# 加载必要的包
library(openxlsx)
library(Peptides)
library(stringr)

# 设置文件路径
input_file1 <- "C:/Users/Administrator/Desktop/re-part1/TL/TL_A1_unit.xlsx"
input_file2 <- "C:/Users/Administrator/Desktop/re-part1/TL/TL_A2_unit.xlsx"
output_file <- "C:/Users/Administrator/Desktop/re-part1/TL/TL_A_rank.xlsx"

# 扩展的N末端修饰分子量计算函数
get_n_terminal_mod_mw <- function(modification) {
  if (is.na(modification) || modification == "" || modification == "Free") {
    return(0)
  }
  
  mod_lower <- tolower(modification)
  
  # 乙酰化及相关修饰
  if (grepl("acetyl", mod_lower) || mod_lower %in% c("acetylation", "acetylization")) {
    return(42.01)
  }
  
  # 脂肪酸修饰
  if (grepl("c6|hexanoyl", mod_lower)) return(100.12)
  if (grepl("c8|octanoyl|octanoic", mod_lower)) return(128.17)
  if (grepl("c10|decanoyl", mod_lower)) return(156.27)
  if (grepl("c12|dodecanoyl|dodecanoic|lauric", mod_lower)) return(184.32)
  if (grepl("c14|myristoyl|myristic", mod_lower)) return(212.37)
  if (grepl("c16|palmitoyl|palmitic|pal,ch3\\(ch2\\)14co", mod_lower)) return(240.43)
  if (grepl("c18|stearoyl|stearic|stearyl|ch3\\(ch2\\)16co", mod_lower)) return(268.48)
  
  # 生物素化
  if (grepl("biotin", mod_lower)) return(226.29)
  
  # 荧光标记
  if (grepl("fitc|fluorescein|5\\(6\\)-carboxyfluorescein|carboxyfluorescein", mod_lower)) return(389.38)
  if (grepl("tamra", mod_lower)) return(430.47)
  if (grepl("rhodamine", mod_lower)) return(479.02)
  
  # 聚乙二醇化（未知聚合度，默认为5）
  if (grepl("mpeg|peg", mod_lower)) return(220.26)
  
  # 其他常见修饰
  if (grepl("boc|tert-butyloxycarbonyl", mod_lower)) return(100.12)
  if (grepl("pivaloyl", mod_lower)) return(100.12)
  if (grepl("maleimide", mod_lower)) return(98.06)
  if (grepl("phenylacetyl", mod_lower)) return(120.15)
  
  # 无法识别的修饰视为free
  return(0)
}

# 扩展的C末端修饰分子量计算函数
get_c_terminal_mod_mw <- function(modification) {
  if (is.na(modification) || modification == "" || modification == "Free") {
    return(0)
  }
  
  mod_lower <- tolower(modification)
  
  # 酰胺化
  if (grepl("amidation|amidated|α-amidation", mod_lower)) {
    return(-0.98)  # 酰胺化减少的质量
  }
  
  # 脂肪酸修饰（与N端相同）
  if (grepl("c6|hexanoyl", mod_lower)) return(100.12)
  if (grepl("c8|octanoyl|octanoic", mod_lower)) return(128.17)
  if (grepl("c10|decanoyl", mod_lower)) return(156.27)
  if (grepl("c12|dodecanoyl|dodecanoic|lauric", mod_lower)) return(184.32)
  if (grepl("c14|myristoyl|myristic", mod_lower)) return(212.37)
  if (grepl("c16|palmitoyl|palmitic", mod_lower)) return(240.43)
  if (grepl("c18|stearoyl|stearic|stearyl", mod_lower)) return(268.48)
  
  # 其他C端修饰
  if (grepl("boc|tert-butyloxycarbonyl", mod_lower)) return(100.12)
  if (grepl("methyl amidation", mod_lower)) return(13.02)
  if (grepl("cysteamid", mod_lower)) return(75.14)
  
  # 无法识别的修饰视为free
  return(0)
}

# 检查修饰是否可识别
is_modification_recognized <- function(modification, terminal_type) {
  if (is.na(modification) || modification == "" || modification == "Free") {
    return(TRUE)  # Free是可识别的
  }
  
  mod_lower <- tolower(modification)
  
  if (terminal_type == "N") {
    recognized_patterns <- c(
      "acetyl", "c6", "c8", "c10", "c12", "c14", "c16", "c18", 
      "biotin", "fitc", "fluorescein", "tamra", "rhodamine", 
      "mpeg", "peg", "boc", "pivaloyl", "maleimide", "phenylacetyl",
      "hexanoyl", "octanoyl", "decanoyl", "dodecanoyl", "myristoyl", 
      "palmitoyl", "stearoyl", "stearyl", "lauric"
    )
  } else { # C terminal
    recognized_patterns <- c(
      "amidation", "amidated", "c6", "c8", "c10", "c12", "c14", "c16", "c18",
      "boc", "methyl amidation", "cysteamid",
      "hexanoyl", "octanoyl", "decanoyl", "dodecanoyl", "myristoyl", 
      "palmitoyl", "stearoyl", "stearyl", "lauric"
    )
  }
  
  for (pattern in recognized_patterns) {
    if (grepl(pattern, mod_lower)) {
      return(TRUE)
    }
  }
  
  return(FALSE)
}

# 函数计算肽的分子量（考虑修饰）
calculate_peptide_mw <- function(sequence, n_terminal, c_terminal) {
  # 基础肽段分子量
  base_mw <- mw(seq = sequence, monoisotopic = FALSE)
  
  # N末端修饰
  n_mod_mw <- get_n_terminal_mod_mw(n_terminal)
  base_mw <- base_mw + n_mod_mw
  
  # C末端修饰
  c_mod_mw <- get_c_terminal_mod_mw(c_terminal)
  base_mw <- base_mw + c_mod_mw
  
  return(base_mw)
}

# 单位转换函数 - 处理包含单位的字符串
convert_to_uM <- function(cell_value, mw) {
  if (is.na(cell_value) || cell_value == "") {
    return(NA)
  }
  
  # 提取数值和单位
  # 匹配模式：数字（可能包含小数点）后跟单位
  matches <- str_match(cell_value, "([0-9.]+)\\s*([a-zA-Zμ/]+)")
  
  if (any(is.na(matches))) {
    return(NA)  # 无法解析的格式
  }
  
  value <- as.numeric(matches[1, 2])
  unit <- matches[1, 3]
  
  # 单位转换
  if (grepl("μg/mL|ug/mL", unit)) {
    return((value * 1000) / mw)      # μg/mL to μM
  } else if (grepl("mg/mL", unit)) {
    return((value * 1000000) / mw)   # mg/mL to μM
  } else if (grepl("μM|uM", unit)) {
    return(value)                    # 已经是μM，无需转换
  } else if (grepl("mM", unit)) {
    return(value * 1000)             # mM to μM
  } else if (grepl("nM", unit)) {
    return(value / 1000)             # nM to μM
  } else if (grepl("^M$", unit)) {   # 单独的"M"表示摩尔
    return(value * 1000000)          # M to μM
  } else {
    return(NA)  # 未知单位
  }
}

# 计算几何平均值函数
calculate_geometric_mean <- function(row, mic_cols) {
  values <- as.numeric(row[mic_cols])
  values <- values[!is.na(values) & values > 0]  # 移除NA和零值
  if (length(values) == 0) {
    return(NA)
  }
  exp(mean(log(values)))
}

# 处理单个文件的函数
process_file <- function(file_path) {
  # 读取Excel文件
  data <- read.xlsx(file_path)
  
  # 检查修饰列是否存在
  if (!"N_terminal" %in% names(data)) {
    data$N_terminal <- "Free"
  }
  if (!"C_terminal" %in% names(data)) {
    data$C_terminal <- "Free"
  }
  
  # 为每行计算分子量并检查修饰识别状态
  molecular_weights <- numeric(nrow(data))
  unrecognized_mods <- logical(nrow(data))
  
  for (i in 1:nrow(data)) {
    n_term <- as.character(data$N_terminal[i])
    c_term <- as.character(data$C_terminal[i])
    
    # 检查修饰是否可识别
    n_recognized <- is_modification_recognized(n_term, "N")
    c_recognized <- is_modification_recognized(c_term, "C")
    
    unrecognized_mods[i] <- !n_recognized || !c_recognized
    
    # 计算分子量
    molecular_weights[i] <- calculate_peptide_mw(
      as.character(data$Sequence[i]), 
      n_term, 
      c_term
    )
  }
  
  # 识别MIC相关的列
  mic_columns <- grep("^MIC_\\d+", names(data), value = TRUE)
  
  # 对每个MIC列进行单位转换
  for (col in mic_columns) {
    # 应用单位转换
    converted_values <- numeric(length = nrow(data))
    
    for (i in 1:nrow(data)) {
      cell_value <- as.character(data[i, col])
      converted_values[i] <- convert_to_uM(cell_value, molecular_weights[i])
    }
    
    data[[col]] <- converted_values
  }
  
  # 添加MIC列（在MIC_1之前）
  mic_1_index <- which(names(data) == "MIC_1")
  mic_values <- apply(data, 1, calculate_geometric_mean, mic_cols = mic_columns)
  data <- cbind(data[, 1:(mic_1_index-1), drop = FALSE], 
                MIC = mic_values, 
                data[, mic_1_index:ncol(data), drop = FALSE])
  
  # 添加警示列
  data$Warning <- ifelse(unrecognized_mods, 
                         "Unrecognized modification - MIC may have error", 
                         "")
  
  return(list(data = data, unrecognized_count = sum(unrecognized_mods)))
}

# 处理两个文件
result1 <- process_file(input_file1)
result2 <- process_file(input_file2)

data1 <- result1$data
data2 <- result2$data

# 合并两个数据框
combined_data <- rbind(data1, data2)

# 按MIC值升序排列（MIC值越小越靠前）
combined_data <- combined_data[order(combined_data$MIC, decreasing = FALSE), ]

# 创建工作簿并设置样式
wb <- createWorkbook()
addWorksheet(wb, "Results")

# 写入数据
writeData(wb, "Results", combined_data, rowNames = FALSE)

# 修复：设置警示行的黄色背景
warning_rows <- which(combined_data$Warning != "") + 1  # +1 for header row
if (length(warning_rows) > 0) {
  warning_style <- createStyle(fgFill = "#FFFF00")  
  
  # 对每个警示行应用样式到所有列
  for (row in warning_rows) {
    addStyle(wb, "Results", warning_style, rows = row, cols = 1:ncol(combined_data))
  }
}

# 保存结果
saveWorkbook(wb, output_file, overwrite = TRUE)

cat("处理完成！结果已保存至:", output_file, "\n")
cat("无法识别的修饰行数:", result1$unrecognized_count + result2$unrecognized_count, "\n")
if ((result1$unrecognized_count + result2$unrecognized_count) > 0) {
  cat("警示行已在Excel中标记为黄色背景\n")
}