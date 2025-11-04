# 加载必要的包
library(openxlsx)
library(Peptides)
library(stringr)
library(dplyr)
library(readr)

# 设置工作路径
setwd("C:/Users/Administrator/Desktop/re-part1")

# 从代码一中复制的分子量计算函数（保持不变）
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

# 处理单个靶标菌的函数（修正MIC优异集获取方式）
process_bacteria <- function(bacteria_code) {
  cat("正在处理", bacteria_code, "...\n")
  
  # 1. 读取MIC预测文件和对应的输入文件
  mic_pred_file <- paste0("251101-repart1/B_", bacteria_code, "_predictions.csv")
  input_file <- paste0("251101-repart1/amp_", bacteria_code, "_B.csv")
  
  # 使用read.csv并指定编码
  mic_pred <- read.csv(mic_pred_file, header = FALSE, skip = 1, col.names = "MIC")
  input_data <- read.csv(input_file)
  
  # 合并数据
  combined_data <- cbind(input_data, MIC = mic_pred$MIC)
  
  # 2. 计算分子量并转换为μM单位
  molecular_weights <- numeric(nrow(combined_data))
  
  for (i in 1:nrow(combined_data)) {
    molecular_weights[i] <- calculate_peptide_mw(
      as.character(combined_data$Sequence[i]),
      as.character(combined_data$N_terminal[i]),
      as.character(combined_data$C_terminal[i])
    )
  }
  
  # 转换为μM单位: MIC_uM = (MIC_μg/mL * 1000) / 分子量
  combined_data$MIC_uM <- (combined_data$MIC * 1000) / molecular_weights
  combined_data <- combined_data[order(combined_data$MIC_uM, decreasing = FALSE), ]
  
  # 3. 获取MIC优异集 - 分别取两个来源的前10%然后取并集
  # 读取已知MIC排序文件
  known_mic_file <- paste0("A_rank/", bacteria_code, "_A_rank.xlsx")
  known_mic_data <- read.xlsx(known_mic_file)
  
  # 计算前10%的界限
  n_known <- nrow(known_mic_data)
  cutoff_index_known <- ceiling(n_known * 0.1)
  known_cutoff <- known_mic_data$MIC[cutoff_index_known]
  
  n_pred <- nrow(combined_data)
  cutoff_index_pred <- ceiling(n_pred * 0.1)
  pred_cutoff <- combined_data$MIC_uM[cutoff_index_pred]
  
  cat(bacteria_code, "已知MIC前10%界限:", known_cutoff, "μM\n")
  cat(bacteria_code, "预测MIC前10%界限:", pred_cutoff, "μM\n")
  
  # 获取已知数据的MIC优异集
  known_excellent <- known_mic_data[1:cutoff_index_known, ]
  
  # 获取预测数据的MIC优异集
  pred_excellent <- combined_data[1:cutoff_index_pred, ]
  
  # 取两个优异集的并集
  mic_excellent <- bind_rows(known_excellent, pred_excellent)
  # 去除重复序列
  mic_excellent <- mic_excellent[!duplicated(mic_excellent$Sequence), ]
  
  cat(bacteria_code, "MIC优异集大小:", nrow(mic_excellent), "条序列\n")
  
  # 4. 获取毒性优异集 - 使用read.csv解决编码问题
  tox_file <- paste0("Tox/", bacteria_code, "_Toxfinal_output.csv")
  
  # 尝试不同编码方式读取毒性文件
  tox_data <- tryCatch({
    read.csv(tox_file, fileEncoding = "UTF-8")
  }, error = function(e) {
    cat("UTF-8编码读取失败，尝试GBK编码...\n")
    read.csv(tox_file, fileEncoding = "GBK")
  })
  
  # 检查列名，适应可能的列名变化
  hybrid_score_col <- NULL
  sequence_col <- NULL
  
  # 查找包含"Hybrid"或"Score"的列
  for (col in names(tox_data)) {
    if (grepl("Hybrid|Score", col, ignore.case = TRUE)) {
      hybrid_score_col <- col
    }
    if (grepl("Sequence", col, ignore.case = TRUE)) {
      sequence_col <- col
    }
  }
  
  if (is.null(hybrid_score_col) || is.null(sequence_col)) {
    cat("警告：在毒性文件中找不到所需的列，使用默认列名\n")
    cat("找到的列：", paste(names(tox_data), collapse = ", "), "\n")
    # 如果找不到，使用前几列作为默认
    hybrid_score_col <- names(tox_data)[4]  # 通常第4列是Hybrid Score
    sequence_col <- names(tox_data)[2]      # 通常第2列是Sequence
  }
  
  # 获取毒性优异集
  tox_excellent <- tox_data[tox_data[[hybrid_score_col]] == 0, ]
  
  # 5. 取交集
  top_data <- mic_excellent[mic_excellent$Sequence %in% tox_excellent[[sequence_col]], ]
  
  # 输出结果
  output_file <- paste0("Top/", bacteria_code, "_top.xlsx")
  write.xlsx(top_data, output_file)
  
  cat(bacteria_code, "处理完成，找到", nrow(top_data), "条Top序列\n")
  
  return(top_data)
}

# 主处理流程
# 创建输出目录
if (!dir.exists("Top")) {
  dir.create("Top")
}

# 处理所有靶标菌
bacteria_codes <- c("TL", "BM", "DC", "FK")
all_top_data <- list()

for (bacteria in bacteria_codes) {
  result <- tryCatch({
    top_data <- process_bacteria(bacteria)
    all_top_data[[bacteria]] <- top_data
    cat(bacteria, "处理成功\n")
  }, error = function(e) {
    cat(bacteria, "处理失败:", e$message, "\n")
    NULL
  })
}

# 合并所有Top集（取并集）
if (length(all_top_data) > 0) {
  combined_top <- bind_rows(all_top_data)
  # 去除重复序列
  combined_top <- combined_top[!duplicated(combined_top$Sequence), ]
  
  # 输出最终合并文件
  write.xlsx(combined_top, "Top/TOP_down.xlsx")
  
  cat("所有处理完成！\n")
  cat("各靶标菌Top序列数量:\n")
  for (bacteria in bacteria_codes) {
    if (!is.null(all_top_data[[bacteria]])) {
      cat(bacteria, ":", nrow(all_top_data[[bacteria]]), "条\n")
    } else {
      cat(bacteria, ": 0条（处理失败）\n")
    }
  }
  cat("最终合并去重后:", nrow(combined_top), "条序列\n")
  cat("结果保存在 C:/Users/Administrator/Desktop/re-part1/Top/ 目录下\n")
} else {
  cat("所有处理都失败了，请检查文件路径和格式\n")
}