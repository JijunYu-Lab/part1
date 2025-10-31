# 加载必要的包
library(readxl)
library(openxlsx)
library(stringr)
library(dplyr)

# 参数修改
# 文件路径参数
input_file <- "C:/Users/Administrator/Desktop/re-part1/BM/amp_BM_1.xlsx"
output_path <- "C:/Users/Administrator/Desktop/re-part1/BM/"
output_prefix <- "amp_BM"  # 输出文件前缀

# 细菌提取参数
target_bacteria <- "Acinetobacter baumannii"  # 目标细菌名称
bacteria_pattern <- "(?i)(?:Acinetobacter[ _-]?baumannii|A[._]?[ _-]?baumannii)[^;&,#]*"# 匹配项正则表达式

# 列名参数
target_organism_col <- "Target_Organism"  # 目标菌列名
other_modifications_col <- "Other_Modifications"  # 其他修饰列名

# MIC提取参数
skip_patterns <- c("MBC", "IC50", "MIC\\s*(50|90|100)")  # 需要跳过的模式
time_units_to_skip <- c("hour", "day", "minute", "sec")  # 需要跳过的时间单位

# 单位规整参数
standard_units <- c("μg/mL", "μM", "mg/mL", "mM", "nM", "M")  # 标准单位列表
highlight_nonstandard <- TRUE  # 是否标记非标准单位（黄色高亮）

# 分组输出参数
save_group_B <- TRUE  # 是否保存B组（无MIC数值）
save_group_A1 <- TRUE  # 是否保存A1组（单个MIC数值）
save_group_A2 <- TRUE  # 是否保存A2组（多个MIC数值）

#############################################################################

# 第一部分：函数定义

# 定义提取菌信息的函数
extract_bacteria <- function(text) {
  if (is.na(text) || text == "") return(NA)
  
  # 使用参数中定义的细菌模式
  pattern <- bacteria_pattern
  
  # 获取所有匹配项
  matches <- str_extract_all(text, pattern)[[1]]
  matches <- matches[!is.na(matches)]
  
  if (length(matches) > 0) {
    # 清理结果
    clean_matches <- trimws(gsub("^[,\\.;\\s]+|[,\\.;\\s]+$", "", matches))
    
    # 去重
    unique_matches <- unique(clean_matches)
    
    return(paste(unique_matches, collapse = "; "))
  } else {
    return(NA)
  }
}

# 定义提取MIC值的函数
extract_mic <- function(text) {
  if (is.na(text) || text == "") return(NA)
  
  # 分割多个条目
  entries <- str_split(text, ";\\s*")[[1]]
  mic_values <- c()
  
  for (entry in entries) {
    # 排除需要跳过的模式
    skip_detected <- FALSE
    for (skip_pattern in skip_patterns) {
      if (str_detect(entry, paste0("(?i)(", skip_pattern, ")\\s*[=:]"))) {
        skip_detected <- TRUE
        break
      }
    }
    if (skip_detected) next
    
    # 匹配MIC
    mic_matches <- str_extract_all(entry, "(?i)(?<!\\w)MIC(?!\\w)\\s*[/]?\\s*[=:]?\\s*[^;)]+")[[1]]
    
    if (length(mic_matches) == 0) next
    
    for (mic_str in mic_matches) {
      # 处理"or"情况，取第一个值
      if (str_detect(mic_str, "\\bor\\b")) {
        parts <- str_split(mic_str, "\\bor\\b")[[1]]
        mic_str <- parts[1]
      }
      
      # 处理MIC/MBC情况，只取MIC部分
      if (str_detect(mic_str, "(?i)MIC\\s*/\\s*MBC")) {
        # 提取MIC值部分（直到下一个斜杠或结束）
        mic_part <- str_extract(mic_str, "(?i)MIC\\s*/\\s*MBC\\s*[=:]?\\s*[^/]+")
        if (!is.na(mic_part)) {
          mic_str <- mic_part
        }
      }
      
      # 移除MIC关键词
      clean_str <- str_replace(mic_str, "(?i)MIC\\s*[/]?\\s*[=:]?\\s*", "")
      clean_str <- str_squish(clean_str)
      
      # 处理带±的值，只取前面的值
      if (str_detect(clean_str, "±")) {
        before_pm <- str_split(clean_str, "±")[[1]][1]
        
        value_unit_match <- str_match(before_pm, "([<>]?=?\\s*)?(\\d+(?:\\.\\d+)?)\\s*([a-zA-Zµμ/]+[a-zA-Zµμ0-9/]*)")
        
        if (!is.na(value_unit_match[1,1])) {
          value <- value_unit_match[1,3]
          unit <- value_unit_match[1,4]
          
          # 若单位是百分比，跳过
          if (!is.na(unit) && str_detect(unit, "%")) {
            next
          }
          
          # 若有单位，组合成结果
          if (!is.na(unit) && unit != "") {
            clean_str <- paste0(value, " ", unit)
          } else {
            # 若未提取到单位，尝试从原始字符串中提取更完整的单位
            unit_from_original <- str_extract(clean_str, "[a-zA-Zµμ/]+[a-zA-Zµμ0-9/]*$")
            if (!is.na(unit_from_original)) {
              clean_str <- paste0(value, " ", unit_from_original)
            } else {
              clean_str <- value
            }
          }
        } else {
          # 备选提取方法：直接提取±前面的数字
          value <- str_extract(before_pm, "\\d+(?:\\.\\d+)?")
          if (!is.na(value)) {
            # 从原始字符串中提取单位
            unit_from_original <- str_extract(clean_str, "[a-zA-Zµμ/]+[a-zA-Zµμ0-9/]*$")
            if (!is.na(unit_from_original) && !str_detect(unit_from_original, "%")) {
              clean_str <- paste0(value, " ", unit_from_original)
            } else {
              clean_str <- value
            }
          } else {
            next
          }
        }
      }
      # 处理范围值
      else if (str_detect(clean_str, "\\d+\\.?\\d*\\s*[-–_]\\s*\\d+\\.?\\d*")) {
        # 提取范围部分和完整单位
        range_match <- str_match(clean_str, "([<>]?=?\\s*)?(\\d+\\.?\\d*\\s*[-–_]\\s*\\d+\\.?\\d*)\\s*([a-zA-Zµμ/]+[a-zA-Zµμ0-9/]*)")
        
        if (!is.na(range_match[1,1])) {
          range_values <- range_match[1,3]
          unit <- range_match[1,4]
          
          # 若单位是百分比，跳过
          if (!is.na(unit) && str_detect(unit, "%")) {
            next
          }
          
          # 分离范围值
          values <- str_split(range_values, "\\s*[-–_]\\s*")[[1]] %>%
            as.numeric()
          
          # 计算中值
          avg_value <- mean(values, na.rm = TRUE)
          
          # 格式化结果
          if (avg_value < 1) {
            formatted_avg <- sprintf("%.3f", avg_value)
          } else if (avg_value < 10) {
            formatted_avg <- sprintf("%.2f", avg_value)
          } else {
            formatted_avg <- sprintf("%.1f", avg_value)
          }
          
          # 移除尾随零
          formatted_avg <- sub("\\.0+$", "", formatted_avg)
          formatted_avg <- sub("(\\.[0-9]*[1-9])0+$", "\\1", formatted_avg)
          
          if (!is.na(unit) && unit != "") {
            clean_str <- paste0(formatted_avg, " ", unit)
          } else {
            # 若未提取到单位，尝试从原始字符串中提取
            unit_from_original <- str_extract(clean_str, "(?<=[-–_]\\s*\\d+(?:\\.\\d+)?\\s*)[a-zA-Zµμ/]+[a-zA-Zµμ0-9/]*")
            if (!is.na(unit_from_original)) {
              clean_str <- paste0(formatted_avg, " ", unit_from_original)
            } else {
              clean_str <- formatted_avg
            }
          }
        }
      }
      # 处理不等号
      else if (str_detect(clean_str, "[<>]")) {
        # 提取数值和完整单位
        value_unit_match <- str_match(clean_str, "[<>]\\s*=?\\s*(\\d+(?:\\.\\d+)?)\\s*([a-zA-Zµμ/]+[a-zA-Zµμ0-9/]*)")
        
        if (!is.na(value_unit_match[1,1])) {
          value <- value_unit_match[1,2]
          unit <- value_unit_match[1,3]
          
          # 若单位是百分比，跳过
          if (!is.na(unit) && str_detect(unit, "%")) {
            next
          }
          
          if (!is.na(unit) && unit != "") {
            clean_str <- paste0(value, " ", unit)
          } else {
            clean_str <- value
          }
        }
      }
      # 一般MIC值提取
      else {
        # 匹配模式：数值 + 空格 + 单位
        value_unit_match <- str_match(clean_str, "([<>]?=?\\s*)?(\\d+(?:\\.\\d+)?)\\s*([a-zA-Zµμ/]+[a-zA-Zµμ0-9/]*)")
        
        if (!is.na(value_unit_match[1,1])) {
          value <- value_unit_match[1,3]
          unit <- value_unit_match[1,4]
          
          # 若单位是百分比，跳过
          if (!is.na(unit) && str_detect(unit, "%")) {
            next
          }
          
          if (!is.na(unit) && unit != "") {
            clean_str <- paste0(value, " ", unit)
          } else {
            # 备选：尝试匹配紧跟在数字后的单位
            value_unit_tight <- str_match(clean_str, "(\\d+(?:\\.\\d+)?)([a-zA-Zµμ/]+[a-zA-Zµμ0-9/]*)")
            if (!is.na(value_unit_tight[1,1])) {
              value <- value_unit_tight[1,2]
              unit <- value_unit_tight[1,3]
              if (!is.na(unit) && !str_detect(unit, "%")) {
                clean_str <- paste0(value, unit)
              } else {
                clean_str <- value
              }
            } else {
              clean_str <- value
            }
          }
        } else {
          # 若无法提取，返回NA
          next
        }
      }
      
      # 移除部分残留的无关文本
      # 移除包含时间单位的条目
      time_detected <- FALSE
      for (time_unit in time_units_to_skip) {
        if (str_detect(clean_str, time_unit)) {
          time_detected <- TRUE
          break
        }
      }
      if (time_detected) next
      
      # 移除等号、不等号
      clean_str <- str_replace_all(clean_str, "[<>]=?", "")
      clean_str <- str_squish(clean_str)
      
      # 检查结果
      # 必须包含数字和合理单位
      if (str_detect(clean_str, "\\d") && 
          !str_detect(clean_str, "%") && 
          nchar(clean_str) > 0 &&
          !str_detect(clean_str, "±")) {  # 确保没有残留的±符号
        mic_values <- c(mic_values, clean_str)
      }
    }
  }
  
  if (length(mic_values) > 0) return(mic_values)
  return(NA)
}

# 定义单位规整函数
standardize_units <- function(text) {
  if (is.na(text) || text == "") {
    return(text)
  }
  
  # 分离数值和单位
  matches <- str_match(text, "^([-+]?[0-9]*\\.?[0-9]+)(\\s*)(.*)$")
  
  if (is.na(matches[1,1])) {
    return(text)  
  }
  
  number <- matches[1,2]
  unit <- matches[1,4]
  
  # 规整单位（不改变数值）
  standardized_unit <- case_when(
    # 将所有质量浓度单位统一为μg/mL
    str_detect(unit, regex("mg/l|mg/L", ignore_case = TRUE)) ~ "μg/mL",
    str_detect(unit, regex("microg/ml|ug/ml|μg/ml", ignore_case = TRUE)) ~ "μg/mL",
    
    # 将所有摩尔浓度单位统一为μM
    str_detect(unit, regex("microm|micromol/l|μmol/l", ignore_case = TRUE)) ~ "μM",
    str_detect(unit, regex("um|uM", ignore_case = TRUE)) ~ "μM",
    
    # 标准单位保持不变
    str_detect(unit, regex("μg/mL", ignore_case = TRUE)) ~ "μg/mL",
    str_detect(unit, regex("μM", ignore_case = TRUE)) ~ "μM",
    str_detect(unit, regex("mg/mL", ignore_case = TRUE)) ~ "mg/mL",
    str_detect(unit, regex("mM", ignore_case = TRUE)) ~ "mM",
    str_detect(unit, regex("nM", ignore_case = TRUE)) ~ "nM",
    str_detect(unit, regex("M", ignore_case = TRUE)) ~ "M",
    # 若无匹配则保持原单位
    TRUE ~ unit
  )
  
  # 返回规整后的数值+单位
  return(paste(number, standardized_unit))
}

# 定义执行单位规整的函数
process_unit_standardization <- function(data, output_filename) {
  # 使用参数中定义的标准单位列表
  standard_units <- standard_units
  
  # 找到目标列
  mic_columns <- which(names(data) == "MIC_1"):ncol(data)
  
  # 创建单位统计表
  unit_stats <- list()
  
  # 创建用于标记非标准单位的向量
  rows_with_nonstandard_units <- c()
  
  # 应用
  for (col in mic_columns) {
    # 提取原始单位
    original_units <- sapply(data[[col]], function(text) {
      if (is.na(text) || text == "") {
        return(NA)
      }
      matches <- str_match(text, "^([-+]?[0-9]*\\.?[0-9]+)(\\s*)(.*)$")
      if (is.na(matches[1,1])) {
        return(NA)
      }
      return(matches[1,4])
    })
    
    # 统计
    original_counts <- table(original_units, useNA = "ifany")
    
    # 应用单位规整
    data[[col]] <- sapply(data[[col]], standardize_units)
    
    # 提取规整后的单位
    standardized_units <- sapply(data[[col]], function(text) {
      if (is.na(text) || text == "") {
        return(NA)
      }
      matches <- str_match(text, "^([-+]?[0-9]*\\.?[0-9]+)(\\s*)(.*)$")
      if (is.na(matches[1,1])) {
        return(NA)
      }
      return(matches[1,4])
    })
    
    # 统计规整后的单位
    standardized_counts <- table(standardized_units, useNA = "ifany")
    
    # 检查是否有非标准单位
    nonstandard_rows <- which(!standardized_units %in% standard_units & !is.na(standardized_units))
    if (length(nonstandard_rows) > 0) {
      rows_with_nonstandard_units <- unique(c(rows_with_nonstandard_units, nonstandard_rows))
    }
    
    # 存储统计结果
    unit_stats[[names(data)[col]]] <- list(
      original = original_counts,
      standardized = standardized_counts
    )
  }
  
  # 创建新的工作簿
  wb <- createWorkbook()
  
  # 添加工作表
  addWorksheet(wb, "Data")
  
  # 写入数据
  writeData(wb, "Data", data)
  
  # 标记非标准单位（如果启用）
  if (highlight_nonstandard && length(rows_with_nonstandard_units) > 0) {
    # 创建黄色填充样式
    yellow_fill <- createStyle(fgFill = "#FFFF00")
    
    # 行号加一（因为第一行是标题）
    addStyle(wb, "Data", yellow_fill, 
             rows = rows_with_nonstandard_units + 1, 
             cols = 1:ncol(data), 
             gridExpand = TRUE)
  }
  
  # 保存工作簿
  saveWorkbook(wb, paste0(output_path, output_filename), overwrite = TRUE)
  
  # 打印单位统计信息
  cat("单位规整完成！处理了", length(mic_columns), "列数据。\n\n")
  
  # 打印详细的单位统计
  for (col_name in names(unit_stats)) {
    cat("列名:", col_name, "\n")
    
    cat("  规整前单位统计:\n")
    if (length(unit_stats[[col_name]]$original) > 0) {
      for (i in 1:length(unit_stats[[col_name]]$original)) {
        unit <- names(unit_stats[[col_name]]$original)[i]
        count <- unit_stats[[col_name]]$original[i]
        cat("    ", unit, ":", count, "\n")
      }
    } else {
      cat("    无单位数据\n")
    }
    
    cat("  规整后单位统计:\n")
    if (length(unit_stats[[col_name]]$standardized) > 0) {
      for (i in 1:length(unit_stats[[col_name]]$standardized)) {
        unit <- names(unit_stats[[col_name]]$standardized)[i]
        count <- unit_stats[[col_name]]$standardized[i]
        cat("    ", unit, ":", count, "\n")
      }
    } else {
      cat("    无单位数据\n")
    }
    
    cat("\n")
  }
  
  # 打印非标准单位行统计
  cat("非标准单位检查结果:\n")
  if (length(rows_with_nonstandard_units) > 0) {
    cat("  发现", length(rows_with_nonstandard_units), "行包含非标准单位")
    if (highlight_nonstandard) {
      cat("，已标记为黄色")
    }
    cat("\n")
    cat("  受影响的行号:", paste(sort(rows_with_nonstandard_units), collapse = ", "), "\n")
  } else {
    cat("  所有行均使用标准单位，无需标记\n")
  }
  
  cat("\n文件已保存至:", paste0(output_path, output_filename), "\n")
  cat("========================================\n\n")
  
  return(data)
}

# 第二部分：主处理流程

# 读取数据
data <- read_excel(input_file)

# 检查必要的列是否存在
if (!target_organism_col %in% names(data)) {
  stop("错误：找不到目标菌列 '", target_organism_col, "'。请检查列名参数。")
}

# 应用提取函数到目标列
data$Extracted <- sapply(data[[target_organism_col]], extract_bacteria)

# 重新排列列顺序（将新列放在目标列后）
target_idx <- which(names(data) == target_organism_col)
if (ncol(data) > target_idx) {
  new_order <- c(
    1:target_idx,
    ncol(data),
    (target_idx + 1):(ncol(data) - 1)
  )
} else {
  new_order <- c(1:target_idx, ncol(data))
}
data <- data[, new_order]

# 处理提取列
mic_results <- lapply(data$Extracted, extract_mic)

max_cols <- max(sapply(mic_results, function(x) ifelse(is.na(x[1]), 0, length(x))))

# 创建新的MIC列
if (other_modifications_col %in% colnames(data)) {
  start_col <- which(colnames(data) == other_modifications_col) + 1
} else {
  start_col <- which(colnames(data) == "Extracted") + 1
}

new_cols <- paste0("MIC_", seq_len(max_cols))

# 添加新列到数据框
for (i in seq_along(new_cols)) {
  data[[new_cols[i]]] <- NA
}

# 填充MIC值
for (i in seq_len(nrow(data))) {
  if (!is.na(mic_results[[i]][1])) {
    for (j in seq_along(mic_results[[i]])) {
      data[i, start_col + j - 1] <- mic_results[[i]][j]
    }
  } else {
    data[i, start_col] <- "NA"
  }
}

# 第三部分：分组
# 索引
mic_cols <- which(names(data) == "MIC_1"):ncol(data)

# 初始化三个分组
group_A1 <- data.frame()
group_A2 <- data.frame()
group_B <- data.frame()

# 遍历每一行数据
for(i in 1:nrow(data)) {
  # 提取当前行的MIC相关列
  mic_values <- data[i, mic_cols]
  
  # 计算包含数字的列数
  numeric_count <- 0
  
  for(j in 1:length(mic_values)) {
    cell_value <- as.character(mic_values[[j]])
    # 检查单元格是否包含数字
    if(!is.na(cell_value) & str_detect(cell_value, "[0-9]")) {
      numeric_count <- numeric_count + 1
    }
  }
  
  # 根据数字计数分类
  if(numeric_count == 0) {
    group_B <- rbind(group_B, data[i, ])
  } else if(numeric_count == 1) {
    group_A1 <- rbind(group_A1, data[i, ])
  } else {
    group_A2 <- rbind(group_A2, data[i, ])
  }
}

# 第四部分：汇总并保存

cat("分组完成！开始单位规整处理...\n")
cat("目标细菌:", target_bacteria, "\n")
cat("原始数据行数：", nrow(data), "\n")
cat("A1组（只有一个MIC数值）：", nrow(group_A1), "行\n")
cat("A2组（有多个MIC数值）：", nrow(group_A2), "行\n")
cat("B组（无MIC数值）：", nrow(group_B), "行\n")
cat("========================================\n\n")

# 保存B组（无需单位规整）
if(save_group_B && nrow(group_B) > 0) {
  b_output_file <- paste0(output_prefix, "_B.xlsx")
  write.xlsx(group_B, paste0(output_path, b_output_file))
  cat("B组已保存至：", paste0(output_path, b_output_file), "\n")
} else if (save_group_B && nrow(group_B) == 0) {
  cat("警告：B组为空，未生成B组文件\n")
}

# 处理A1组（单位规整）
if(save_group_A1 && nrow(group_A1) > 0) {
  cat("开始处理A1组...\n")
  a1_output_file <- paste0(output_prefix, "_A1_unit.xlsx")
  process_unit_standardization(group_A1, a1_output_file)
} else if (save_group_A1 && nrow(group_A1) == 0) {
  cat("警告：A1组为空，未生成A1组文件\n")
}

# 处理A2组（单位规整）
if(save_group_A2 && nrow(group_A2) > 0) {
  cat("开始处理A2组...\n")
  a2_output_file <- paste0(output_prefix, "_A2_unit.xlsx")
  process_unit_standardization(group_A2, a2_output_file)
} else if (save_group_A2 && nrow(group_A2) == 0) {
  cat("警告：A2组为空，未生成A2组文件\n")
}

cat("所有处理完成！\n")