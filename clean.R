# 加载必要的包
if (!require("readxl")) install.packages("readxl")
if (!require("writexl")) install.packages("writexl")
if (!require("dplyr")) install.packages("dplyr")
if (!require("stringr")) install.packages("stringr")

library(readxl)
library(writexl)
library(dplyr)
library(stringr)

# 设置文件路径
data_path <- "C:/Users/Administrator/Desktop/re-part1/data/"

# 定义数据清洗函数（氨基酸清洗）
clean_peptide_data <- function(input_file, output_file) {
  # 读取数据
  data <- read_excel(paste0(data_path, input_file))
  
  # 记录原始序列数量
  total_original <- nrow(data)
  
  # 检查非天然氨基酸 (B, J, O, U, X, Z)
  non_standard_aa <- data %>% 
    filter(grepl("[BJOUXZ]", Sequence, ignore.case = FALSE))
  count_non_standard <- nrow(non_standard_aa)
  
  # 检查小写字母 (D型氨基酸)
  d_amino_acids <- data %>% 
    filter(grepl("[a-z]", Sequence))
  count_d_amino <- nrow(d_amino_acids)
  
  # 数据清洗：移除包含非天然氨基酸或小写字母的序列
  clean_data <- data %>% 
    filter(!grepl("[BJOUXZ]", Sequence, ignore.case = FALSE)) %>% 
    filter(!grepl("[a-z]", Sequence))
  
  # 记录清洗后的序列数量
  total_clean <- nrow(clean_data)
  
  # 输出日志信息
  cat("=== 氨基酸清洗处理日志 ===\n")
  cat("输入文件:", input_file, "\n")
  cat("原始序列数量:", total_original, "\n")
  cat("包含非天然氨基酸的序列:", count_non_standard, "\n")
  cat("包含D型氨基酸的序列:", count_d_amino, "\n")
  cat("清洗后序列数量:", total_clean, "\n")
  cat("输出文件:", output_file, "\n")
  cat("==========================\n\n")
  
  # 返回清洗后的数据和统计信息
  return(list(
    data = clean_data,
    stats = list(
      original = total_original,
      non_standard = count_non_standard,
      d_amino = count_d_amino,
      cleaned = total_clean
    )
  ))
}

# 定义形状过滤函数
filter_peptide_shape <- function(data, input_file, output_file) {
  # 查找肽形状相关的列名
  shape_columns <- c(
    "Linear/Cyclic/Branched",
    "Structure_type",
    "Topology",
    "Shape",
    "Peptide_shape",
    "Structure"
  )
  
  # 在数据中查找实际存在的形状相关列
  existing_shape_cols <- shape_columns[shape_columns %in% names(data)]
  
  # 如果没有找到相关列，直接返回原数据
  if (length(existing_shape_cols) == 0) {
    cat("在文件", input_file, "中未找到肽形状相关列，直接保存文件\n")
    return(list(
      data = data,
      stats = list(
        original = nrow(data),
        filtered = nrow(data),
        removed = 0
      )
    ))
  }
  
  # 使用找到的第一个形状相关列
  shape_col <- existing_shape_cols[1]
  cat("在文件", input_file, "中找到肽形状列:", shape_col, "\n")
  
  # 定义需要删除的非线性肽标识符（不区分大小写）
  nonlinear_identifiers <- c(
    "branched", "cyclic", "dimer", "trimer", "tetramer", "stapled", 
    "multimer", "polymer", "branched peptide", "cyclic peptide", 
    "dimer peptide", "trimer peptide", "tetramer peptide", "stapled peptide"
  )
  
  # 过滤数据：删除明确标注为非线性的肽序列
  # 保留线性肽、空白值、NA值以及其他未明确标注为非线性的值
  filtered_data <- data %>%
    dplyr::mutate(
      shape_value = as.character(.data[[shape_col]]),
      shape_value = ifelse(is.na(shape_value), "", trimws(shape_value)),
      shape_value_lower = tolower(shape_value)
    ) %>%
    dplyr::filter(
      # 保留不包含任何非线性标识符的行
      !str_detect(shape_value_lower, paste(nonlinear_identifiers, collapse = "|")) |
        shape_value == "" |
        is.na(.data[[shape_col]])
    ) %>%
    dplyr::select(-shape_value, -shape_value_lower)
  
  # 输出处理信息
  original_rows <- nrow(data)
  filtered_rows <- nrow(filtered_data)
  removed_rows <- original_rows - filtered_rows
  
  cat("=== 形状过滤处理日志 ===\n")
  cat("输入文件:", input_file, "\n")
  cat("原始行数:", original_rows, "\n")
  cat("过滤后行数:", filtered_rows, "\n")
  cat("删除行数:", removed_rows, "\n")
  cat("输出文件:", output_file, "\n")
  cat("========================\n\n")
  
  # 返回处理后的数据和统计信息
  return(list(
    data = filtered_data,
    stats = list(
      original = original_rows,
      filtered = filtered_rows,
      removed = removed_rows
    )
  ))
}

# 定义长度过滤函数（来自代码二）
filter_peptide_length <- function(data, input_file, output_file) {
  # 计算序列长度并过滤掉长度>=33的序列
  filtered_data <- data %>%
    mutate(Sequence_Length = nchar(Sequence)) %>%
    filter(Sequence_Length < 33) %>%
    select(-Sequence_Length)  # 移除临时添加的长度列
  
  # 输出处理信息
  original_rows <- nrow(data)
  filtered_rows <- nrow(filtered_data)
  removed_rows <- original_rows - filtered_rows
  
  cat("=== 长度过滤处理日志 ===\n")
  cat("输入文件:", input_file, "\n")
  cat("原始行数:", original_rows, "\n")
  cat("过滤后行数:", filtered_rows, "\n")
  cat("删除行数:", removed_rows, "\n")
  cat("输出文件:", output_file, "\n")
  cat("========================\n\n")
  
  # 保存处理后的文件
  writexl::write_xlsx(filtered_data, paste0(data_path, output_file))
  
  # 返回处理后的数据和统计信息
  return(list(
    data = filtered_data,
    stats = list(
      original = original_rows,
      filtered = filtered_rows,
      removed = removed_rows
    )
  ))
}

# 主处理流程
process_peptide_pipeline <- function(input_file, final_output_file) {
  cat("开始处理文件:", input_file, "\n")
  
  # 第一步：氨基酸清洗
  step1_result <- clean_peptide_data(input_file, "temp_amino_clean.xlsx")
  
  # 第二步：形状过滤
  step2_result <- filter_peptide_shape(step1_result$data, "temp_amino_clean.xlsx", "temp_shape_clean.xlsx")
  
  # 第三步：长度过滤
  step3_result <- filter_peptide_length(step2_result$data, "temp_shape_clean.xlsx", final_output_file)
  
  cat("文件处理完成:", input_file, "->", final_output_file, "\n\n")
  
  # 返回总体统计信息
  return(list(
    amino_cleaning = step1_result$stats,
    shape_filtering = step2_result$stats,
    length_filtering = step3_result$stats
  ))
}

# 执行主处理流程
cat("正在启动肽数据处理流程...\n\n")

# 处理第一个数据集：01_dbAMP3_pepinfo.xlsx -> 05_dbAMP_clean.xlsx
cat("正在处理 dbAMP 数据集...\n")
stats_dbAMP <- process_peptide_pipeline("01_dbAMP3_pepinfo.xlsx", "05_dbAMP_clean.xlsx")

# 处理第二个数据集：04_dramp_all.xlsx -> 06_DRAMP_clean.xlsx  
cat("正在处理 DRAMP 数据集...\n")
stats_DRAMP <- process_peptide_pipeline("04_dramp_all.xlsx", "06_DRAMP_clean.xlsx")

# 输出总体统计
cat("=== 总体处理摘要 ===\n")
cat("dbAMP 数据集:\n")
cat("  氨基酸清洗: 保留", stats_dbAMP$amino_cleaning$cleaned, "/", stats_dbAMP$amino_cleaning$original, "条序列\n")
cat("  形状过滤: 保留", stats_dbAMP$shape_filtering$filtered, "/", stats_dbAMP$shape_filtering$original, "条序列\n")
cat("  长度过滤: 保留", stats_dbAMP$length_filtering$filtered, "/", stats_dbAMP$length_filtering$original, "条序列\n")

cat("DRAMP 数据集:\n")
cat("  氨基酸清洗: 保留", stats_DRAMP$amino_cleaning$cleaned, "/", stats_DRAMP$amino_cleaning$original, "条序列\n")
cat("  形状过滤: 保留", stats_DRAMP$shape_filtering$filtered, "/", stats_DRAMP$shape_filtering$original, "条序列\n")
cat("  长度过滤: 保留", stats_DRAMP$length_filtering$filtered, "/", stats_DRAMP$length_filtering$original, "条序列\n")

total_original <- stats_dbAMP$amino_cleaning$original + stats_DRAMP$amino_cleaning$original
total_final <- stats_dbAMP$length_filtering$filtered + stats_DRAMP$length_filtering$filtered

cat("总保留率:", round(total_final / total_original * 100, 2), "%\n")
cat("===================\n")

cat("所有文件处理完成！\n")