# 加载必要的库
library(readxl)
library(stringr)
library(openxlsx)

# 定义精准提取肺炎克雷伯菌信息的函数
extract_klebsiella <- function(text) {
  # 若输入为空则返回空
  if (is.na(text) || text == "") return(NA)
  
  # 统一匹配模式 - 覆盖所有常见表达方式
  pattern <- "\\b(?:Klebsiella[ -]?pneumoniae|K\\.?[ _]?pneumoniae)\\b(?:[^;&,#]|\\([^)]*\\))*"
  
  # 获取所有匹配项
  matches <- str_extract_all(text, pattern)[[1]]
  matches <- matches[!is.na(matches)]
  
  if (length(matches) > 0) {
    # 清理结果并去除首尾空白/标点
    clean_matches <- trimws(gsub("^[,\\.;\\s]+|[,\\.;\\s]+$", "", matches))
    
    # 精准去重（保留顺序）
    unique_matches <- unique(clean_matches)
    
    return(paste(unique_matches, collapse = "; "))
  } else {
    return(NA)
  }
}

# 设置文件路径
input_file <- "C:/Users/Administrator/Desktop/re-part1/FK/amp_FK_1.xlsx"
output_file <- "C:/Users/Administrator/Desktop/re-part1/FK/amp_FK_only.xlsx"

# 读取Excel数据
data <- read_excel(input_file)

# 应用提取函数到目标列
data$Extracted_Klebsiella <- sapply(data$Target_Organism, extract_klebsiella)

# 重新排列列顺序（将新列放在目标列后）
target_idx <- which(names(data) == "Target_Organism")
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

# 创建新工作簿并设置样式
wb <- createWorkbook()
addWorksheet(wb, "Sheet1")

# 将数据写入工作簿
writeData(wb, sheet = 1, data, startRow = 1, startCol = 1)

# 创建蓝色样式
blueStyle <- createStyle(fgFill = "#CCE5FF") # 浅蓝色背景

# 识别需要人工处理的行（Extracted_Klebsiella为NA但Target_Organism不为空）
need_manual_rows <- which(is.na(data$Extracted_Klebsiella) & !is.na(data$Target_Organism) & data$Target_Organism != "")

# 为需要人工处理的行应用蓝色样式
if (length(need_manual_rows) > 0) {
  # 注意：Excel行号从1开始，而我们数据从第2行开始（第1行是标题）
  excel_rows <- need_manual_rows + 1
  
  # 应用样式到整行
  for (row in excel_rows) {
    addStyle(wb, sheet = 1, blueStyle, rows = row, cols = 1:ncol(data), gridExpand = TRUE)
  }
  
  cat("标记了", length(need_manual_rows), "行需要人工处理的数据\n")
} else {
  cat("没有发现需要人工处理的数据\n")
}

# 保存处理后的数据
saveWorkbook(wb, output_file, overwrite = TRUE)

cat("处理完成！结果已保存至:", output_file)