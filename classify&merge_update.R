library(readxl)
library(writexl)
library(openxlsx)
library(dplyr)
library(stringr)

# ===== 第一部分：根据靶标菌描述筛选数据 =====

# 主函数
filter_sequences_by_organism <- function(input_file, 
                                         target_column = "Target_Organism",
                                         target_organism_patterns) {
  
  # 读取Excel文件
  cat("正在读取文件:", input_file, "\n")
  data <- read_excel(input_file)
  
  # 检查目标列是否存在
  if (!target_column %in% colnames(data)) {
    stop("错误: 文件中未找到列 '", target_column, "'")
  }
  
  # 构建正则表达式模式（不区分大小写）
  # 将多个模式用|连接，表示"或"的关系
  regex_pattern <- paste(target_organism_patterns, collapse = "|")
  
  cat("使用正则表达式模式:", regex_pattern, "\n")
  
  # 筛选包含目标菌描述的序列
  filtered_data <- data %>%
    filter(str_detect(.[[target_column]], 
                      regex(regex_pattern, ignore_case = TRUE)))
  
  cat("找到", nrow(filtered_data), "条匹配的序列\n")
  
  return(filtered_data)
}

# 菌的多种表示方式模式
acinetobacter_patterns <- c(
  "Acinetobacter baumannii",   # 全名
  "A\\. baumannii",            # A. baumannii（带空格）
  "A\\.baumannii",             # A.baumannii（不带空格）
  "A baumannii",               # A baumannii（不带点）
  "Abaumannii",                # Abaumannii（紧凑形式）
  "鲍曼不动杆菌",             # 中文名称
  "Acinetobacter baumannii\\s*\\(.*\\)"  # 可能带有括号的变体
)

# 处理第一个文件：05_dbAMP_clean.xlsx
cat("\n=== 处理dbAMP文件 ===\n")
result1 <- filter_sequences_by_organism(
  input_file = "C:/Users/Administrator/Desktop/re-part1/data/05_dbAMP_clean.xlsx",
  target_organism_patterns = acinetobacter_patterns
)

# 处理第二个文件：06_DRAMP_clean.xlsx
cat("\n=== 处理DRAMP文件 ===\n")
result2 <- filter_sequences_by_organism(
  input_file = "C:/Users/Administrator/Desktop/re-part1/data/06_DRAMP_clean.xlsx",
  target_organism_patterns = acinetobacter_patterns
)

# 打印第一部分汇总信息
cat("\n=== 筛选完成 ===\n")
cat("文件1筛选结果:", nrow(result1), "条序列\n")
cat("文件2筛选结果:", nrow(result2), "条序列\n")

# ===== 第二部分：合并和去重处理 =====

cat("\n=== 开始合并处理 ===\n")

# 设置文件路径
output_dir <- "C:/Users/Administrator/Desktop/re-part1/BM"

# 确保输出目录存在
if (!dir.exists(output_dir)) {
  dir.create(output_dir, recursive = TRUE)
  cat("创建输出目录:", output_dir, "\n")
}

# 将第一部分的结果赋值给变量
dbamp <- result1
dramp <- result2

# 确保Sequence列是字符类型
dbamp$Sequence <- as.character(dbamp$Sequence)
dramp$Sequence <- as.character(dramp$Sequence)

# 1. 检查列名问题
cat("dbAMP列名:", names(dbamp), "\n")
cat("dramp列名:", names(dramp), "\n")

# 2. 修正Target_Organism列名（如果实际列名不同）
# 查找可能的列名变体
possible_names <- c("Target_Organism", "Target organism", "Target.Organism", 
                    "target_organism", "TargetOrganism", "Organism_Target")

# 找到实际存在的列名
target_col <- intersect(possible_names, names(dramp))[1]

# 如果没有找到，停止执行
if(is.na(target_col)) {
  stop("未找到Target_Organism列！请检查dramp文件中的列名。实际列名：", 
       paste(names(dramp), collapse = ", "))
} else {
  cat("使用列名:", target_col, "作为Target_Organism列\n")
}

# 3. 处理重复序列
# 对两个数据集按Sequence去重（保留每个序列的第一行记录）
dbamp <- dbamp %>% distinct(Sequence, .keep_all = TRUE)
dramp <- dramp %>% distinct(Sequence, .keep_all = TRUE)

# 创建函数检查是否包含数字信息
contains_digits <- function(x) {
  if (is.na(x) || x == "") return(FALSE)
  str_detect(x, "\\d")  # 使用正则表达式检测数字
}

# 找出需要更新的序列
update_rows <- dbamp %>%
  filter(Sequence %in% dramp$Sequence) %>%  # 只保留共同序列
  # 添加DRAMP中对应序列的Target信息
  left_join(select(dramp, Sequence, Target_Org = all_of(target_col)), 
            by = "Sequence", relationship = "one-to-one") %>%
  # 检查DRAMP中的Target是否缺失数字
  rowwise() %>%
  mutate(no_digits_in_dramp = !contains_digits(Target_Org)) %>%
  ungroup() %>%
  filter(no_digits_in_dramp) %>%
  select(-no_digits_in_dramp, -Target_Org)  # 移除临时列

# 找出全新序列（不存在于DRAMP中）
new_rows <- dbamp %>% 
  filter(!Sequence %in% dramp$Sequence)

# 合并需要添加的数据
rows_to_add <- bind_rows(update_rows, new_rows) %>%
  select(any_of(names(dramp)))  # 保留dramp中存在的列

# 添加缺失的列
for (col in names(dramp)) {
  if (!(col %in% names(rows_to_add))) {
    rows_to_add[[col]] <- NA
  }
}

# 从DRAMP中移除需要更新的序列
dramp_updated <- dramp %>%
  filter(!Sequence %in% update_rows$Sequence)

# 合并数据
combined <- bind_rows(dramp_updated, rows_to_add) %>%
  distinct(Sequence, .keep_all = TRUE)

# ===== 新增功能：ID列处理和列筛选 =====

cat("\n=== 处理ID列 ===\n")

# 从原始文件中重新读取完整的ID信息
dbamp_original <- read_excel("C:/Users/Administrator/Desktop/re-part1/data/05_dbAMP_clean.xlsx")
dramp_original <- read_excel("C:/Users/Administrator/Desktop/re-part1/data/06_DRAMP_clean.xlsx")

# 确保第一列是ID列
dbamp_id_col <- names(dbamp_original)[1]
dramp_id_col <- names(dramp_original)[1]

cat("dbAMP ID列名:", dbamp_id_col, "\n")
cat("DRAMP ID列名:", dramp_id_col, "\n")

# 为两个原始数据集创建包含ID和Sequence的查找表
dbamp_lookup <- dbamp_original %>% 
  select(Sequence, ID_dbAMP = all_of(dbamp_id_col)) %>% 
  distinct(Sequence, .keep_all = TRUE)

dramp_lookup <- dramp_original %>% 
  select(Sequence, ID_DRAMP = all_of(dramp_id_col)) %>% 
  distinct(Sequence, .keep_all = TRUE)

# 调试：检查查找表的内容
cat("dbAMP查找表记录数:", nrow(dbamp_lookup), "\n")
cat("DRAMP查找表记录数:", nrow(dramp_lookup), "\n")

# 确定每条序列的最终来源并分配对应的ID
# 首先添加来源标记
combined_with_source <- combined %>%
  mutate(
    source = ifelse(Sequence %in% dramp_updated$Sequence, "DRAMP", "dbAMP")
  )

# 然后分别添加两个数据集的ID
combined_with_ids <- combined_with_source %>%
  left_join(dramp_lookup, by = "Sequence") %>%
  left_join(dbamp_lookup, by = "Sequence")

# 调试：检查列名
cat("合并后的列名:", paste(names(combined_with_ids), collapse = ", "), "\n")

# 根据来源选择正确的ID
combined_with_ids <- combined_with_ids %>%
  mutate(
    ID = case_when(
      source == "DRAMP" & !is.na(ID_DRAMP) ~ as.character(ID_DRAMP),
      source == "dbAMP" & !is.na(ID_dbAMP) ~ as.character(ID_dbAMP),
      TRUE ~ "Unknown"  # 如果都没有，设为Unknown
    )
  ) %>%
  # 将ID列移到第一列，并移除临时列
  select(ID, everything(), -source, -ID_DRAMP, -ID_dbAMP)

cat("ID列处理完成，共", nrow(combined_with_ids), "条记录\n")
cat("序列来源统计:\n")
cat("  - 来自DRAMP数据集:", sum(combined_with_source$source == "DRAMP"), "\n")
cat("  - 来自dbAMP数据集:", sum(combined_with_source$source == "dbAMP"), "\n")

# 功能2：筛选需要的列
cat("\n=== 筛选需要的列 ===\n")

# 定义需要保留的列（更全面的列名列表）
required_columns <- c(
  "ID", 
  "Sequence", 
  # 名称相关列
  "name", "Name", "NAME", "AMP_name", "AMP_Name", 
  # 结构相关列
  "structure", "Structure", "STRUCTURE", 
  # 靶标菌相关列
  "target_organism", "Target_Organism", "Target organism", "target organism",
  "Target_organism", "Organism", "organism",
  # 线性/环状/分支相关列
  "linear/cyclic/branched", "Linear/Cyclic/Branched", "structure_type", "Structure_type",
  "Structure_Type", "Type", "type", "Linear_Cyclic_Branched",
  # N-末端相关列
  "N-terminal", "N_terminal", "N-Terminal", "N_terminal_modification",
  "N_terminal", "N-terminal_modification", "N_terminal_mod", "N-terminal_Modification",
  # C-末端相关列
  "C-terminal", "C_terminal", "C-Terminal", "C_terminal_modification",
  "C_terminal", "C-terminal_modification", "C_terminal_mod", "C-terminal_Modification",
  # 其他修饰相关列
  "Other_Modifications", "Other_modifications", "other_modifications", "Modifications",
  "Modification", "modification", "Other_mods"
)

# 找出实际存在于数据中的列
existing_columns <- character(0)
for (col in required_columns) {
  if (col %in% names(combined_with_ids)) {
    existing_columns <- c(existing_columns, col)
  }
}

cat("找到以下列将被保留:", paste(existing_columns, collapse = ", "), "\n")

# 如果没有任何匹配的列，报错
if (length(existing_columns) == 0) {
  stop("错误: 没有找到任何需要保留的列。请检查列名。")
}

# 筛选列
final_data <- combined_with_ids %>%
  select(all_of(existing_columns))

# 标准化列名
standard_names <- c(
  "ID" = "ID",
  "Sequence" = "Sequence",
  "name" = "Name", "Name" = "Name", "NAME" = "Name", "AMP_name" = "Name", "AMP_Name" = "Name",
  "structure" = "Structure", "Structure" = "Structure", "STRUCTURE" = "Structure",
  "target_organism" = "Target_Organism", "Target_Organism" = "Target_Organism", 
  "Target organism" = "Target_Organism", "target organism" = "Target_Organism",
  "Target_organism" = "Target_Organism", "Organism" = "Target_Organism", "organism" = "Target_Organism",
  "linear/cyclic/branched" = "Structure_Type", 
  "Linear/Cyclic/Branched" = "Structure_Type",
  "structure_type" = "Structure_Type", "Structure_type" = "Structure_Type",
  "Structure_Type" = "Structure_Type", "Type" = "Structure_Type", "type" = "Structure_Type",
  "Linear_Cyclic_Branched" = "Structure_Type",
  "N-terminal" = "N_terminal", "N_terminal" = "N_terminal", 
  "N-Terminal" = "N_terminal", "N_terminal_modification" = "N_terminal",
  "N_terminal" = "N_terminal", "N-terminal_Modification" = "N_terminal", "N_terminal_mod" = "N_terminal",
  "C-terminal" = "C_terminal", "C_terminal" = "C_terminal", 
  "C-Terminal" = "C_terminal", "C_terminal_modification" = "C_terminal",
  "C_terminal" = "C_terminal", "C-terminal_Modification" = "C_terminal", "C_terminal_mod" = "C_terminal",
  "Other_Modifications" = "Other_Modifications", 
  "Other_modifications" = "Other_Modifications", 
  "other_modifications" = "Other_Modifications", 
  "Modifications" = "Other_Modifications",
  "Modification" = "Other_Modifications", "modification" = "Other_Modifications",
  "Other_mods" = "Other_Modifications"
)

# 应用标准列名
for (i in 1:ncol(final_data)) {
  old_name <- names(final_data)[i]
  if (old_name %in% names(standard_names)) {
    new_name <- standard_names[old_name]
    names(final_data)[i] <- new_name
  }
}

# 确保ID在第一列
final_data <- final_data %>%
  select(ID, everything())

cat("列筛选完成，最终保留", ncol(final_data), "列\n")
cat("最终列名:", paste(names(final_data), collapse = ", "), "\n")

# 保存结果
output_file <- file.path(output_dir, "amp_BM_1.xlsx")
write.xlsx(final_data, output_file)

# 输出最终统计信息
cat("\n===== 处理完成 =====\n")
cat("原始dbAMP记录数:", nrow(result1), "\n")
cat("原始DRAMP记录数:", nrow(result2), "\n")
cat("去重后dbAMP记录数:", nrow(dbamp), "\n")
cat("去重后DRAMP记录数:", nrow(dramp), "\n")
cat("共同序列数量:", sum(dbamp$Sequence %in% dramp$Sequence), "\n")
cat("需要更新的序列数量:", nrow(update_rows), "\n")
cat("全新序列数量:", nrow(new_rows), "\n")
cat("合并后总记录数:", nrow(combined), "\n")
cat("最终筛选后记录数:", nrow(final_data), "\n")
cat("最终保留列数:", ncol(final_data), "\n")
cat("序列来源统计:\n")
cat("  - 来自DRAMP数据集:", sum(combined_with_source$source == "DRAMP"), "\n")
cat("  - 来自dbAMP数据集:", sum(combined_with_source$source == "dbAMP"), "\n")
cat("最终结果已保存到:", output_file, "\n")