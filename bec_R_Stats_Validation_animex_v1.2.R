################################################################################
#                                                                              #
#     Validation Check of the Annual Animal Experimentation Statistics         #
#                                                                              #
#     Version: v1.2 (26.3.2025)                                                #
#     Author: Cristian Berce - Federal Food Safety and Veterinary Office FSVO  #
#     Description: Validation check of the data for the annual                 #
#                      animal experimentation statistics from the animex       #
#                      generated AC Reports (the xlsx Files)                   #
#     Email: oberaufsicht-tv@blv.admin.ch                                      #
#                                                                              #
################################################################################


### The Validation rules in this script are :
### Rule 1: 9.2 needs to be greater than 9.4.1
### Rule 2: 10.1 must equal the sum of 9.1.1 + 9.1.2 + 9.1.3 + 9.1.4 + 9.4.2
### Rule 3 11.1 + 11.2 + 11.3 + 11.4 must equal to the sum of 9.2 + 9.1.1 + 9.1.2 + 9.1.3 + 9.1.4
### Rule 4: 12.1 + 12.2.1 + 12.2.2 + 12.2.3 + 12.2.4 + 12.3 must equal to the sum of 10.1 + 10.2 + 10.3 + 10.4
### Rule 5: If 9.4.1 is bigger than 0 then 9.2 must be greater than the sum of 10.1 to 10.4
### If either of these rules fails then it means your report was not correctly filled out and will be returned to you for corrections, either by the Cantonal Office or by the FSVO

library(readxl)
library(openxlsx)
library(dplyr)

# This is where you set your paths. Please update with the relevant paths
input_dir <- "C:/Rtests/statistics/animal_excels/a" # Modfiy this with the path of the folder where you will save your xlsx files
output_file <- "C:/Rtests/statistics/Validation_Report.xlsx" # Modify this with the path of the folder where you want your validation report

# Chapter keys that will be extracted from your reports - MODIFY STUFF BELOW AT YOUR OWN RISK - IT MAY RENDER THE SCRIPT UNFUNCTIONAL
chapter_keys <- c(
  "9.1.1", "9.1.2", "9.1.3", "9.1.4", "9.2", "9.4.1", "9.4.2",
  "10.1", "10.2", "10.3", "10.4",
  "11.1", "11.2", "11.3", "11.4",
  "12.1", "12.2.1", "12.2.2", "12.2.3", "12.2.4", "12.3"
)

# Validation rules start below
process_block <- function(values, category, filename) {
  rows <- list()
  col_9_1 <- as.numeric(values[c("9.1.1", "9.1.2", "9.1.3", "9.1.4")])
  col_9_2 <- as.numeric(values["9.2"])
  col_9_4_1 <- as.numeric(values["9.4.1"])
  col_9_4_2 <- as.numeric(values["9.4.2"])
  col_10 <- as.numeric(values[c("10.1", "10.2", "10.3", "10.4")])
  col_11 <- as.numeric(values[c("11.1", "11.2", "11.3", "11.4")])
  col_12 <- as.numeric(values[c("12.1", "12.2.1", "12.2.2", "12.2.3", "12.2.4", "12.3")])
  
  is_empty <- all(is.na(c(col_9_1, col_9_2, col_9_4_1, col_9_4_2, col_10, col_11, col_12))) ||
    all(c(col_9_1, col_9_2, col_9_4_1, col_9_4_2, col_10, col_11, col_12) == 0, na.rm = TRUE)
  
  if (!is_empty) {
    if (!is.na(col_9_2) && !is.na(col_9_4_1) &&
        !(col_9_2 == 0 && col_9_4_1 == 0) &&
        !(col_9_2 > col_9_4_1)) {
      rows[[length(rows)+1]] <- c(list(File = filename, Category = category, Failed_Rule = "Rule 1: 9.2 not > 9.4.1"), as.list(values))
    }
    
    if (!(values[["10.1"]] == sum(col_9_1) + col_9_4_2)) {
      rows[[length(rows)+1]] <- c(list(File = filename, Category = category, Failed_Rule = "Rule 2: 10.1 ≠ 9.1.x + 9.4.2"), as.list(values))
    }
    if (!(sum(col_11) == sum(col_9_2, col_9_1))) {
      rows[[length(rows)+1]] <- c(list(File = filename, Category = category, Failed_Rule = "Rule 3: 11.x ≠ 9.x"), as.list(values))
    }
    if (!(sum(col_12) == sum(col_10))) {
      rows[[length(rows)+1]] <- c(list(File = filename, Category = category, Failed_Rule = "Rule 4: 12.x ≠ 10.x"), as.list(values))
    }
    if (col_9_4_1 > 0 && !(col_9_2 > sum(col_10))) {
      rows[[length(rows)+1]] <- c(list(File = filename, Category = category, Failed_Rule = "Rule 5: 9.2 not > sum of 10.x when 9.4.1 > 0"), as.list(values))
    }
  }
  
  if (length(rows) == 0 && !is_empty) {
    rows[[1]] <- c(list(File = filename, Category = category, Failed_Rule = NA), as.list(values))
  }
  
  return(rows)
}

# Main loop
files <- list.files(input_dir, pattern = "*.xlsx", full.names = TRUE)
results <- list()

for (file in files) {
  sheet <- read_excel(file, sheet = "Sheet0", col_names = FALSE)
  filename <- tools::file_path_sans_ext(basename(file))
  
  current_category <- NA
  block_values <- list()
  block_values$values <- setNames(rep(NA, length(chapter_keys)), chapter_keys)
  
  for (i in seq_len(nrow(sheet))) {
    row <- sheet[i, ]
    first_col <- as.character(row[[1]])
    
    if (!is.na(first_col) && trimws(first_col) == "8") {
      if (!is.null(current_category)) {
        rows <- process_block(block_values$values, current_category, filename)
        results <- append(results, rows)
        block_values$values <- setNames(rep(NA, length(chapter_keys)), chapter_keys)
      }
      current_category <- trimws(as.character(row[[2]]))  # Column B
    }
    
    chap_key <- trimws(first_col)
    if (!is.na(chap_key) && chap_key %in% chapter_keys) {
      for (j in 0:2) {
        if ((i + j) <= nrow(sheet)) {
          subrow <- sheet[i + j, ]
          if (toupper(trimws(as.character(subrow[[8]]))) == "INST") {
            inst_value <- suppressWarnings(as.numeric(subrow[[9]]))
            block_values$values[[chap_key]] <- inst_value
            break
          }
        }
      }
    }
  }
  
  if (!is.null(current_category)) {
    rows <- process_block(block_values$values, current_category, filename)
    results <- append(results, rows)
  }
}

# Excel output
result_df <- bind_rows(results)
wb <- createWorkbook()
addWorksheet(wb, "Validation_Results")
writeData(wb, sheet = 1, result_df)

# Taste the rainbow :)
style_rule1 <- createStyle(fgFill = "#FF9999")
style_rule2 <- createStyle(fgFill = "#FFD580")
style_rule3 <- createStyle(fgFill = "#FFFF99")
style_rule4 <- createStyle(fgFill = "#ADD8E6")
style_rule5 <- createStyle(fgFill = "#D9B3FF")
style_valid <- createStyle(fgFill = "#CCFFCC")

fail_col <- which(colnames(result_df) == "Failed_Rule")

for (i in 1:nrow(result_df)) {
  value <- result_df[i, fail_col]
  if (is.na(value) || value == "") {
    addStyle(wb, 1, style_valid, rows = i + 1, cols = fail_col, gridExpand = TRUE)
  } else {
    if (grepl("Rule 1", value)) addStyle(wb, 1, style_rule1, rows = i + 1, cols = fail_col, gridExpand = TRUE)
    if (grepl("Rule 2", value)) addStyle(wb, 1, style_rule2, rows = i + 1, cols = fail_col, gridExpand = TRUE)
    if (grepl("Rule 3", value)) addStyle(wb, 1, style_rule3, rows = i + 1, cols = fail_col, gridExpand = TRUE)
    if (grepl("Rule 4", value)) addStyle(wb, 1, style_rule4, rows = i + 1, cols = fail_col, gridExpand = TRUE)
    if (grepl("Rule 5", value)) addStyle(wb, 1, style_rule5, rows = i + 1, cols = fail_col, gridExpand = TRUE)
  }
}

saveWorkbook(wb, output_file, overwrite = TRUE)
cat("Validation complete. Output saved to:", output_file, "\n")
