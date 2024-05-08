# Load packages
packages_all = c("data.table", "readxl", "reshape2")
packages_installed = packages_all %in% rownames(installed.packages())
if (any(packages_installed == FALSE)) {
    install.packages(packages_all[!packages_installed])
}
invisible(lapply(packages_all, library, character.only = TRUE))



# Load local constants
root_folder <- "/Users/xlisda/Downloads/it/server/research/phd/taa_sr/output/docx/done"



# count the numbr of unique reports included in meta-analysis
unique_reports_in_meta_analyses <- c()
xlsx_files <- list.files(path = root_folder, pattern = "^meta-analysis-.*\\.xlsx$", recursive = TRUE, full.names = TRUE)
# go through each file and extract the report name
for (xlsx_file in xlsx_files) {
    # Read the Excel file
    suppressMessages(df <- read_excel(xlsx_file))
    # rename Lodge 2014a to Lodge 20141, Lodge 2014b to Lodge 20142, Shahunja 2022a to Shahunja 20221, Shahunja 2022b to Shahunja 20222, Hu 2020a to Hu 20201, Hu 2020b to Hu 20202
    df$data_id <- gsub("Lodge 2014a", "Lodge 20141", df$data_id)
    df$data_id <- gsub("Lodge 2014b", "Lodge 20142", df$data_id)
    df$data_id <- gsub("Shahunja 2022a", "Shahunja 20221", df$data_id)
    df$data_id <- gsub("Shahunja 2022b", "Shahunja 20222", df$data_id)
    df$data_id <- gsub("Hu 2020a", "Hu 20201", df$data_id)
    df$data_id <- gsub("Hu 2020b", "Hu 20202", df$data_id)
    df$data_id <- gsub("([0-9]{4})[a-z]", "\\1", df$data_id)
    # remove any text/values in parentheses
    df$data_id <- gsub("\\(.*\\)", "", df$data_id)
    # remove any whitespace at beginning/end of string
    df$data_id <- trimws(df$data_id)
    # remove any non-unique values
    data_id <- unique(df$data_id)
    # add reports (values in data_id) not already in unique_reports_in_meta_analyses to unique_reports_in_meta_analyses
    unique_reports_in_meta_analyses <- c(unique_reports_in_meta_analyses, data_id[!data_id %in% unique_reports_in_meta_analyses])
}
print('Number of unique reports included in meta-analyses:')
print(length(unique_reports_in_meta_analyses))
print('Unique reports included in meta-analyses:')
print(unique_reports_in_meta_analyses)



# count the number of unique meta-analyses performed
svg_files <- list.files(path = root_folder, pattern = "^raw-forest-plots-.*\\.svg$", recursive = TRUE, full.names = TRUE)
print('Number of unique meta-analyses performed:')
print(length(svg_files))