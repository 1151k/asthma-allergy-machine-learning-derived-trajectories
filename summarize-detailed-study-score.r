# Load packages
packages_all = c("data.table", "readxl", "ggplot2", "dplyr", "stringr")
packages_installed = packages_all %in% rownames(installed.packages())
if (any(packages_installed == FALSE)) {
    install.packages(packages_all[!packages_installed])
}
invisible(lapply(packages_all, library, character.only = TRUE))



# Load global constants
source('CONSTANTS.R')
# Load local constants
primary_study_list <- readLines(paste0(folder_path, "output/txt/primary-study-ids.txt"))
primary_study_list <- unlist(strsplit(primary_study_list, ","))
# add to primary_study_list the following (Caudri 2013 [for Savenije 2011], Collin 2015 [for Henderson 2008] and Kotsapas 2022 [for Simpson 2010])
primary_study_list <- c(primary_study_list, "Caudri 2013", "Collin 2015", "Kotsapas 2022")
print(primary_study_list)
# quit()



cat('READING EXCEL FILES\n')
cat('-------- start\n')
# List all Excel files in the folder with names containing "Data extraction"
excel_files <- list.files(path = paste0(folder_path, "input/xlsx"), pattern = "^[^~].*Quality assessment.xlsx$", full.names = TRUE)
# Check if there are any matching files
if (length(excel_files) == 0) {
    stop("No files found.")
}
print(paste("Loaded ", length(excel_files), " Excel files", sep = ""))
cat('-------- end\n')



cat('EXTRACTING DATA\n')
cat('-------- start\n')
score_data <- data.table(
    study_ID = character(),
    score = numeric()
)
# Loop through each Excel file
i <- 0
for (file in excel_files) {
    # Read the Excel file
    suppressMessages(
        df <- read_excel(file, col_names = FALSE)
    )

    # get data items
    study_ID = df[3, 3][[1]]
    if (study_ID %in% primary_study_list) {
        # C (withdrawals and drop-outs)
        missingness_1 <- ifelse(grepl("1\\.|3\\.", df[31,3][[1]]), 1, 0)
        # D (preprocessing)
        missingness_2 <- ifelse(grepl("1\\.", df[39,3][[1]]), 1, 0)
        time_metrics_variance <- ifelse(grepl("1\\.", df[41,3][[1]]), 1, 0)
        noise_variation <- ifelse(grepl("1\\.", df[43,3][[1]]), 1, 0)
        trajectory_variable_selection <- ifelse(grepl("1\\.", df[45,3][[1]]), 1, 0)
        data_processing <- ifelse(grepl("1\\.", df[47,3][[1]]), 1, 0)
        preprocessing_reproducibility <- ifelse(grepl("1\\.", df[49,3][[1]]), 2, 0)
        # E (trajectory modelling)
        modelling_selection_basis <- ifelse(grepl("1\\.", df[57,3][[1]]), 1, 0)
        modelling_optimization <- ifelse(grepl("1\\.", df[59,3][[1]]), 2, 0)
        optimal_model_selection <- ifelse(grepl("1\\.", df[61,3][[1]]), 2, 0)
        multiple_modelling_techniques <- ifelse(grepl("1\\.", df[63,3][[1]]), 1, 0)
        modelling_replicability <- ifelse(grepl("1\\.", df[65,3][[1]]), 2, 0)
        # G (evaluation and reporting of results)
        modelling_reporting <- ifelse(grepl("1\\.", df[89,3][[1]]), 2, 0)
        external_validation <- ifelse(grepl("1\\.", df[91,3][[1]]), 1, 0)
        limitations_report <- ifelse(grepl("1\\.", df[93,3][[1]]), 2, 0)

        print(study_ID)
        print(paste('Missingness 1: ', missingness_1, sep = ''))
        print(paste('Missingness 2: ', missingness_2, sep = ''))
        print(paste('Time metrics variance: ', time_metrics_variance, sep = ''))
        print(paste('Noise variation: ', noise_variation, sep = ''))
        print(paste('Trajectory variable selection: ', trajectory_variable_selection, sep = ''))
        print(paste('Data processing: ', data_processing, sep = ''))
        print(paste('Preprocessing reproducibility: ', preprocessing_reproducibility, sep = ''))
        print(paste('Modelling selection basis: ', modelling_selection_basis, sep = ''))
        print(paste('Modelling optimization: ', modelling_optimization, sep = ''))
        print(paste('Optimal model selection: ', optimal_model_selection, sep = ''))
        print(paste('Multiple modelling techniques: ', multiple_modelling_techniques, sep = ''))
        print(paste('Modelling replicability: ', modelling_replicability, sep = ''))
        print(paste('Modelling reporting: ', modelling_reporting, sep = ''))
        print(paste('External validation: ', external_validation, sep = ''))
        print(paste('Limitations report: ', limitations_report, sep = ''))

        # summarize scores
        score <- missingness_1 + missingness_2 + time_metrics_variance + noise_variation + data_processing + preprocessing_reproducibility + modelling_selection_basis + modelling_optimization + optimal_model_selection + multiple_modelling_techniques + modelling_replicability + modelling_reporting + external_validation + limitations_report
        print(paste('Score: ', score, sep = ''))

        # add to data table
        score_data <- rbind(score_data, data.table(study_ID = study_ID, score = score))
        print('---')
    }
}
print('Done with all files')
cat('-------- end\n')
print(paste("Total number of studies: ", i, sep = ""))



cat('SAVING DATA EXTRACTION\n')
cat('-------- start\n')
p <- ggplot(score_data, aes(x = study_ID, y = score, color = score)) +
    geom_point(size = 2) +
    geom_text(aes(label = paste0(sub(" 2", "\n2", study_ID), " (", score, ")")), size = 3, vjust = 1) + #, angle = 90, vjust = 3, hjust = 0.5) +
    scale_color_gradient2(low = "#a64d4c", mid = "#d2a635", high = "#5d9969", midpoint = 13) +
    theme(
        panel.background = element_rect(fill = "transparent"),
        axis.line = element_line(color='black', linewidth = 0.75),
        axis.text.x = element_blank(),
        axis.text.y = element_text(size = 16, color = "black", margin = margin(r = 5)),
        axis.title.x =element_text(size = 16, color = "black", margin = margin(r = 9)),
        axis.title.y = element_text(size = 18, color = "black", margin = margin(r = 9)),
        axis.ticks.length = unit(0.2, "cm"),
        legend.position = "none",
        axis.ticks.x = element_blank()
    ) +
    labs(
        # title = plot_name,
        x = "Study ID",
        y = "Computational methodology robustness score"
    ) +
  scale_y_continuous(breaks = seq(from = 0, to = 21, by = 3), limits = c(0, 21), expand = c(0, 0))




ggsave(paste0(folder_path, "output/svg/detailed-study-score.svg"), p, width = 14.2, height = 6.5)
cat('-------- end\n')