# Load packages
packages_all = c("stringr")
packages_installed = packages_all %in% rownames(installed.packages())
if (any(packages_installed == FALSE)) {
    install.packages(packages_all[!packages_installed])
}
invisible(lapply(packages_all, library, character.only = TRUE))



# Load global constants
source('CONSTANTS.R')
# Load local constants



# read excel file
suppressMessages(
    df <- readxl::read_excel(paste0("input/xlsx/study-summary/study-summary.xlsx"), col_names = FALSE)
)
print('Loaded cohort description file')



secondary_analyses <- 0
study_rows <- 0
unique_cohorts <- c()
primary_study_ids <- c()
# loop all rows in the data frame
for (i in 1:nrow(df)) {
    # do not read rows where it is written who screened and extracted data, respectively
    if (!grepl( '–', df[i, 1][[1]], fixed = TRUE)) {
        study_id <- df[i, 1][[1]]
        trajectories_from <- df[i, 3][[1]]
        cohort <- df[i, 2][[1]]
        # split cohort by ";" into an array and remove excess space before/after string
        cohorts <- str_trim(unlist(str_split(cohort, ';')))
        # check if any cohorts are not already in the list of unique cohorts
        for (cohort in cohorts) {
            if (!(cohort %in% unique_cohorts)) {
                unique_cohorts <- c(unique_cohorts, cohort)
            }
        }
        is_secondary_analysis <- df[i, 4][[1]]
        print(is_secondary_analysis)
        if (!is.na(is_secondary_analysis)) {
            secondary_analyses <- secondary_analyses + 1
        } else {
            primary_study_ids <- c(primary_study_ids, study_id)
        }
        study_rows <- study_rows + 1
    }
}
primary_analyses <- study_rows - secondary_analyses
print('Study:')
print(study_rows)
print('Unique cohorts:')
print(unique_cohorts)
print('Number of unique cohorts:')
print(length(unique_cohorts))
print(paste0('Primary analyses: ', primary_analyses))
print(paste0('Secondary analyses: ', secondary_analyses))
print(paste0('Total: ', primary_analyses + secondary_analyses))
# save the summary data to a .txt file
write(paste0('Unique cohorts: ', length(unique_cohorts), '\n', 'Primary analyses: ', primary_analyses, '\n', 'Secondary analyses: ', secondary_analyses, '\n', 'Total: ', primary_analyses + secondary_analyses), file = paste0('output/txt/', 'study-level-summary.txt'))
# save the primary study ids to a .txt file
write(paste(primary_study_ids, collapse = ','), file = paste0('output/txt/', 'primary-study-ids.txt'))