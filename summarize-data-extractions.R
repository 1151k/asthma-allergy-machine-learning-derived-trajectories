# Load packages
packages_all = c("data.table", "readxl", "officer", "flextable", "dplyr", "stringr")
packages_installed = packages_all %in% rownames(installed.packages())
if (any(packages_installed == FALSE)) {
    install.packages(packages_all[!packages_installed])
}
invisible(lapply(packages_all, library, character.only = TRUE))



# Load global constants
source('CONSTANTS.R')
# Load local constants
# data.table to hold all data extraction items
data_extraction <- data.table(
    study_id_and_country = character(),
    n_subjects = character(),
    source_and_characteristics = character(),
    subject_age = character(),
    diseases = character(),
    trajectory_modelling = character(),
    model_selection = character(),
    trajectory_characteristics = character()
)



cat('READING EXCEL FILES\n')
cat('-------- start\n')
# List all Excel files in the folder with names containing "Data extraction"
excel_files <- list.files(path = paste0(folder_path, "input/xlsx"), pattern = "^[^~].*Data extraction.xlsx$", full.names = TRUE)
# Check if there are any matching files
if (length(excel_files) == 0) {
    stop("No files found.")
}
print(paste("Loaded ", length(excel_files), " Excel files", sep = ""))
cat('-------- end\n')



cat('EXTRACTING DATA\n')
cat('-------- start\n')
# hold studies for tabulation
unique_diseases <- c()
added_studies <- c() # studies that have been added to any list
asthma_wheezing_studies <- c()
AD_eczema_studies <- c()
AR_rhinitis_rhinoconjunctivitis_studies <- c()
FA_studies <- c()
atopy_studies <- c()
multiple_diseases_studies <- c()
secondary_analysis_studies <- c()
# Loop through each Excel file
for (file in excel_files) {
    # Read the Excel file
    suppressMessages(
        df <- read_excel(file, col_names = FALSE)
    )

    # get data items
    study_ID = df[3, 3][[1]]
    year = stringr::str_extract_all(study_ID, "\\d+")[[1]]
    country = df[11, 3][[1]]
    ## study ID and country
    study_id_and_country <- paste(study_ID, country, sep = ", ")
    study_id_and_country <- trimws(gsub("  ", " ", study_id_and_country))
    ## subject information
    n_subjects = df[17, 3][[1]]
    # add space before and after em dash
    n_subjects <- gsub("–", "–", n_subjects)
    n_subjects <- gsub("—", "–", n_subjects)
    n_subjects <- gsub("-", "–", n_subjects)
    # take number to the right of – in n_subjects and divide by the number to the left of – and multiply by 100 and round to one decimal
    percentage <- round(as.numeric(gsub(",", "", strsplit(n_subjects, "–")[[1]][2])) / as.numeric(gsub(",", "", strsplit(n_subjects, "–")[[1]][1])) * 100, 1)
    n_subjects <- paste(n_subjects, " (", percentage, "%)", sep = "")
    age_span = df[19, 3][[1]]
    age_span <- gsub("–", "–", age_span)
    age_span <- gsub("—", "–", age_span)
    age_span <- gsub("-", "–", age_span)
    years_followed = df[21, 3][[1]]
    source_and_characteristics = df[23, 3][[1]]
    percentage_agreed = df[25, 3][[1]]
    percentage_completed_follow_up = df[27, 3][[1]]
    ## trajectory-defining data and preprocessing
    rationale_trajectory_variables = df[33, 3][[1]]
    trajectory_variables = df[35, 3][[1]]
    diseases_short = df[35, 5][[1]]
    trajectory_assessment_points = df[37, 3][[1]]
    n_assessments = strsplit(df[37, 5][[1]], ";")[[1]]
    preprocessing = df[39, 3][[1]]
    preprocessing_reproducibility = df[41, 3][[1]]
    ## trajectory modelling
    rationale_trajectory_modelling = df[47, 3][[1]]
    trajectory_modelling = df[49, 3][[1]]
    trajectory_modelling_optimization = df[51, 3][[1]]
    trajectory_modelling_selection = df[53, 3][[1]]
    trajectory_modelling_reproducibility = df[55, 3][[1]]
    # evaluation/validation of trajectories and associated risk factors/outcomes
    external_validation = df[61, 3][[1]]
    clinical_epidemiological_pathophysiological_evaluation = df[63, 3][[1]]
    associated_risk_factors = strsplit(df[65, 3][[1]], ",")[[1]]
    associated_outcomes = df[67, 3][[1]]
    trajectory_characteristics = df[69, 3][[1]]
    diseases_investigated = df[69, 5][[1]]

    # START DISEASES INVESTIGATED
    # remove the plus 
    diseases_investigated <- gsub("\\+", "", diseases_investigated)
    # remove the minus
    diseases_investigated <- gsub("\\-", "", diseases_investigated)
    # # check if "+" is in diseases_investigated
    # plots_possible <- FALSE
    # if (grepl("\\+", diseases_investigated)) {
    #     # print(paste("XYZ DRAW POSSIBLE: ", diseases_investigated, sep = ""))
    #     plots_possible <- TRUE
    # } else if (grepl("\\-", diseases_investigated)) {
    #     # print(paste("XYZ NO PLOTS: ", diseases_investigated, sep = ""))
    # } else {
    #     # print(paste("XYZ NOTHING AT ALL: ", diseases_investigated, sep = ""))
    # }
    # check if string length is above 0
    if (!is.na(diseases_investigated)) {
        # make list separated by comma
        diseases_investigated <- strsplit(diseases_investigated, ",")[[1]]
        # trim spaces at the beginning and end
        diseases_investigated <- trimws(diseases_investigated)
        # get number of diseases
        n_diseases_investigated <- length(diseases_investigated)
        # check if any of the itms in diseases_investigated are not in unique_diseases, add if not
        for (disease in diseases_investigated) {
            if (disease %in% unique_diseases == FALSE) {
                unique_diseases <- c(unique_diseases, disease)
            }
        }
        # for studies that do not have trajectory data for individual diseases, do not add these to disease-specific lists
        if (study_ID %in% c("Apel 2019", "Belgrave 2014", "Bougas 2019", "Bui 2021", "Dharma 2018", "Forster 2022", "Gabet 2019", "Khan 2023", "Kilanowski 2023", "Lee 2022", "Panico 2014", "Peng 2022", "Rancière 2013", "Schoos 2016", "Shi 2023", "Tang 2018")) {
            multiple_diseases_no_individual_trajectory_data <- TRUE
        } else {
            multiple_diseases_no_individual_trajectory_data <- FALSE
        }
        # check if "asthma", "wheezing", "dyspnea", "sleep disturbance", or "coughing" is in any of the diseases
        if (any(grepl("asthma", diseases_investigated, ignore.case = TRUE)) | any(grepl("wheezing", diseases_investigated, ignore.case = TRUE)) | any(grepl("dyspnea", diseases_investigated, ignore.case = TRUE)) | any(grepl("sleep disturbance", diseases_investigated, ignore.case = TRUE)) | any(grepl("coughing", diseases_investigated, ignore.case = TRUE))) {
            print(paste("ASTHMA:"))
            if (!multiple_diseases_no_individual_trajectory_data) asthma_wheezing_studies <- c(asthma_wheezing_studies, study_id_and_country)
            if (study_id_and_country %in% added_studies == FALSE) {
                added_studies <- c(added_studies, study_id_and_country)
            }
        }
        # check if "AD" or "eczema" is in any of the diseases
        if (any(grepl("AD", diseases_investigated, ignore.case = TRUE)) | any(grepl("eczema", diseases_investigated, ignore.case = TRUE))) {
            print(paste("AD/ECZEMA:"))
            if (!multiple_diseases_no_individual_trajectory_data) AD_eczema_studies <- c(AD_eczema_studies, study_id_and_country)
            if (study_id_and_country %in% added_studies == FALSE) {
                added_studies <- c(added_studies, study_id_and_country)
            }
        }
        # check if "AR" or "rhinitis" or "rhinoconjunctivitis" is in any of the diseases
        if (any(grepl("AR", diseases_investigated, ignore.case = TRUE)) | any(grepl("rhinitis", diseases_investigated, ignore.case = TRUE)) | any(grepl("rhinoconjunctivitis", diseases_investigated, ignore.case = TRUE))) {
            print(paste("AR/RHINITIS/RHINOCONJUNCTIVITIS:"))
            if (!multiple_diseases_no_individual_trajectory_data) AR_rhinitis_rhinoconjunctivitis_studies <- c(AR_rhinitis_rhinoconjunctivitis_studies, study_id_and_country)
            if (study_id_and_country %in% added_studies == FALSE) {
                added_studies <- c(added_studies, study_id_and_country)
            }
        }
        # check if "FA" is in any of the diseases
        if (any(grepl("FA", diseases_investigated, ignore.case = TRUE))) {
            print(paste("FA:"))
            if (!multiple_diseases_no_individual_trajectory_data) FA_studies <- c(FA_studies, study_id_and_country)
            if (study_id_and_country %in% added_studies == FALSE) {
                added_studies <- c(added_studies, study_id_and_country)
            }
        }
        # check if "atopy" is in any of the diseases
        if (any(grepl("atopy", diseases_investigated, ignore.case = TRUE))) {
            print(paste("ATOPY:"))
            if (!multiple_diseases_no_individual_trajectory_data) atopy_studies <- c(atopy_studies, study_id_and_country)
            if (study_id_and_country %in% added_studies == FALSE) {
                added_studies <- c(added_studies, study_id_and_country)
            }
        }
        # check if multiple diseases are investigated
        if (n_diseases_investigated > 1) {
            print("MULTIPLE DISEASES:")
            print(diseases_investigated)
            multiple_diseases_studies <- c(multiple_diseases_studies, study_id_and_country)
            if (study_id_and_country %in% added_studies == FALSE) {
                added_studies <- c(added_studies, study_id_and_country)
            }
        }
    } else {
        # secondary analysis
        secondary_analysis_studies <- c(secondary_analysis_studies, study_id_and_country)
        if (study_id_and_country %in% added_studies == FALSE) {
            added_studies <- c(added_studies, study_id_and_country)
        }
    }
    # END DISEASES INVESTIGATED
 
    # for table of characteristics
    # study ID and country
    # see above
    # number of subjects
    n_subjects = n_subjects
    # source and characteristics of subjects
    source_and_characteristics <- source_and_characteristics
    # age
    subject_age <- paste(age_span, trajectory_assessment_points, sep = "\n")
    # diseases
    diseases = trajectory_variables
    # trajectory modelling
    trajectory_modelling = trajectory_modelling
    # model selection
    trajectory_modelling_selection = trajectory_modelling_selection
    # trajectory characteristics
    trajectory_characteristics = trajectory_characteristics
    # add to table
    data_extraction <- rbind(
        data_extraction, 
        list(
            study_id_and_country = study_id_and_country,
            n_subjects = n_subjects,
            source_and_characteristics = source_and_characteristics,
            subject_age = subject_age,
            diseases = diseases,
            trajectory_modelling = trajectory_modelling,
            model_selection = trajectory_modelling_selection,
            trajectory_characteristics = trajectory_characteristics
        )
    )
    print(paste('Done with ', study_ID, sep = ""))
    print('-----')
}
print('Done with all files')
unique_diseases <- sort(unique_diseases)
print(unique_diseases)
print('----')
print(length(added_studies))
print('----')
print(asthma_wheezing_studies)
cat('-------- end\n')



cat('SAVING DATA EXTRACTION\n')
cat('-------- start\n')
# save as table in Word document
save_table <- function(study_ID_and_countries, filename) {
    # set data_table to those rows in data_extraction where there column study_id_and_country is in study_ID_and_countries
    data_table <- data_extraction[study_id_and_country %in% study_ID_and_countries]
    # data_table <- data_extraction
    doc <- officer::read_docx(path = paste0(folder_path, "input/docx/templates/landscape.dotx"))
    colnames(data_table) <- c(
        "Reference, countrya",
        "Number of subjectsb",
        "Source and characteristics of subjectsc",
        "Subject aged",
        "Outcome definitiond",
        "Trajectory modelling technique",
        "Criteria for model selection",
        "Trajectory characteristics"
    )
    ft <- flextable::flextable(data_table)
    ft <- flextable::set_table_properties(ft, layout = "autofit", align = "left")
    doc <- body_add_flextable(doc, ft)
    print(doc, target = paste0(folder_path, paste0("output/docx/data-extraction-", filename, ".docx")))
}
# save asthma/wheezing/dyspnea/sleep disturbance/coughing studies
save_table(asthma_wheezing_studies, "asthma-wheezing")
# save AD/eczema studies
save_table(AD_eczema_studies, "AD-eczema")
# save AR/rhinitis/rhinoconjunctivitis studies
save_table(AR_rhinitis_rhinoconjunctivitis_studies, "AR-rhinitis-rhinoconjunctivitis")
# save FA studies
save_table(FA_studies, "FA")
# save atopy studies
save_table(atopy_studies, "atopy")
# save multiple diseases studies
save_table(multiple_diseases_studies, "multiple-diseases")

# save the number of studies in asthma_wheezing_studies, AD_eczema_studies, AR_rhinitis_rhinoconjunctivitis_studies, FA_studies, atopy_studies, and multiple_diseases_studies to a .txt file
write.table(
    c(
        paste0("Asthma/wheezing/dyspnea/sleep disturbance/coughing: ", length(asthma_wheezing_studies)),
        paste0("AD/eczema: ", length(AD_eczema_studies)),
        paste0("AR/rhinitis/rhinoconjunctivitis: ", length(AR_rhinitis_rhinoconjunctivitis_studies)),
        paste0("FA: ", length(FA_studies)),
        paste0("Atopy: ", length(atopy_studies)),
        paste0("Multiple diseases: ", length(multiple_diseases_studies))
    ),
    file = paste0(folder_path, "output/txt/number-of-studies.txt"),
    row.names = FALSE,
    col.names = FALSE
)
cat('-------- end\n')