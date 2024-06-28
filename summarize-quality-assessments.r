# Load packages
packages_all = c("readxl", "stringr", "officer", "ggplot2")
packages_installed = packages_all %in% rownames(installed.packages())
if (any(packages_installed == FALSE)) {
    install.packages(packages_all[!packages_installed])
}
invisible(lapply(packages_all, library, character.only = TRUE))



# Load global constants
source('CONSTANTS.R')
# Load local constants
# data.frame to hold all quality assessment values
quality_assessment <- data.frame(
    studyID = character(),
    selection_bias = character(),
    data_collection_methods = character(),
    withdrawals_and_dropouts = character(),
    preprocessing = character(),
    trajectory_modelling = character(),
    associated_risk_factors_and_outcomes = character(),
    evaluation_and_reporting_of_results = character(),
    overall_rating = character()
)



# list all Excel files in the folder with names containing "Data extraction"
cat('READING EXCEL FILES\n')
cat('-------- start\n')
excel_files <- list.files(path = paste0(folder_path, "input/xlsx"), pattern = "^[^~].*Quality assessment.xlsx$", full.names = TRUE)
# check if there are any matching files
if (length(excel_files) == 0) {
    stop("No files found.")
}
print(paste("Loaded ", length(excel_files), " Excel files", sep = ""))
cat('-------- end\n')



# loop through each Excel file
cat('EXTRACTING DATA\n')
cat('-------- start\n')
secondary_analyses <- 0
secondary_ratings <- c()
for (file in excel_files) {
    # read the Excel file
    suppressMessages(
        df <- read_excel(file, col_names = FALSE)
    )
    # get study ID
    study_ID = df[3, 3][[1]]
    # get overall score
    overall_rating <- df[101,3][[1]]
    # get section ratings
    selection_bias = df[13, 3][[1]]
    data_collection_methods = df[23, 3][[1]]
    withdrawals_and_dropouts = df[33, 3][[1]]
    preprocessing = df[51, 3][[1]]
    trajectory_modelling = df[67, 3][[1]]
    associated_risk_factors_and_outcomes = df[81, 3][[1]]
    evaluation_and_reporting_of_results = df[97, 3][[1]]
    section_ratings = c(selection_bias, data_collection_methods, withdrawals_and_dropouts, preprocessing, trajectory_modelling, associated_risk_factors_and_outcomes, evaluation_and_reporting_of_results)
    # make a data.frame for the new row
    new_row <- data.frame(
        studyID = study_ID,
        selection_bias = selection_bias,
        data_collection_methods = data_collection_methods,
        withdrawals_and_dropouts = withdrawals_and_dropouts,
        preprocessing = preprocessing,
        trajectory_modelling = trajectory_modelling,
        associated_risk_factors_and_outcomes = associated_risk_factors_and_outcomes,
        evaluation_and_reporting_of_results = evaluation_and_reporting_of_results,
        overall_rating = overall_rating
    )
    # add row to the overall data.frame
    quality_assessment <- rbind(quality_assessment, new_row)
    # look for papers where the overall rating does not make sense
    weak_ratings = 0
    for (section_rating in section_ratings) {
        if (section_rating == "Weak") {
            weak_ratings = weak_ratings + 1
        }
    }
    # print addded data
    print(new_row)
    # print number of weak ratings and the overall coresponding rating
    if (overall_rating != 'SELECT OPTION') {
        print(paste0("Weak ratings: ", weak_ratings, ". Overall rating: ", overall_rating))
    } else {
        secondary_analyses = secondary_analyses + 1
        secondary_ratings <- c(secondary_ratings, associated_risk_factors_and_outcomes)
    }
    # print error in case of overall rating that does not make sense
    if (overall_rating != 'SELECT OPTION' & !(weak_ratings == 0 & overall_rating == "Strong") & !(weak_ratings == 1 & overall_rating == "Moderate") & !(weak_ratings > 1 & overall_rating == "Weak")) {
        print("INCORRECT OVERALL RATING")
        print("===============")
        quit()
    }
}
cat('-------- end\n')



cat('SUMMARIZE DATA\n')
cat('-------- start\n')
# present summary of results
n_complete <- nrow(quality_assessment) - secondary_analyses
# output the results
print('Quality assessment (quality_assessment):')
print(quality_assessment)
print('--------')
print('Overall ratings:')
# count the number of "Strong" in overall_rating column
strong_overall <- paste0("Strong: ", round(sum(quality_assessment$overall_rating == "Strong") / n_complete * 100, 1), "%")
print(strong_overall)
# count the number of "Moderate" in overall_rating column
moderate_overall <- paste0("Moderate: ", round(sum(quality_assessment$overall_rating == "Moderate") / n_complete * 100, 1), "%")
print(moderate_overall)
# count the number of "Weak" in overall_rating column
weak_overall <- paste0("Weak: ", round(sum(quality_assessment$overall_rating == "Weak") / n_complete * 100, 1), "%")
print(weak_overall)
print('--------')
print('Quality assessment by section:')
# for each column in quality_assessment starting from the second one, calculate the percentage of "Strong" ratings, "Moderate" ratings, and "Weak" ratings
section_rating_summaries <- c()
for (i in 2:ncol(quality_assessment)) {
    section_name <- names(quality_assessment)[i]
    if (section_name == "associated_risk_factors_and_outcomes") {
        denominator <- nrow(quality_assessment)
    } else {
        denominator <- n_complete
    }
    section_strong <- paste0("Strong: ", round(sum(quality_assessment[,i] == "Strong") / denominator * 100, 1), "%")
    section_moderate <- paste0("Moderate: ", round(sum(quality_assessment[,i] == "Moderate") / denominator * 100, 1), "%")
    section_weak <- paste0("Weak: ", round(sum(quality_assessment[,i] == "Weak") / denominator * 100, 1), "%")
    section_rating_summary <- paste0(section_name, ": ", section_strong, ", ", section_moderate, ", ", section_weak)
    print(section_rating_summary)
    section_rating_summaries <- c(section_rating_summaries, section_rating_summary)
    print('---')
}
print(secondary_ratings)
# for each rating ("Strong", "Moderate", "Weak", count the percentage in secondary_ratings
secondary_analysis_summary <- c()
for (rating in c("Strong", "Moderate", "Weak")) {
    rating_count <- sum(secondary_ratings == rating)
    rating_percentage <- paste0(rating, ": ", round(rating_count / length(secondary_ratings) * 100, 1), "%")
    secondary_analysis_summary <- c(secondary_analysis_summary, rating_percentage)
}
# save strong_overall, moderate_overall, weak_overall to one .txt file
write(c(strong_overall, moderate_overall, weak_overall), file = paste0(folder_path, "output/txt/quality-assessment-overall.txt"))
# save section_rating_summaries to one .txt file
write(section_rating_summaries, file = paste0(folder_path, "output/txt/quality-assessment-sections.txt"))
# save secondary_analysis_summary to one .txt file
write(secondary_analysis_summary, file = paste0(folder_path, "output/txt/quality-assessment-secondary.txt"))

# perform graphical base designs for visual presentations
# create a bar plot for ratings with ggplot2
stacked_barplot <- function(data, filename) {
    # if there are data elements of value "SELECT OPTION", remove these
    data <- data[data != "SELECT OPTION"]
    n_denominator <- length(data)
    data <- data.frame(
        rating = c("Strong", "Moderate", "Weak"),
        percentage = c(
            sum(data == "Strong") / n_denominator * 100,
            sum(data == "Moderate") / n_denominator * 100,
            sum(data == "Weak") / n_denominator * 100
        )
    )
    data$rating <- factor(data$rating, levels = c("Strong", "Moderate", "Weak"))
    plot <- ggplot(data, aes(x = "", y = percentage, fill = rating)) +
        geom_bar(stat = "identity") +
        coord_flip() +
        geom_text(aes(label = round(percentage,1)), position = position_stack(vjust = 0.5)) +
        scale_fill_manual(values = c("Strong" = "#88b591", "Moderate" = "#e4b97d", "Weak" = "#d1706f")) +
        theme(
            axis.line = element_blank(),
            axis.text = element_blank(),
            axis.ticks = element_blank(),
            panel.grid = element_blank(),
            panel.border = element_blank(),
            plot.background = element_blank(),
            panel.background = element_blank(),
            axis.title.x = element_blank(),
            axis.title.y = element_blank()
        ) +
        ggtitle(filename)
    ggsave(paste0(folder_path, "output/svg/", filename, ".svg"), plot = plot, device = "svg", dpi = 300, height = 0.84, width = 7)
}
print(table(quality_assessment$overall_rating))
stacked_barplot(quality_assessment$overall_rating, "quality-assessment-overall")

# create a bar plot for section ratings as per above
stacked_barplot(quality_assessment$selection_bias, "quality-assessment-selection-bias")
stacked_barplot(quality_assessment$data_collection_methods, "quality-assessment-data-collection-methods")
stacked_barplot(quality_assessment$withdrawals_and_dropouts, "quality-assessment-withdrawals-and-dropouts")
stacked_barplot(quality_assessment$preprocessing, "quality-assessment-preprocessing")
stacked_barplot(quality_assessment$trajectory_modelling, "quality-assessment-trajectory-modelling")
stacked_barplot(quality_assessment$associated_risk_factors_and_outcomes, "quality-assessment-associated-risk-factors-and-outcomes")
stacked_barplot(quality_assessment$evaluation_and_reporting_of_results, "quality-assessment-evaluation-and-reporting-of-results")
stacked_barplot(secondary_ratings, "quality-assessment-secondary")

cat('-------- end\n')



# make Word table with pooled quality assessments
cat('WRITING TO WORD\n')
cat('-------- start\n')
ft <- flextable::flextable(quality_assessment)
# Loop over all columns in the data frame
for (column_name in names(quality_assessment)) {
  # Color cells based on their values
  ft <- ft %>%
    flextable::bg(j = column_name, i = ~ get(column_name) == "Weak", bg = "#d1706f") %>%
    flextable::bg(j = column_name, i = ~ get(column_name) == "Moderate", bg = "#e4b97d") %>%
    flextable::bg(j = column_name, i = ~ get(column_name) == "Strong", bg = "#88b591") %>%
    flextable::bg(j = column_name, i = ~ get(column_name) == "Not applicable", bg = "#c9c4b2")
}
# set width
ft <- flextable::width(ft, width = .99)
# Change column names
ft <- flextable::set_header_labels(
    ft,
    values = list(
        studyID = "Study ID",
        selection_bias = "Selection bias",
        data_collection_methods = "Data collection methods",
        withdrawals_and_dropouts = "Withdrawals and dropouts",
        preprocessing = "Preprocessing",
        trajectory_modelling = "Trajectory modelling",
        associated_risk_factors_and_outcomes = "Associated risk factors and outcomes",
        evaluation_and_reporting_of_results = "Evaluation and reporting of results",
        overall_rating = "Overall rating"
    )
)
# Create a new Word document using the landscape template
doc <- read_docx(path = paste0(folder_path, "input/docx/templates/landscape.dotx"))
# Add the flextable to the Word document
doc <- flextable::body_add_flextable(doc, value = ft)
# Save the Word document
print(doc, target = paste0(folder_path, "output/docx/quality-assessment.docx"))
cat('-------- end\n')