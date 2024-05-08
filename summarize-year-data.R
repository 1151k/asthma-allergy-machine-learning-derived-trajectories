# Purpose: generate plot that shows the number of studies published each year from 2013 to 2023



# Load packages
packages_all = c("data.table", "readxl", "ggplot2")
packages_installed = packages_all %in% rownames(installed.packages())
if (any(packages_installed == FALSE)) {
    install.packages(packages_all[!packages_installed])
}
invisible(lapply(packages_all, library, character.only = TRUE))



# Load global constants
source('CONSTANTS.R')
# Load local constants



cat('READING QUALITY ASSESSMENT FILES\n')
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
studies_per_year <- list(
    '2013' = 0,
    '2014' = 0,
    '2015' = 0,
    '2016' = 0,
    '2017' = 0,
    '2018' = 0,
    '2019' = 0,
    '2020' = 0,
    '2021' = 0,
    '2022' = 0,
    '2023' = 0
)
high_rating_per_year <- list(
    '2013' = 0,
    '2014' = 0,
    '2015' = 0,
    '2016' = 0,
    '2017' = 0,
    '2018' = 0,
    '2019' = 0,
    '2020' = 0,
    '2021' = 0,
    '2022' = 0,
    '2023' = 0
)
moderate_rating_per_year <- list(
    '2013' = 0,
    '2014' = 0,
    '2015' = 0,
    '2016' = 0,
    '2017' = 0,
    '2018' = 0,
    '2019' = 0,
    '2020' = 0,
    '2021' = 0,
    '2022' = 0,
    '2023' = 0
)
low_rating_per_year <- list(
    '2013' = 0,
    '2014' = 0,
    '2015' = 0,
    '2016' = 0,
    '2017' = 0,
    '2018' = 0,
    '2019' = 0,
    '2020' = 0,
    '2021' = 0,
    '2022' = 0,
    '2023' = 0
)
# Loop through each Excel file
for (file in excel_files) {
    # Read the Excel file
    suppressMessages(
        df <- read_excel(file, col_names = FALSE)
    )
    # get data items
    study_ID = df[3, 3][[1]]
    year = stringr::str_extract_all(study_ID, "\\d+")[[1]]
    studies_per_year[[year]] <- studies_per_year[[year]] + 1
    overall_rating <- df[101,3][[1]]
    if (overall_rating != "SELECT OPTION") {
        if (overall_rating == "Weak") {
            low_rating_per_year[[year]] <- low_rating_per_year[[year]] + 1
        } else if (overall_rating == "Moderate") {
            moderate_rating_per_year[[year]] <- moderate_rating_per_year[[year]] + 1
        } else if (overall_rating == "Strong") {
            high_rating_per_year[[year]] <- high_rating_per_year[[year]] + 1
        }
    }
    print(paste('Done with ', study_ID, sep = ''))
}
print('Done with all files')
cat('-------- end\n')



cat('PLOT AND SAVE DATA\n')
cat('-------- start\n')
# convert list to data.table
studies_per_year <- data.table(
    year = as.numeric(names(studies_per_year)),
    n_studies = as.numeric(unlist(studies_per_year))
)
low_rating_per_year <- data.table(
    year = as.numeric(names(low_rating_per_year)),
    n_studies = as.numeric(unlist(low_rating_per_year))
)
moderate_rating_per_year <- data.table(
    year = as.numeric(names(moderate_rating_per_year)),
    n_studies = as.numeric(unlist(moderate_rating_per_year))
)
high_rating_per_year <- data.table(
    year = as.numeric(names(high_rating_per_year)),
    n_studies = as.numeric(unlist(high_rating_per_year))
)
print(studies_per_year)
print(low_rating_per_year)
print(moderate_rating_per_year)
print(high_rating_per_year)
# plot the data with ggplot2
plot <- ggplot2::ggplot(data = studies_per_year) +
    geom_bar(aes(x = year, y = n_studies), stat = "identity", fill = "#e3e3e3") +
    geom_line(data = low_rating_per_year, aes(x = year, y = n_studies), color = "#d1706f", linewidth = 1.5) +
    geom_line(data = moderate_rating_per_year, aes(x = year, y = n_studies), color = "#E8ba42", linewidth = 1.5) +
    geom_line(data = high_rating_per_year, aes(x = year, y = n_studies), color = "#88b591", linewidth = 1.5) +
    labs(x = "Year", y = "Number of studies") +
    theme(
        panel.background = element_rect(fill = "#ffffff"),
        axis.line = element_line(color='black', linewidth = 0.7),
        axis.title.x = element_text(size = 22, color = "black", margin = margin(t = 18)),
        axis.title.y = element_text(size = 22, color = "black", margin = margin(r = 13)),
        axis.text.x = element_text(size = 19, color = "black", margin = margin(t = 13)),
        axis.text.y = element_text(size = 19, color = "black", margin = margin(r = 13)),
        axis.ticks = element_line(linewidth=0.7), 
        axis.ticks.length = unit(0.36, "cm")
    ) +
    scale_x_continuous(breaks = seq(from = 2013, to = 2023, by = 2), limits = c(2012, 2024), expand = c(0, 0)) +
    scale_y_continuous(breaks = seq(from = 0, to = 15, by = 5), limits = c(0, 16), expand = c(0, 0))
ggsave(paste0('output/svg/', 'studies-per-year.svg'), plot, dpi = 300, width = 7, height = 4.5)

cat('-------- end\n')