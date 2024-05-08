# Load packages
packages_all = c("data.table", "readxl", "ggplot2", "dplyr", "stringr", "gridExtra")
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
n_assessments_data <- data.table(
    x = numeric(),
    y = numeric()
)
n_assessments_values <- c()
threshold_assessments_data <- data.table(
    x = numeric(),
    y = numeric()
)
threshold_diseases_data <- data.table(
    x = numeric(),
    y = numeric()
)
years_follow_up_data <- data.table(
    x = numeric(), # years of follow-up
    y = numeric() # number of trajectories identified
)
# Loop through each Excel file
for (file in excel_files) {
    # Read the Excel file
    suppressMessages(
        df <- read_excel(file, col_names = FALSE)
    )

    # get data items
    study_ID = df[3, 3][[1]]
    if (study_ID %in% primary_study_list) {
        print(study_ID)
        n_trajectories = strsplit(as.character(df[69, 4][[1]]), ",")[[1]]
        n_assessments = trimws(strsplit(df[37, 5][[1]], ",")[[1]])
        n_assessments = max(as.numeric(n_assessments))
        years_followed = as.numeric(df[19, 4][[1]])
        print(paste('years followed:', years_followed))
        n_assessments_values <- c(n_assessments_values, n_assessments)
        n_trajectories_max <- max(as.numeric(n_trajectories))
        years_follow_up_data <- rbind(
            years_follow_up_data,
            data.table(
                x = years_followed,
                y = n_trajectories_max
            )
        )
        # make three groups of number of assessments
        if (n_assessments < 5) {
            threshold_assessments <- 1
        } else if (n_assessments < 10) {
            threshold_assessments <- 2
        } else if (n_assessments >= 10) {
            threshold_assessments <- 3
        }
        # number of diseases (set to ≥2 if there are multiple diseases)
        n_diseases = df[69, 5][[1]]
        if (grepl(",", n_diseases)) {
            n_diseases <- 2
        } else {
            n_diseases <- 1
        }
        # loop each number of trajectories
        for (ng in n_trajectories) {
            n_assessments_data <- rbind(
                n_assessments_data,
                data.table(
                    x = as.numeric(ng),
                    y = n_assessments
                )
            )
            threshold_assessments_data <- rbind(
                threshold_assessments_data,
                data.table(
                    x = as.numeric(ng),
                    y = threshold_assessments
                )
            )
            threshold_diseases_data <- rbind(
                threshold_diseases_data,
                data.table(
                    x = as.numeric(ng),
                    y = n_diseases
                )
            )
        }
        print('----')
    }
}
print('Done with all files')
cat('-------- end\n')



cat('SAVING DATA EXTRACTION\n')
cat('-------- start\n')
# get median of the column y in n_assessments_data
n_assessments_median <- median(n_assessments_values)
n_assessments_min <- min(n_assessments_values)
n_assessments_max <- max(n_assessments_values)
# save these three above values to a .txt file
write(
    paste("Median: ", n_assessments_median, "\nMin: ", n_assessments_min, "\nMax: ", n_assessments_max),
    file = paste0(folder_path, "output/txt/number-of-assessments.txt")
)
# make histogram of the number of assessments (y in n_assessments_data)
p <- ggplot(n_assessments_data, aes(x = y)) +
    geom_bar(width = 0.84, fill = "#e3e3e3") +
    labs(
        x = "Number of assessments",
        y = "Number of studies"
    ) +
    theme(
        panel.background = element_rect(fill = "#ffffff"),
        axis.line = element_line(color='black', linewidth = 0.7),
        axis.title.x = element_text(size = 19, color = "black", margin = margin(t = 9)),
        axis.title.y = element_text(size = 19, color = "black", margin = margin(r = 9)),
        axis.text.x = element_text(size = 16, color = "black", margin = margin(t = 13)),
        axis.text.y = element_text(size = 16, color = "black", margin = margin(r = 13)),
        axis.ticks = element_line(linewidth=0.7), 
        axis.ticks.length = unit(0.3, "cm")
    ) +
    scale_x_continuous(breaks = seq(from = 2, to = 25, by = 4), limits = c(0, 25), expand = c(0, 0)) +
    scale_y_continuous(breaks = seq(from = 0, to = 13, by = 4), limits = c(0, 13), expand = c(0, 0))
ggsave(
    paste0(folder_path, "output/svg/number-of-assessments.svg"),
    plot = p,
    device = "svg",
    dpi = 300,
    width = 10,
    height = 6
)
# make stacked bar plot of the number of assessments (y is in threshold_assessments_data)
p_assessments <- ggplot(threshold_assessments_data, aes(x = x, fill = factor(y))) +
    geom_bar(position = "stack", stat = "count") +
    labs(
        x = "Number of trajectories",
        y = "Number of studies",
        fill = "Number of assessments"
    ) +
    theme(
        panel.background = element_rect(fill = "#ffffff"),
        axis.line = element_line(color='black', linewidth = 0.7),
        axis.title.x = element_text(size = 22, color = "black", margin = margin(t = 18)),
        axis.title.y = element_text(size = 22, color = "black", margin = margin(r = 13)),
        axis.text.x = element_text(size = 19, color = "black", margin = margin(t = 13)),
        axis.text.y = element_text(size = 19, color = "black", margin = margin(r = 13)),
        axis.ticks = element_line(linewidth=0.7), 
        axis.ticks.length = unit(0.36, "cm"),
        legend.position = "top",
        legend.title = element_text(size = 19, color = "black"),
        legend.text = element_text(size = 16, color = "black")
    ) +
    scale_x_continuous(breaks = seq(from = 2, to = 8, by = 2), limits = c(1, 9), expand = c(0, 0)) +
    scale_y_continuous(breaks = seq(from = 0, to = 28, by = 5), limits = c(0, 29), expand = c(0, 0)) +
    scale_fill_manual(
        values = c("#e3e3e3", "#A6AAC9", "#585E92"),
        labels = c("1-4", "5-9", "≥10")
    )
# make stacked bar plot of the number of diseases (y is in threshold_diseases_data)
p_diseases <- ggplot(threshold_diseases_data, aes(x = x, fill = factor(y))) +
    geom_bar(position = "stack", stat = "count") +
    labs(
        x = "Number of trajectories",
        y = "Number of studies",
        fill = "Number of diseases"
    ) +
    theme(
        panel.background = element_rect(fill = "#ffffff"),
        axis.line = element_line(color='black', linewidth = 0.7),
        axis.title.x = element_text(size = 22, color = "black", margin = margin(t = 18)),
        axis.title.y = element_text(size = 22, color = "black", margin = margin(r = 13)),
        axis.text.x = element_text(size = 19, color = "black", margin = margin(t = 13)),
        axis.text.y = element_text(size = 19, color = "black", margin = margin(r = 13)),
        axis.ticks = element_line(linewidth=0.7), 
        axis.ticks.length = unit(0.36, "cm"),
        legend.position = "top",
        legend.title = element_text(size = 19, color = "black"),
        legend.text = element_text(size = 16, color = "black")
    ) +
    scale_x_continuous(breaks = seq(from = 2, to = 8, by = 2), limits = c(1, 9), expand = c(0, 0)) +
    scale_y_continuous(breaks = seq(from = 0, to = 28, by = 5), limits = c(0, 29), expand = c(0, 0)) +
    scale_fill_manual(
        values = c("#e3e3e3", "#585E92"),
        labels = c("1", "≥2")
    )
# merge p_assessments and p_diseases with gridExtra side by side
p_combined <- gridExtra::grid.arrange(
    p_assessments,
    p_diseases,
    ncol = 2
)
ggsave(
    paste0(folder_path, "output/svg/number-of-assessments-diseases-threshold.svg"),
    plot = p_combined,
    device = "svg",
    dpi = 300,
    width = 18,
    height = 6.5
)
# make scatter plot of years of follow-up (y in years_follow_up_data)
p_years_follow_up <- ggplot(years_follow_up_data, aes(x = x, y = y)) +
    geom_point(size = 2, color = '#595e92') +
    labs(
        x = "Years of follow-up",
        y = "Number of trajectories"
    ) +
    theme(
        panel.background = element_rect(fill = "#ffffff"),
        axis.line = element_line(color='black', linewidth = 0.7),
        axis.title.x = element_text(size = 19, color = "black", margin = margin(t = 9)),
        axis.title.y = element_text(size = 19, color = "black", margin = margin(r = 9)),
        axis.text.x = element_text(size = 16, color = "black", margin = margin(t = 13)),
        axis.text.y = element_text(size = 16, color = "black", margin = margin(r = 13)),
        axis.ticks = element_line(linewidth=0.7), 
        axis.ticks.length = unit(0.3, "cm")
    ) +
    scale_x_continuous(breaks = seq(from = 0, to = 18.5, by = 4), limits = c(0, 18.5), expand = c(0, 0)) +
    scale_y_continuous(breaks = seq(from = 0, to = 9, by = 2), limits = c(0, 9), expand = c(0, 0))
ggsave(
    paste0(folder_path, "output/svg/years-of-follow-up.svg"),
    plot = p_years_follow_up,
    device = "svg",
    dpi = 300,
    width = 10,
    height = 6
)
# merge p_years_follow_up and p with gridExtra side by side
p_combined <- gridExtra::grid.arrange(
    p,
    p_years_follow_up,
    ncol = 2
)
ggsave(
    paste0(folder_path, "output/svg/number-of-assessments-years-follow-up.svg"),
    plot = p_combined,
    device = "svg",
    dpi = 300,
    width = 18,
    height = 6.5
)
cat('-------- end\n')