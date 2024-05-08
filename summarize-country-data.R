# Load packages
packages_all = c("data.table", "readxl", "reshape2")
packages_installed = packages_all %in% rownames(installed.packages())
if (any(packages_installed == FALSE)) {
    install.packages(packages_all[!packages_installed])
}
invisible(lapply(packages_all, library, character.only = TRUE))



# Load global constants
source('CONSTANTS.R')
# Load local constants



cat("READING PRIMARY STUDY LIST\n")
cat('-------- start\n')
# Read the primary study list (the file is in .txt format)
primary_study_list <- readLines(paste0(folder_path, "output/txt/primary-study-ids.txt"))
primary_study_list <- unlist(strsplit(primary_study_list, ","))
print(primary_study_list)
# to get the secondary analyses on studies prior to 2013
primary_study_list <- c("Caudri 2013", "Collin 2015", "Kotsapas 2022")
cat('-------- end\n')



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
country_num <- data.table(
    "United Kingdom" = 0
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
        countries_all <- df[11, 3][[1]]
        # split by "\r\n"
        countries_all <- unlist(strsplit(countries_all, "\r\n"))
        # in each item, delete everything before ":"
        countries_all <- gsub(".*:", "", countries_all)
        # trim all whitespace from the beginning and end of each item
        countries_all <- trimws(countries_all)
        # loop each value in country, if it contains ",", split by ","
        for (countries_slice in countries_all) {
            countries_slice <- unlist(strsplit(countries_slice, ","))
            for (country in countries_slice) {
                country <- trimws(country)
                print(country)
                if (country %in% colnames(country_num)) {
                    # if it is, increment the value in the corresponding cell
                    country_num[1, (country) := country_num[1, get(country) + 1]]
                } else {
                    # if it is not, add a new column with the value 1
                    country_num[1, (country) := 1]
                }
            }
        }
        print(paste('---- Done with ', study_ID, sep = ''))
    }
    ## study ID and country
}
print('Done with all files')
cat('-------- end\n')



cat('SAVING DATA\n')
cat('-------- start\n')
long_format <- reshape2::melt(country_num, measure.vars = names(country_num))
setorder(long_format, -value)
print(head(long_format))
# add a row at the beginning with the total
long_format <- rbind(data.table(variable = "Total", value = dim(long_format)[1]), long_format)
# save to .txt file without the column names, but add a row at the end with the total
write.table(long_format, paste0(folder_path, "output/txt/country-data.txt"), col.names = FALSE, row.names = FALSE, quote = FALSE, sep = "\t")
cat('-------- end\n')