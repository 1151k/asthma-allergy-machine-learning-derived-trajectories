# Specify the directory
dir_path <- "/Users/xlisda/Downloads/it/server/research/phd/taa_sr/input/pdf"

# Get a list of all files in the directory
files <- list.files(path = dir_path, full.names = TRUE)

# Loop over the files
for (file in files) {
  # Generate the new filename by removing "(number)"
  new_filename <- gsub(" \\(\\d+\\)", "", file)
  
  # Rename the file
  file.rename(from = file, to = new_filename)
}