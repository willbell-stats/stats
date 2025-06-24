# Load required packages
library(readxl)
library(dplyr)
library(openxlsx)  # For writing Excel files

# Define the file paths using current working directory
markers_file_path <- file.path(getwd(), "Marker_Data.xlsx")
pre_allocation_file_path <- file.path(getwd(), "Pre_Allocation_of_Markers.xlsx")
empty_school_files_directory <- getwd()  # Use current working directory

# Read the data
markers_data_frame <- read_excel(markers_file_path, col_names = FALSE,sheet=2)
pre_allocation_data_frame <- read_excel(pre_allocation_file_path, col_names = TRUE)

# Extract marker names from the markers_data_frame (assuming they're in the first column, skipping header)
marker_names <- markers_data_frame[[1]][-1]  # Remove first row (header)
marker_names <- marker_names[!is.na(marker_names)]  # Remove NA values
marker_names <- unique(marker_names)  # Get unique names


# Define the list of possible Marker 4s (can be expanded as needed)
possible_marker4s <- c("Mibeka Tan Panza")  # Add more names here as needed

# Combine all marker names (from data) with possible Marker 4s
all_folders_to_create <- unique(c(marker_names, possible_marker4s))

# Create a data frame for the folder names output
folder_names_output <- data.frame(`Folder Names` = all_folders_to_create)
names(folder_names_output) <- "Folder Names"  # Ensure proper column name with space

# Write the folder names to an Excel file
output_file_path <- file.path(getwd(), "Created_Folder_Names.xlsx")
write.xlsx(folder_names_output, output_file_path, rowNames = FALSE)

cat("Found markers:", paste(marker_names, collapse = ", "), "\n")
cat("Possible Marker 4s:", paste(possible_marker4s, collapse = ", "), "\n")
cat("Folder names list saved to:", output_file_path, "\n")

# Create folders for each marker and possible Marker 4
for (folder_name in all_folders_to_create) {
  # Create folder path
  folder_path <- file.path(empty_school_files_directory, folder_name)
  
  # Create the folder if it doesn't exist
  if (!dir.exists(folder_path)) {
    dir.create(folder_path, recursive = TRUE)
    cat("Created folder:", folder_path, "\n")
  } else {
    cat("Folder already exists:", folder_path, "\n")
  }
}

# Function to copy school files to marker folders
copy_school_files_to_markers <- function() {
  # Iterate through each row in pre_allocation_data_frame
  for (i in 1:nrow(pre_allocation_data_frame)) {
    school_serial <- pre_allocation_data_frame$`School Serial No.`[i]
    
    # Define the source file path (original format: FXXX - Empty.xlsx)
    source_file <- file.path(empty_school_files_directory, paste0(school_serial, " - Empty.xlsx"))
    
    # Check if the source file exists
    if (!file.exists(source_file)) {
      cat("Warning: File not found:", source_file, "\n")
      next
    }
    
    # Check each marker column and copy file to respective marker folders
    marker_columns <- c("Marker 1", "Marker 2", "Marker 3", "Marker 4")
    
    for (col in marker_columns) {
      marker_name <- pre_allocation_data_frame[[col]][i]
      
      # Skip if marker name is NA
      if (is.na(marker_name)) {
        next
      }
      
      # Define destination path (renamed format: FXXX.xlsx)
      marker_folder_path <- file.path(empty_school_files_directory, marker_name)
      destination_file <- file.path(marker_folder_path, paste0(school_serial, ".xlsx"))
      
      # Copy the file with new name
      if (file.exists(source_file)) {
        file.copy(source_file, destination_file, overwrite = TRUE)
        cat("Copied", paste0(school_serial, " - Empty.xlsx"), "to", marker_name, "folder as", paste0(school_serial, ".xlsx"), "\n")
      }
    }
  }
}

# Execute the file copying
copy_school_files_to_markers()

# Summary
cat("\n=== SUMMARY ===\n")
cat("Created folders for", length(all_folders_to_create), "markers and possible Marker 4s:\n")
for (folder in all_folders_to_create) {
  folder_path <- file.path(empty_school_files_directory, folder)
  if (dir.exists(folder_path)) {
    files_in_folder <- list.files(folder_path, pattern = "\\.xlsx$")
    cat("-", folder, ":", length(files_in_folder), "files\n")
  }
}

cat("\nFile organisation complete!\n")
cat("List of created folders saved to:", output_file_path, "\n")