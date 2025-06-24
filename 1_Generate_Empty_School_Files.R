# Load required packages
library(readxl)
library(writexl)

# Get the current working directory
project_dir <- getwd()

# Define the file paths relative to the working directory
school_template_file_path <- file.path(project_dir, "P18745 FXXX Marked File.xlsx")
schedule_file_path <- file.path(project_dir, "School_Test_Delivery_Schedule.xlsx")

# Read the template file (both sheets)
school_template_sheet1 <- read_excel(school_template_file_path, sheet = 1)
datamap_sheet <- read_excel(school_template_file_path, sheet = "Datamap")

# Get the name of the first sheet
template_sheet_names <- excel_sheets(school_template_file_path)
first_sheet_name <- template_sheet_names[1]

# Read schedule data
schedule_data_frame <- read_excel(schedule_file_path, col_names = TRUE)

# Get the list of school serial numbers
school_serial_numbers <- schedule_data_frame$`School Serial No.`

# Function to create data for each school
create_school_data <- function(school_serial) {
  # Create 30 rows of data for this school
  school_data <- school_template_sheet1[1:30, ]  # Create 30 empty rows with same structure
  
  # Fill in the School Serial No. column
  school_data$`School Serial No.` <- school_serial
  
  # Generate Pupil Serial Numbers (FXXX001 through FXXX030)
  pupil_serials <- sprintf("%s%03d", school_serial, 1:30)
  school_data$`Pupil Serial No.` <- pupil_serials
  
  return(school_data)
}

# Create files for each school
for (school_serial in school_serial_numbers) {
  # Create the data for this school
  school_data <- create_school_data(school_serial)
  
  # Create the workbook with both sheets
  # First sheet named after the school serial number
  # Second sheet is the unchanged Datamap
  workbook_data <- list()
  workbook_data[[paste0("School ", school_serial)]] <- school_data
  workbook_data[["Datamap"]] <- datamap_sheet
  
  # Define the output file path (in the project directory)
  output_file_path <- file.path(project_dir, paste0(school_serial, " - Empty.xlsx"))
  
  # Write the Excel file with both sheets
  write_xlsx(workbook_data, output_file_path)
  
  # Print progress message
  cat("Created file:", output_file_path, "with sheets: School", school_serial, "and Datamap\n")
}

cat("All school files have been created successfully!\n")
cat("Total files created:", length(school_serial_numbers), "\n")
cat("Each file contains two sheets:\n")
cat("- Sheet 1: 'School [SerialNo]' with 30 pupils\n")
cat("- Sheet 2: 'Datamap' (unchanged from template)\n")