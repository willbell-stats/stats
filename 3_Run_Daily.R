# This is the code that should be run daily. 
# It carries out several essential tasks in the work specification for the project.
# Specifically, it:
# 1.) Cleans up the directories, removing or adding empty versions of the school files if changes are made to the pre-allocation spreadsheet.
# 2.) Identifies third or fourth markers if needed.
# 3.) Emails the third marker to notify them that they have been selected.
# 4.) Assigns the third/fourth marker an appropriate copy of the school sheet to complete, with relevant cells highlighted and the rest locked.
# 5.) Locks any completed school files to prevent markers retroactively changing them - this may otherwise lead to inconsistencies in the final output.
# 6.) Puts the locked, completed school files into an archive folder once the marks for all questions and all pupils have been finalised.
# 7.) Produces a checks dashboard for the day.
# NB: Emailing is done through an address I created called eefmarking@gmail.com. The password to access this account is 123Alpha.Plus. 
# The account also has a 16-character app password, which is the one required by mailR: dyhogkbcqzrjxcau.
# The emails are sent out every time a third marker is NEWLY appointed and their names are recorded in an updated version of the pre-allocation spreadsheet.
# Probably not needed, but I have left the option for you to have more than one Marker 4 in the pool at Alpha Plus. 
# If more than one person wishes to be considered for fourth marker, add their name to the appropriate lists and this code will select a name at random from the list. 
# This is useful if more than one fourth marker is involved from Alpha Plus and they each want their own folder.
# Markers should be advised before the marking period not to change the files they are given in any way other than inputting the marks into the relevant cells.
# This code works under the assumption that we have exactly 30 pupils from each school, as stated in the work specification.

#############
# LIBRARIES #
#############

library(readxl)
library(writexl)
library(fs)
library(purrr)
library(dplyr)
library(openxlsx)
library(mailR) 
library(readxl)

#############################################################################
# CODE TO CHECK MARKERS' FOLDERS AND THEIR CONTENTS VS. PRE-ALLOCATION FILE #
#############################################################################

# Define paths and file names
project_dir <- getwd()
pre_allocation_file <- file.path(project_dir, "Pre_Allocation_of_Markers.xlsx")
folder_names_file <- file.path(project_dir, "Created_Folder_Names.xlsx")

# Read the data
pre_allocation <- read_excel(pre_allocation_file)
folder_names <- read_excel(folder_names_file)

# Get list of all marker names from the folder names
all_markers <- folder_names$`Folder Names`

# Function to process each school serial number
process_serial <- function(serial) {
  # Get markers allocated to this serial
  markers <- pre_allocation %>% 
    filter(`School Serial No.` == serial) %>% 
    select(`Marker 1`, `Marker 2`, `Marker 3`, `Marker 4`) %>% 
    unlist() %>% 
    na.omit() %>% 
    as.character()
  
  # File names
  target_file <- paste0(serial, ".xlsx")
  empty_template_file <- file.path(project_dir, paste0(serial, " - Empty.xlsx"))
  
  # Process each marker's folder
  for (marker in all_markers) {
    marker_folder <- file.path(project_dir, marker)
    
    if (dir_exists(marker_folder)) {
      file_path <- file.path(marker_folder, target_file)
      
      if (marker %in% markers) {
        # Marker is allocated - ensure file exists
        if (!file_exists(file_path)) {
          # Check if the specific empty template exists
          if (file_exists(empty_template_file)) {
            # Copy the empty template to the correct name
            file_copy(empty_template_file, file_path)
            message("Created ", target_file, " in ", marker, "'s folder using ", basename(empty_template_file))
          } else {
            warning("Empty template not found for serial ", serial, " (looking for: ", basename(empty_template_file), ")")
          }
        }
      } else {
        # Marker is not allocated - delete file if exists
        if (file_exists(file_path)) {
          file_delete(file_path)
          message("Deleted ", target_file, " from ", marker, "'s folder (no longer allocated)")
        }
      }
    }
  }
}

# Process all serial numbers
walk(pre_allocation$`School Serial No.`, process_serial)

message("Folder and file verification complete!")

#################################################################################
# CODE TO PRODUCE THE FIRST TWO SHEETS OF THE 'CHECKS DASHBOARD' EXCEL WORKBOOK #
#################################################################################

# Define question columns (Q1a to Q15)
question_cols <- c("Q1a", "Q1b", "Q1c", "Q2a", "Q2b", "Q2c", "Q3", "Q4a", "Q4b", "Q4c", "Q4d", 
                   "Q5a", "Q5b", "Q5c", "Q5d", "Q5e", "Q6a", "Q6b", "Q7a", "Q7b", "Q7c", 
                   "Q8a", "Q8b", "Q9a", "Q9b", "Q9c", "Q10a", "Q10b", "Q11a", "Q11b", "Q11c", 
                   "Q12", "Q13", "Q14a", "Q14b", "Q14c", "Q14d", "Q15")

# Function to check if a paper (row) is complete (no blank cells in question columns)
is_paper_complete <- function(paper_row) {
  q_values <- paper_row[question_cols]
  all(!is.na(q_values) & q_values != "")
}

# Function to check if two markers agree on all questions for a pupil's paper
markers_agree <- function(marker1_data, marker2_data, pupil_idx) {
  for (q_col in question_cols) {
    mark1 <- marker1_data[pupil_idx, q_col][[1]]
    mark2 <- marker2_data[pupil_idx, q_col][[1]]
    
    # Skip if either mark is NA or empty
    if (is.na(mark1) || is.na(mark2) || mark1 == "" || mark2 == "") {
      return(FALSE)  # If either is blank, they don't agree
    }
    
    if (mark1 != mark2) {
      return(FALSE)
    }
  }
  return(TRUE)
}

# Function to get disagreement positions between marker 1 and 2
get_disagreements <- function(marker1_data, marker2_data, pupil_idx) {
  disagreements <- c()
  for (q_col in question_cols) {
    mark1 <- marker1_data[pupil_idx, q_col][[1]]
    mark2 <- marker2_data[pupil_idx, q_col][[1]]
    
    # Skip if either mark is NA or empty
    if (is.na(mark1) || is.na(mark2) || mark1 == "" || mark2 == "") next
    
    if (mark1 != mark2) {
      disagreements <- c(disagreements, q_col)
    }
  }
  return(disagreements)
}

# Function to check if all disagreement cells are filled by marker 3/4
all_disagreement_cells_filled <- function(marker_data, disagreement_cols, pupil_idx) {
  if (length(disagreement_cols) == 0) return(TRUE)
  
  for (q_col in disagreement_cols) {
    mark <- marker_data[pupil_idx, q_col][[1]]
    if (is.na(mark) || mark == "") {
      return(FALSE)
    }
  }
  return(TRUE)
}

# Function to check if marker 3 resolves all disagreements properly
marker3_resolves_disagreements <- function(marker1_data, marker2_data, marker3_data, pupil_idx, disagreement_cols) {
  if (length(disagreement_cols) == 0) return(TRUE)
  
  for (q_col in disagreement_cols) {
    mark1 <- marker1_data[pupil_idx, q_col][[1]]
    mark2 <- marker2_data[pupil_idx, q_col][[1]]
    mark3 <- marker3_data[pupil_idx, q_col][[1]]
    
    # Marker 3 must have filled this cell
    if (is.na(mark3) || mark3 == "") {
      return(FALSE)
    }
    
    # Marker 3 must agree with at least one of marker 1 or 2
    if (mark3 != mark1 && mark3 != mark2) {
      return(FALSE)
    }
  }
  return(TRUE)
}

# Function to count papers with disagreements (allocated to Marker 3)
count_papers_with_disagreements <- function(serial, pre_allocation_row) {
  marker1 <- pre_allocation_row$`Marker 1`
  marker2 <- pre_allocation_row$`Marker 2`
  
  target_file <- paste0(serial, ".xlsx")
  marker1_path <- file.path(project_dir, marker1, target_file)
  marker2_path <- file.path(project_dir, marker2, target_file)
  
  # Check if marker files exist
  if (!file_exists(marker1_path) || !file_exists(marker2_path)) {
    return(0)
  }
  
  marker1_data <- read_excel(marker1_path)
  marker2_data <- read_excel(marker2_path)
  
  # Count papers with disagreements
  papers_with_disagreements <- 0
  for (pupil_idx in 1:min(nrow(marker1_data), nrow(marker2_data))) {
    disagreements <- get_disagreements(marker1_data, marker2_data, pupil_idx)
    if (length(disagreements) > 0) {
      papers_with_disagreements <- papers_with_disagreements + 1
    }
  }
  
  return(papers_with_disagreements)
}

# Function to count papers completed by Marker 3
count_marker3_completed_papers <- function(serial, pre_allocation_row) {
  marker1 <- pre_allocation_row$`Marker 1`
  marker2 <- pre_allocation_row$`Marker 2`
  marker3 <- pre_allocation_row$`Marker 3`
  
  # If no Marker 3 assigned, return 0
  if (is.na(marker3) || marker3 == "") {
    return(0)
  }
  
  target_file <- paste0(serial, ".xlsx")
  marker1_path <- file.path(project_dir, marker1, target_file)
  marker2_path <- file.path(project_dir, marker2, target_file)
  marker3_path <- file.path(project_dir, marker3, target_file)
  
  # Check if all required files exist
  if (!file_exists(marker1_path) || !file_exists(marker2_path) || !file_exists(marker3_path)) {
    return(0)
  }
  
  marker1_data <- read_excel(marker1_path)
  marker2_data <- read_excel(marker2_path)
  marker3_data <- read_excel(marker3_path)
  
  # Check if files have consistent structure
  if (nrow(marker1_data) < 1 || nrow(marker2_data) < 1 || nrow(marker3_data) < 1) {
    return(0)
  }
  
  # Count completed papers (where Marker 3 has resolved all disagreements)
  completed_papers <- 0
  for (pupil_idx in 1:min(nrow(marker1_data), nrow(marker2_data), nrow(marker3_data))) {
    disagreements <- get_disagreements(marker1_data, marker2_data, pupil_idx)
    if (length(disagreements) > 0) {
      # Check if Marker 3 has filled all disagreement cells
      if (all_disagreement_cells_filled(marker3_data, disagreements, pupil_idx)) {
        completed_papers <- completed_papers + 1
      }
    }
  }
  
  return(completed_papers)
}

# Function to check if Marker 3 has completed all work for a school
is_marker3_school_complete <- function(serial, pre_allocation_row) {
  marker3 <- pre_allocation_row$`Marker 3`
  
  # If no Marker 3 assigned, return TRUE (not applicable)
  if (is.na(marker3) || marker3 == "") {
    return(TRUE)
  }
  
  allocated_papers <- count_papers_with_disagreements(serial, pre_allocation_row)
  completed_papers <- count_marker3_completed_papers(serial, pre_allocation_row)
  
  return(allocated_papers == completed_papers)
}

# Main function to determine if a paper is "marked" according to the criteria
is_paper_marked <- function(serial, pupil_idx, pre_allocation_row) {
  marker1 <- pre_allocation_row$`Marker 1`
  marker2 <- pre_allocation_row$`Marker 2`  
  marker3 <- pre_allocation_row$`Marker 3`
  marker4 <- pre_allocation_row$`Marker 4`
  
  # File paths
  target_file <- paste0(serial, ".xlsx")
  marker1_path <- file.path(project_dir, marker1, target_file)
  marker2_path <- file.path(project_dir, marker2, target_file)
  
  # Check if marker 1 and 2 files exist
  if (!file_exists(marker1_path) || !file_exists(marker2_path)) {
    return(FALSE)
  }
  
  # Read marker 1 and 2 data
  marker1_data <- read_excel(marker1_path)
  marker2_data <- read_excel(marker2_path)
  
  # Verify files have same structure and pupil exists
  if (nrow(marker1_data) < pupil_idx || nrow(marker2_data) < pupil_idx ||
      marker1_data$`Pupil Serial No.`[pupil_idx] != marker2_data$`Pupil Serial No.`[pupil_idx]) {
    return(FALSE)
  }
  
  # Criterion 1: For Markers 1 and 2, none of the cells for Q1a-15 for that row/pupil/paper is blank
  if (!is_paper_complete(marker1_data[pupil_idx,]) || !is_paper_complete(marker2_data[pupil_idx,])) {
    return(FALSE)
  }
  
  # Get disagreements between marker 1 and 2
  disagreements <- get_disagreements(marker1_data, marker2_data, pupil_idx)
  
  # Check if marker 3 is assigned
  if (is.na(marker3) || marker3 == "" || is.na(marker3)) {
    # Criterion 3: If no Marker 3 has yet been assigned, there are no discrepancies at all 
    # on this row/pupil/paper between Marker 1 and Marker 2
    return(length(disagreements) == 0)
  }
  
  # Marker 3 is assigned - check if marker 3 file exists
  marker3_path <- file.path(project_dir, marker3, target_file)
  if (!file_exists(marker3_path)) {
    return(FALSE)  # Marker 3 file doesn't exist yet
  }
  
  marker3_data <- read_excel(marker3_path)
  if (nrow(marker3_data) < pupil_idx) {
    return(FALSE)
  }
  
  # Check if marker 4 is assigned
  if (is.na(marker4) || marker4 == "" || is.na(marker4)) {
    # Criterion 4: If Marker 3 has been assigned, but marker 4 has not yet been assigned, 
    # there are no discrepancies between Marker 3 and both Marker 1 and Marker 2
    
    # First check that all disagreement cells are filled by marker 3
    if (!all_disagreement_cells_filled(marker3_data, disagreements, pupil_idx)) {
      return(FALSE)
    }
    
    # Then check that marker 3 resolves disagreements properly
    return(marker3_resolves_disagreements(marker1_data, marker2_data, marker3_data, pupil_idx, disagreements))
  }
  
  # Marker 4 is assigned - check if marker 4 file exists
  marker4_path <- file.path(project_dir, marker4, target_file)
  if (!file_exists(marker4_path)) {
    return(FALSE)  # Marker 4 file doesn't exist yet
  }
  
  marker4_data <- read_excel(marker4_path)
  if (nrow(marker4_data) < pupil_idx) {
    return(FALSE)
  }
  
  # Criterion 2: For Marker 4, ALL cells for Q1a-15 for that row/pupil/paper are NOT yellow, 
  # OR if there are some yellow cells for this row/pupil/paper, they are all not blank
  # Yellow cells are those where marker 3 disagreed with both 1&2
  # For now, we check that all disagreement positions from original markers 1&2 are filled by marker 4
  return(all_disagreement_cells_filled(marker4_data, disagreements, pupil_idx))
}

# Initialise progress tracking
overall_progress <- data.frame(
  `School Serial No.` = pre_allocation$`School Serial No.`,
  `Expected Papers, n` = 30,
  `Papers Marked, n` = 0,
  `Papers Marked, %` = 0,
  check.names = FALSE
)

# Create marker progress for Marker 1, Marker 2, AND Marker 3 (excluding external marker)
# Get unique markers from Marker 1, Marker 2, and Marker 3 columns
markers_1_2_and_3 <- unique(c(
  pre_allocation$`Marker 1`[!is.na(pre_allocation$`Marker 1`)],
  pre_allocation$`Marker 2`[!is.na(pre_allocation$`Marker 2`)],
  pre_allocation$`Marker 3`[!is.na(pre_allocation$`Marker 3`)]
))

marker_progress <- data.frame(
  `Marker Name` = markers_1_2_and_3,
  `Allocated Schools, n` = 0,
  `Schools Complete, n` = 0,
  `Schools Complete, %` = 0,
  `Allocated Papers, n` = 0,
  `Papers Complete, n` = 0,
  `Papers Complete, %` = 0,
  check.names = FALSE
)

# Track which schools are complete for which markers
marker_school_completion <- list()

# Initialise completion tracking
for (marker in markers_1_2_and_3) {
  marker_school_completion[[marker]] <- list()
}

# Analyse each school
for (i in 1:nrow(pre_allocation)) {
  serial <- pre_allocation$`School Serial No.`[i]
  school_row <- pre_allocation[i, ]
  
  # Get Marker 1, Marker 2, and Marker 3 for this school
  marker1 <- school_row$`Marker 1`
  marker2 <- school_row$`Marker 2`
  marker3 <- school_row$`Marker 3`
  
  # Skip if marker is the external marker or missing
  valid_markers <- c()
  if (!is.na(marker1)) {
    valid_markers <- c(valid_markers, marker1)
  }
  if (!is.na(marker2)) {
    valid_markers <- c(valid_markers, marker2)
  }
  if (!is.na(marker3)) {
    valid_markers <- c(valid_markers, marker3)
  }
  
  if (length(valid_markers) == 0) {
    next
  }
  
  target_file <- paste0(serial, ".xlsx")
  
  # Update allocated schools count for all valid markers
  for (marker in valid_markers) {
    marker_progress$`Allocated Schools, n`[marker_progress$`Marker Name` == marker] <- 
      marker_progress$`Allocated Schools, n`[marker_progress$`Marker Name` == marker] + 1
  }
  
  # Check marker files
  marker1_path <- file.path(project_dir, marker1, target_file)
  marker2_path <- file.path(project_dir, marker2, target_file)
  marker3_path <- if (!is.na(marker3)) {
    file.path(project_dir, marker3, target_file)
  } else {
    NULL
  }
  
  # Get number of papers
  num_papers <- 30  # Default expected
  if (file_exists(marker1_path)) {
    marker1_data <- read_excel(marker1_path)
    num_papers <- nrow(marker1_data)
  }
  
  # Check Marker 1's completion
  if (!is.na(marker1)) {
    if (file_exists(marker1_path)) {
      marker1_data <- read_excel(marker1_path)
      papers_complete <- sum(sapply(1:nrow(marker1_data), function(i) is_paper_complete(marker1_data[i,])))
      
      # Update allocated and completed papers count
      marker_progress$`Allocated Papers, n`[marker_progress$`Marker Name` == marker1] <- 
        marker_progress$`Allocated Papers, n`[marker_progress$`Marker Name` == marker1] + nrow(marker1_data)
      marker_progress$`Papers Complete, n`[marker_progress$`Marker Name` == marker1] <- 
        marker_progress$`Papers Complete, n`[marker_progress$`Marker Name` == marker1] + papers_complete
      
      # Check if Marker 1 has completed all papers for this school
      if (papers_complete == nrow(marker1_data)) {
        marker_school_completion[[marker1]][[serial]] <- TRUE
      }
    } else {
      # If file doesn't exist, still count allocated papers as expected number
      marker_progress$`Allocated Papers, n`[marker_progress$`Marker Name` == marker1] <- 
        marker_progress$`Allocated Papers, n`[marker_progress$`Marker Name` == marker1] + num_papers
    }
  }
  
  # Check Marker 2's completion
  if (!is.na(marker2)) {
    if (file_exists(marker2_path)) {
      marker2_data <- read_excel(marker2_path)
      papers_complete <- sum(sapply(1:nrow(marker2_data), function(i) is_paper_complete(marker2_data[i,])))
      
      # Update allocated and completed papers count
      marker_progress$`Allocated Papers, n`[marker_progress$`Marker Name` == marker2] <- 
        marker_progress$`Allocated Papers, n`[marker_progress$`Marker Name` == marker2] + nrow(marker2_data)
      marker_progress$`Papers Complete, n`[marker_progress$`Marker Name` == marker2] <- 
        marker_progress$`Papers Complete, n`[marker_progress$`Marker Name` == marker2] + papers_complete
      
      # Check if Marker 2 has completed all papers for this school
      if (papers_complete == nrow(marker2_data)) {
        marker_school_completion[[marker2]][[serial]] <- TRUE
      }
    } else {
      # If file doesn't exist, still count allocated papers as expected number
      marker_progress$`Allocated Papers, n`[marker_progress$`Marker Name` == marker2] <- 
        marker_progress$`Allocated Papers, n`[marker_progress$`Marker Name` == marker2] + num_papers
    }
  }
  
  # Check Marker 3's completion
  if (!is.na(marker3)) {
    # Count allocated papers for Marker 3 (papers with disagreements)
    allocated_papers_m3 <- count_papers_with_disagreements(serial, school_row)
    completed_papers_m3 <- count_marker3_completed_papers(serial, school_row)
    
    # Update allocated and completed papers count for Marker 3
    marker_progress$`Allocated Papers, n`[marker_progress$`Marker Name` == marker3] <- 
      marker_progress$`Allocated Papers, n`[marker_progress$`Marker Name` == marker3] + allocated_papers_m3
    marker_progress$`Papers Complete, n`[marker_progress$`Marker Name` == marker3] <- 
      marker_progress$`Papers Complete, n`[marker_progress$`Marker Name` == marker3] + completed_papers_m3
    
    # Check if Marker 3 has completed all work for this school
    if (is_marker3_school_complete(serial, school_row)) {
      marker_school_completion[[marker3]][[serial]] <- TRUE
    }
  }
  
  # Update overall progress using the corrected comprehensive marking criteria
  if (file_exists(marker1_path)) {
    marker1_data <- read_excel(marker1_path)
    fully_marked_papers <- 0
    
    for (pupil_idx in 1:nrow(marker1_data)) {
      if (is_paper_marked(serial, pupil_idx, school_row)) {
        fully_marked_papers <- fully_marked_papers + 1
      }
    }
    
    # Update overall progress
    overall_progress$`Papers Marked, n`[overall_progress$`School Serial No.` == serial] <- fully_marked_papers
    overall_progress$`Papers Marked, %`[overall_progress$`School Serial No.` == serial] <- 
      round(fully_marked_papers / num_papers * 100, 1)
  }
}

# Update schools complete count for each marker based on individual completion
for (marker in names(marker_school_completion)) {
  marker_progress$`Schools Complete, n`[marker_progress$`Marker Name` == marker] <- 
    length(marker_school_completion[[marker]])
  marker_progress$`Schools Complete, %`[marker_progress$`Marker Name` == marker] <- 
    ifelse(marker_progress$`Allocated Schools, n`[marker_progress$`Marker Name` == marker] > 0,
           round(length(marker_school_completion[[marker]]) / 
                   marker_progress$`Allocated Schools, n`[marker_progress$`Marker Name` == marker] * 100, 1),
           0)
}

# Calculate marker progress percentages
marker_progress$`Papers Complete, %` <- ifelse(
  marker_progress$`Allocated Papers, n` > 0,
  round(marker_progress$`Papers Complete, n` / marker_progress$`Allocated Papers, n` * 100, 1),
  0
)

# Create the output workbook with today's date in dd-mm-yyyy format
today_date <- format(Sys.Date(), "%d-%m-%Y")
output_file <- file.path(project_dir, paste0("Checks_Dashboard_", today_date, ".xlsx"))

wb <- createWorkbook()
addWorksheet(wb, "Overall Progress")
addWorksheet(wb, "Marker Progress")

writeData(wb, "Overall Progress", overall_progress)
writeData(wb, "Marker Progress", marker_progress)

saveWorkbook(wb, output_file, overwrite = TRUE)

message("Final progress report generated at: ", output_file)

##########################################################################
# CODE THAT CREATES THIRD SHEET OF THE 'CHECKS DASHBOARD' EXCEL WORKBOOK #
##########################################################################

# Function to get marker pairs from pre-allocation
get_marker_pairs_accuracy <- function() {
  pairs <- list()
  for (i in 1:nrow(pre_allocation)) {
    serial <- pre_allocation$`School Serial No.`[i]
    marker1 <- pre_allocation$`Marker 1`[i]
    marker2 <- pre_allocation$`Marker 2`[i]
    
    if (!is.na(marker1) && !is.na(marker2)) {
      pairs[[length(pairs) + 1]] <- list(
        serial = serial,
        marker1 = marker1,
        marker2 = marker2,
        marker3 = if (!is.na(pre_allocation$`Marker 3`[i])) pre_allocation$`Marker 3`[i] else NA,
        marker4 = if (!is.na(pre_allocation$`Marker 4`[i])) pre_allocation$`Marker 4`[i] else NA
      )
    }
  }
  return(pairs)
}

# Function to determine final mark for a question
get_final_mark <- function(mark1, mark2, mark3 = NA, mark4 = NA) {
  # If no discrepancy between marker 1 and 2, return their mark
  if (!is.na(mark1) && !is.na(mark2) && mark1 == mark2) {
    return(mark1)
  }
  
  # If there's a discrepancy and marker 3 exists
  if (!is.na(mark3)) {
    # If marker 3 agrees with marker 1
    if (!is.na(mark1) && mark3 == mark1) {
      return(mark1)
    }
    # If marker 3 agrees with marker 2
    if (!is.na(mark2) && mark3 == mark2) {
      return(mark2)
    }
    # If marker 3 disagrees with both, marker 4's score is final (if available)
    if (!is.na(mark4)) {
      return(mark4)
    }
    # If marker 3 disagrees with both but no marker 4, no final mark yet
    return(NA)
  }
  
  # If there's a discrepancy but no marker 3, no final mark yet
  return(NA)
}

# Initialise results dataframe for all markers involved in marking (Marker 1 and 2 pairs)
marker_pairs <- get_marker_pairs_accuracy()
unique_markers <- unique(c(sapply(marker_pairs, function(x) x$marker1), 
                           sapply(marker_pairs, function(x) x$marker2)))
unique_markers <- unique_markers[!is.na(unique_markers)]

marker_results <- data.frame(
  `Marker Name` = unique_markers,
  `Allocated Schools, n` = 0,
  `Allocated Papers, n` = 0,
  `Papers Completed, n` = 0,
  `Papers with Discrepancies, n` = 0,
  `Papers with Discrepancies, %` = 0,
  `Papers with Discrepancies to Final Mark, n` = 0,
  `Papers with Discrepancies to Final Mark, %` = 0,
  `Items Marked, n` = 0,
  `Items with Discrepancies, n` = 0,
  `Items with Discrepancies, %` = 0,
  `Items with Discrepancies to Final Mark, n` = 0,
  `Items with Discrepancies to Final Mark, %` = 0,
  check.names = FALSE
)

# Process each marker pair
for (pair in marker_pairs) {
  serial <- pair$serial
  marker1 <- pair$marker1
  marker2 <- pair$marker2
  marker3 <- pair$marker3
  marker4 <- pair$marker4
  
  # File paths for both markers
  file1_path <- file.path(project_dir, marker1, paste0(serial, ".xlsx"))
  file2_path <- file.path(project_dir, marker2, paste0(serial, ".xlsx"))
  file3_path <- if (!is.na(marker3)) file.path(project_dir, marker3, paste0(serial, ".xlsx")) else NULL
  file4_path <- if (!is.na(marker4)) file.path(project_dir, marker4, paste0(serial, ".xlsx")) else NULL
  
  # Check if both marker files exist
  if (file_exists(file1_path) && file_exists(file2_path)) {
    # Read the data
    data1 <- read_excel(file1_path)
    data2 <- read_excel(file2_path)
    data3 <- if (!is.null(file3_path) && file_exists(file3_path)) read_excel(file3_path) else NULL
    data4 <- if (!is.null(file4_path) && file_exists(file4_path)) read_excel(file4_path) else NULL
    
    # Update assigned schools count
    marker_results$`Allocated Schools, n`[marker_results$`Marker Name` == marker1] <- 
      marker_results$`Allocated Schools, n`[marker_results$`Marker Name` == marker1] + 1
    marker_results$`Allocated Schools, n`[marker_results$`Marker Name` == marker2] <- 
      marker_results$`Allocated Schools, n`[marker_results$`Marker Name` == marker2] + 1
    
    # Update assigned papers count
    marker_results$`Allocated Papers, n`[marker_results$`Marker Name` == marker1] <- 
      marker_results$`Allocated Papers, n`[marker_results$`Marker Name` == marker1] + nrow(data1)
    marker_results$`Allocated Papers, n`[marker_results$`Marker Name` == marker2] <- 
      marker_results$`Allocated Papers, n`[marker_results$`Marker Name` == marker2] + nrow(data2)
    
    # Process each paper (row)
    for (i in 1:nrow(data1)) {
      # Check if papers are complete for each marker
      paper1_complete <- is_paper_complete(data1[i,])
      paper2_complete <- is_paper_complete(data2[i,])
      
      # Update completed papers count
      if (paper1_complete) {
        marker_results$`Papers Completed, n`[marker_results$`Marker Name` == marker1] <- 
          marker_results$`Papers Completed, n`[marker_results$`Marker Name` == marker1] + 1
      }
      if (paper2_complete) {
        marker_results$`Papers Completed, n`[marker_results$`Marker Name` == marker2] <- 
          marker_results$`Papers Completed, n`[marker_results$`Marker Name` == marker2] + 1
      }
      
      # Count items marked for each marker
      items_marked_m1 <- sum(!is.na(data1[i, question_cols]) & data1[i, question_cols] != "")
      items_marked_m2 <- sum(!is.na(data2[i, question_cols]) & data2[i, question_cols] != "")
      
      marker_results$`Items Marked, n`[marker_results$`Marker Name` == marker1] <- 
        marker_results$`Items Marked, n`[marker_results$`Marker Name` == marker1] + items_marked_m1
      marker_results$`Items Marked, n`[marker_results$`Marker Name` == marker2] <- 
        marker_results$`Items Marked, n`[marker_results$`Marker Name` == marker2] + items_marked_m2
      
      # Check for discrepancies only if both papers are complete
      if (paper1_complete && paper2_complete) {
        has_discrepancy <- FALSE
        has_discrepancy_to_final_m1 <- FALSE  # Track separately for each marker
        has_discrepancy_to_final_m2 <- FALSE  # Track separately for each marker
        
        items_with_discrepancies_m1 <- 0
        items_with_discrepancies_m2 <- 0
        items_with_discrepancies_to_final_m1 <- 0
        items_with_discrepancies_to_final_m2 <- 0
        
        # Check each question
        for (q_col in question_cols) {
          mark1 <- data1[i, q_col][[1]]
          mark2 <- data2[i, q_col][[1]]
          mark3 <- if (!is.null(data3) && i <= nrow(data3)) data3[i, q_col][[1]] else NA
          mark4 <- if (!is.null(data4) && i <= nrow(data4)) data4[i, q_col][[1]] else NA
          
          # Skip if either mark is NA or empty
          if (is.na(mark1) || is.na(mark2) || mark1 == "" || mark2 == "") next
          
          # Check for discrepancy between marker 1 and 2
          if (mark1 != mark2) {
            has_discrepancy <- TRUE
            items_with_discrepancies_m1 <- items_with_discrepancies_m1 + 1
            items_with_discrepancies_m2 <- items_with_discrepancies_m2 + 1
            
            # Get final mark
            final_mark <- get_final_mark(mark1, mark2, mark3, mark4)
            
            # Only check discrepancy to final mark if there IS a final mark
            if (!is.na(final_mark)) {
              if (mark1 != final_mark) {
                has_discrepancy_to_final_m1 <- TRUE
                items_with_discrepancies_to_final_m1 <- items_with_discrepancies_to_final_m1 + 1
              }
              if (mark2 != final_mark) {
                has_discrepancy_to_final_m2 <- TRUE
                items_with_discrepancies_to_final_m2 <- items_with_discrepancies_to_final_m2 + 1
              }
            }
            # If final_mark is NA, we don't count any discrepancies to final mark
          }
        }
        
        # Update paper-level discrepancy counts (only for completed papers)
        if (has_discrepancy) {
          marker_results$`Papers with Discrepancies, n`[marker_results$`Marker Name` == marker1] <- 
            marker_results$`Papers with Discrepancies, n`[marker_results$`Marker Name` == marker1] + 1
          marker_results$`Papers with Discrepancies, n`[marker_results$`Marker Name` == marker2] <- 
            marker_results$`Papers with Discrepancies, n`[marker_results$`Marker Name` == marker2] + 1
        }
        
        # Update paper-level discrepancy to final mark counts - SEPARATELY for each marker
        if (has_discrepancy_to_final_m1) {
          marker_results$`Papers with Discrepancies to Final Mark, n`[marker_results$`Marker Name` == marker1] <- 
            marker_results$`Papers with Discrepancies to Final Mark, n`[marker_results$`Marker Name` == marker1] + 1
        }
        
        if (has_discrepancy_to_final_m2) {
          marker_results$`Papers with Discrepancies to Final Mark, n`[marker_results$`Marker Name` == marker2] <- 
            marker_results$`Papers with Discrepancies to Final Mark, n`[marker_results$`Marker Name` == marker2] + 1
        }
        
        # Update item-level discrepancy counts
        marker_results$`Items with Discrepancies, n`[marker_results$`Marker Name` == marker1] <- 
          marker_results$`Items with Discrepancies, n`[marker_results$`Marker Name` == marker1] + items_with_discrepancies_m1
        marker_results$`Items with Discrepancies, n`[marker_results$`Marker Name` == marker2] <- 
          marker_results$`Items with Discrepancies, n`[marker_results$`Marker Name` == marker2] + items_with_discrepancies_m2
        
        marker_results$`Items with Discrepancies to Final Mark, n`[marker_results$`Marker Name` == marker1] <- 
          marker_results$`Items with Discrepancies to Final Mark, n`[marker_results$`Marker Name` == marker1] + items_with_discrepancies_to_final_m1
        marker_results$`Items with Discrepancies to Final Mark, n`[marker_results$`Marker Name` == marker2] <- 
          marker_results$`Items with Discrepancies to Final Mark, n`[marker_results$`Marker Name` == marker2] + items_with_discrepancies_to_final_m2
      }
    }
  }
}

# Calculate percentages - NOW USING PAPERS COMPLETED AS DENOMINATOR FOR PAPER DISCREPANCY PERCENTAGES
marker_results$`Papers with Discrepancies, %` <- ifelse(
  marker_results$`Papers Completed, n` > 0,
  round(marker_results$`Papers with Discrepancies, n` / marker_results$`Papers Completed, n` * 100, 1),
  0
)

marker_results$`Papers with Discrepancies to Final Mark, %` <- ifelse(
  marker_results$`Papers Completed, n` > 0,
  round(marker_results$`Papers with Discrepancies to Final Mark, n` / marker_results$`Papers Completed, n` * 100, 1),
  0
)

marker_results$`Items with Discrepancies, %` <- ifelse(
  marker_results$`Items Marked, n` > 0,
  round(marker_results$`Items with Discrepancies, n` / marker_results$`Items Marked, n` * 100, 1),
  0
)

marker_results$`Items with Discrepancies to Final Mark, %` <- ifelse(
  marker_results$`Items Marked, n` > 0,
  round(marker_results$`Items with Discrepancies to Final Mark, n` / marker_results$`Items Marked, n` * 100, 1),
  0
)

# Add to existing Checks Dashboard workbook
today_date <- format(Sys.Date(), "%d-%m-%Y")
dashboard_file <- file.path(project_dir, paste0("Checks_Dashboard_", today_date, ".xlsx"))

# Load existing workbook or create new one if it doesn't exist
if (file_exists(dashboard_file)) {
  wb <- loadWorkbook(dashboard_file)
  message("Adding to existing Checks Dashboard...")
} else {
  wb <- createWorkbook()
  message("Creating new Checks Dashboard...")
}

# Add 1st and 2nd Marker Accuracy sheet
addWorksheet(wb, "1st and 2nd Marker Accuracy")
writeData(wb, "1st and 2nd Marker Accuracy", marker_results)

# Save the workbook
saveWorkbook(wb, dashboard_file, overwrite = TRUE)

message("1st and 2nd Marker Accuracy sheet added to: ", dashboard_file)

# Display summary
print(marker_results)

##########################################
# CODE TO ASSIGN THIRD MARKERS IF NEEDED #
##########################################

# Read marker data for triple marking eligibility
markers_file_path <- file.path(project_dir, "Marker_Data.xlsx")
marker_eligibility <- read_excel(markers_file_path, sheet = 2)

# Get list of markers approved for triple marking
approved_markers <- marker_eligibility$`Marker Name`[marker_eligibility$`Approved for Triple Marking` == "Y"]
approved_markers <- approved_markers[!is.na(approved_markers)]

message("Markers approved for triple marking: ", paste(approved_markers, collapse = ", "))

# Function to check if both markers have completed all papers for a school
both_markers_complete <- function(serial, marker1, marker2) {
  file1_path <- file.path(project_dir, marker1, paste0(serial, ".xlsx"))
  file2_path <- file.path(project_dir, marker2, paste0(serial, ".xlsx"))
  
  if (!file_exists(file1_path) || !file_exists(file2_path)) {
    return(FALSE)
  }
  
  data1 <- read_excel(file1_path)
  data2 <- read_excel(file2_path)
  
  # Check if all papers are complete for both markers
  all_complete_m1 <- all(sapply(1:nrow(data1), function(i) is_paper_complete(data1[i,])))
  all_complete_m2 <- all(sapply(1:nrow(data2), function(i) is_paper_complete(data2[i,])))
  
  return(all_complete_m1 && all_complete_m2)
}

# Function to check if there are any discrepancies between two markers for a school
has_discrepancies_for_school <- function(serial, marker1, marker2) {
  file1_path <- file.path(project_dir, marker1, paste0(serial, ".xlsx"))
  file2_path <- file.path(project_dir, marker2, paste0(serial, ".xlsx"))
  
  if (!file_exists(file1_path) || !file_exists(file2_path)) {
    return(FALSE)
  }
  
  data1 <- read_excel(file1_path)
  data2 <- read_excel(file2_path)
  
  # Check each paper for discrepancies
  for (i in 1:nrow(data1)) {
    # Only check if both papers are complete
    if (is_paper_complete(data1[i,]) && is_paper_complete(data2[i,])) {
      # Check each question
      for (q_col in question_cols) {
        mark1 <- data1[i, q_col][[1]]
        mark2 <- data2[i, q_col][[1]]
        
        # Skip if either mark is NA or empty
        if (is.na(mark1) || is.na(mark2) || mark1 == "" || mark2 == "") next
        
        # If there's a discrepancy, return TRUE
        if (mark1 != mark2) {
          return(TRUE)
        }
      }
    }
  }
  
  return(FALSE)
}

# Function to randomly select a third marker
select_third_marker <- function(marker1, marker2, approved_markers) {
  # Remove marker1 and marker2 from the approved list
  available_markers <- approved_markers[!approved_markers %in% c(marker1, marker2)]
  
  if (length(available_markers) == 0) {
    return(NA)
  }
  
  # Randomly select one marker
  set.seed(NULL)  # Ensure randomness
  selected_marker <- sample(available_markers, 1)
  return(selected_marker)
}

# Track assignments made
assignments_made <- data.frame(
  `School Serial No.` = character(0),
  `Marker 1` = character(0),
  `Marker 2` = character(0),
  `Assigned Marker 3` = character(0),
  stringsAsFactors = FALSE
)

# Process each school to check for third marker assignment needs
for (i in 1:nrow(pre_allocation)) {
  serial <- pre_allocation$`School Serial No.`[i]
  marker1 <- pre_allocation$`Marker 1`[i]
  marker2 <- pre_allocation$`Marker 2`[i]
  marker3 <- pre_allocation$`Marker 3`[i]
  
  # Skip if any essential data is missing
  if (is.na(serial) || is.na(marker1) || is.na(marker2)) {
    next
  }
  
  # Skip if a third marker is already assigned
  if (!is.na(marker3) && marker3 != "") {
    message("School ", serial, " already has Marker 3 assigned: ", marker3)
    next
  }
  
  # Check if both markers have completed all papers
  if (both_markers_complete(serial, marker1, marker2)) {
    message("Both markers completed for school ", serial, " - checking for discrepancies...")
    
    # Check if there are discrepancies
    if (has_discrepancies_for_school(serial, marker1, marker2)) {
      message("Discrepancies found for school ", serial, " between ", marker1, " and ", marker2)
      
      # Select a third marker
      third_marker <- select_third_marker(marker1, marker2, approved_markers)
      
      if (!is.na(third_marker)) {
        # Update the pre-allocation data
        pre_allocation$`Marker 3`[i] <- third_marker
        
        # Track the assignment
        assignments_made <- rbind(assignments_made, data.frame(
          `School Serial No.` = serial,
          `Marker 1` = marker1,
          `Marker 2` = marker2,
          `Assigned Marker 3` = third_marker,
          stringsAsFactors = FALSE,
          check.names = FALSE
        ))
        
        message("Assigned ", third_marker, " as Marker 3 for school ", serial)
      } else {
        warning("No available markers for triple marking for school ", serial)
      }
    } else {
      message("No discrepancies found for school ", serial, " - no third marker needed")
    }
  } else {
    message("School ", serial, " - not both markers have completed all papers yet")
  }
}

# Save the updated pre-allocation file
write_xlsx(pre_allocation, pre_allocation_file)

# Display summary of assignments made
if (nrow(assignments_made) > 0) {
  message("\n=== THIRD MARKER ASSIGNMENTS MADE ===")
  print(assignments_made)
  message("Updated Pre_Allocation_of_Markers.xlsx with ", nrow(assignments_made), " new third marker assignments")
} else {
  message("\n=== NO NEW THIRD MARKER ASSIGNMENTS NEEDED ===")
  message("Either no schools have discrepancies, markers haven't completed their work, or third markers are already assigned")
}

################################################################
# CODE TO SEND AN EMAIL TO NOTIFY NEWLY-ASSIGNED THIRD MARKERS #
################################################################

# Function to send email to newly assigned third markers
send_third_marker_emails <- function(assignments_made, project_dir) {
  # Check if there are any assignments made
  if (nrow(assignments_made) == 0) {
    message("No new third marker assignments - no emails to send")
    return()
  }
  
  # Read marker data - ensure we're reading the correct columns
  markers_file_path <- file.path(project_dir, "Marker_Data.xlsx")
  marker_data <- read_excel(markers_file_path, sheet = 2) %>%
    select(`Marker Name`, `Email Address`) %>%
    filter(!is.na(`Marker Name`), !is.na(`Email Address`), `Email Address` != "")
  
  # Group assignments by third marker
  assignments_by_marker <- assignments_made %>%
    group_by(`Assigned Marker 3`) %>%
    summarise(Schools = paste(sort(`School Serial No.`), collapse = ", "), .groups = "drop")
  
  # Process each marker's assignments
  for (i in 1:nrow(assignments_by_marker)) {
    marker_name <- assignments_by_marker$`Assigned Marker 3`[i]
    schools_list <- assignments_by_marker$Schools[i]
    
    # Get marker's email
    marker_info <- marker_data %>% 
      filter(`Marker Name` == marker_name) %>%
      slice(1) # Take first match if multiple
    
    if (nrow(marker_info) == 0) {
      warning("Marker '", marker_name, "' not found in Marker_Data.xlsx or missing email address")
      next
    }
    
    email_address <- marker_info$`Email Address`[1]
    first_name <- strsplit(marker_name, " ")[[1]][1]
    
    # Create email subject and body
    subject <- "Third Marking"
    body <- sprintf(
      "Dear %s,\n\nYou have been allocated as the third marker for the following schools: %s.\n\nPlease review these files first thing tomorrow.\n\nKind regards,\nAlpha Plus",
      first_name,
      schools_list
    )
    
    # Send email
    tryCatch({
      send.mail(
        from = "eefmarking@gmail.com",
        to = email_address,
        subject = subject,
        body = body,
        smtp = list(
          host.name = "smtp.gmail.com",
          port = 587,
          user.name = "eefmarking@gmail.com",
          passwd = "dyhogkbcqzrjxcau",
          tls = TRUE
        ),
        authenticate = TRUE,
        send = TRUE
      )
      message(sprintf("Successfully sent email to %s (%s) for schools: %s", 
                      marker_name, email_address, schools_list))
    }, error = function(e) {
      warning(sprintf("Failed to send email to %s (%s): %s", 
                      marker_name, email_address, e$message))
    })
  }
}

# Call the function with the assignments made earlier
send_third_marker_emails(assignments_made, project_dir)

################################################################
# CODE TO CREATE SCHOOL FILES FOR NEWLY-ASSIGNED THIRD MARKERS #
################################################################

# Function to create third marker files with appropriate formatting
create_third_marker_files <- function(assignments_made, pre_allocation, project_dir, question_cols) {
  # Check if there are any assignments made
  if (nrow(assignments_made) == 0) {
    message("No new third marker assignments - no files to create")
    return()
  }
  
  # Process each assignment
  for (i in 1:nrow(assignments_made)) {
    serial <- assignments_made$`School Serial No.`[i]
    marker1 <- assignments_made$`Marker 1`[i]
    marker2 <- assignments_made$`Marker 2`[i]
    marker3 <- assignments_made$`Assigned Marker 3`[i]
    
    # Define file paths
    target_file <- paste0(serial, ".xlsx")
    marker1_path <- file.path(project_dir, marker1, target_file)
    marker2_path <- file.path(project_dir, marker2, target_file)
    marker3_folder <- file.path(project_dir, marker3)
    marker3_path <- file.path(marker3_folder, target_file)
    
    # Check if marker 1 and 2 files exist
    if (!file_exists(marker1_path) || !file_exists(marker2_path)) {
      warning("Missing files for school ", serial, " - cannot create third marker file")
      next
    }
    
    # Ensure marker 3 folder exists
    if (!dir_exists(marker3_folder)) {
      dir_create(marker3_folder)
      message("Created folder for ", marker3)
    }
    
    # Read data from both markers
    data1 <- read_excel(marker1_path)
    data2 <- read_excel(marker2_path)
    
    # Verify both files have same structure
    if (nrow(data1) != nrow(data2) || 
        !all(data1$`Pupil Serial No.` == data2$`Pupil Serial No.`)) {
      warning("File structure mismatch for school ", serial, " - cannot create third marker file")
      next
    }
    
    # Create workbook starting from Marker 2's file
    wb <- loadWorkbook(marker2_path)
    
    # Define styles that preserve existing formatting WITHOUT black borders
    # Only yellow fill for disagreements, no border styling
    yellow_fill <- createStyle(fgFill = "yellow", locked = FALSE)
    # Only locking for agreements, no border styling
    locked_style <- createStyle(locked = TRUE)
    
    # Process each row (pupil)
    for (row_idx in 1:nrow(data1)) {
      # Only process if both papers are complete
      if (is_paper_complete(data1[row_idx,]) && is_paper_complete(data2[row_idx,])) {
        
        # Check each question column
        for (col_idx in 1:length(question_cols)) {
          q_col <- question_cols[col_idx]
          
          # Get marks from both markers
          mark1 <- data1[row_idx, q_col][[1]]
          mark2 <- data2[row_idx, q_col][[1]]
          
          # Skip if either mark is NA or empty
          if (is.na(mark1) || is.na(mark2) || mark1 == "" || mark2 == "") next
          
          # Find the Excel column index for this question
          excel_col_idx <- which(names(data2) == q_col)
          excel_row_idx <- row_idx + 1  # +1 because Excel is 1-indexed and has header row
          
          if (mark1 == mark2) {
            # Markers agree - lock the cell while preserving original formatting
            addStyle(wb, sheet = 1, style = locked_style, 
                     rows = excel_row_idx, cols = excel_col_idx, stack = TRUE)
          } else {
            # Markers disagree - highlight in yellow and clear content
            # Clear the cell content
            writeData(wb, sheet = 1, x = "", 
                      startRow = excel_row_idx, startCol = excel_col_idx)
            # Apply yellow highlighting and unlock while preserving original formatting
            addStyle(wb, sheet = 1, style = yellow_fill, 
                     rows = excel_row_idx, cols = excel_col_idx, stack = TRUE)
          }
        }
      }
    }
    
    # Protect the worksheet to enforce cell locking
    # Note: You may want to set a password here
    protectWorksheet(wb, sheet = 1, protect = TRUE, password = NULL)
    
    # Save the modified file to Marker 3's folder
    tryCatch({
      saveWorkbook(wb, marker3_path, overwrite = TRUE)
      message("Created third marker file for ", marker3, ": ", target_file)
    }, error = function(e) {
      warning("Failed to create file for ", marker3, " (", serial, "): ", e$message)
    })
  }
}

# Call the function with the assignments made earlier
create_third_marker_files(assignments_made, pre_allocation, project_dir, question_cols)

message("Third marker file creation complete!")

###################################################################
# CODE TO PRODUCE SHEET FOUR OF 'CHECKS DASHBOARD' EXCEL WORKBOOK #
###################################################################

# Function to analyse third marker accuracy
analyse_third_marker_accuracy <- function(pre_allocation, project_dir, question_cols) {
  # Get all unique third markers assigned so far (excluding NAs)
  third_markers <- unique(pre_allocation$`Marker 3`[!is.na(pre_allocation$`Marker 3`)])
  
  # If no third markers assigned yet, return empty dataframe
  if (length(third_markers) == 0) {
    return(data.frame(
      `Marker Name` = character(),
      `Assigned Items, n` = integer(),
      `Items Marked, n` = integer(),
      `Items Marked, %` = numeric(),
      `Items Flagged for Fourth Marking, n` = integer(),
      `Items Flagged for Fourth Marking, %` = numeric(),
      check.names = FALSE
    ))
  }
  
  # Initialise results dataframe
  third_marker_results <- data.frame(
    `Marker Name` = third_markers,
    `Assigned Items, n` = 0,
    `Items Marked, n` = 0,
    `Items Marked, %` = 0,
    `Items Flagged for Fourth Marking, n` = 0,
    `Items Flagged for Fourth Marking, %` = 0,
    check.names = FALSE
  )
  
  # Process each school with a third marker assigned
  for (i in 1:nrow(pre_allocation)) {
    serial <- pre_allocation$`School Serial No.`[i]
    marker1 <- pre_allocation$`Marker 1`[i]
    marker2 <- pre_allocation$`Marker 2`[i]
    marker3 <- pre_allocation$`Marker 3`[i]
    
    # Skip if no third marker assigned
    if (is.na(marker3)) next
    
    # File paths
    file1_path <- file.path(project_dir, marker1, paste0(serial, ".xlsx"))
    file2_path <- file.path(project_dir, marker2, paste0(serial, ".xlsx"))
    file3_path <- file.path(project_dir, marker3, paste0(serial, ".xlsx"))
    
    # Check if all files exist
    if (!file_exists(file1_path) || !file_exists(file2_path) || !file_exists(file3_path)) {
      warning("Missing files for school ", serial, " - skipping")
      next
    }
    
    # Read data
    data1 <- read_excel(file1_path)
    data2 <- read_excel(file2_path)
    data3 <- read_excel(file3_path)
    
    # Verify all files have same structure
    if (nrow(data1) != nrow(data2) || nrow(data1) != nrow(data3) ||
        !all(data1$`Pupil Serial No.` == data2$`Pupil Serial No.`) ||
        !all(data1$`Pupil Serial No.` == data3$`Pupil Serial No.`)) {
      warning("File structure mismatch for school ", serial, " - skipping")
      next
    }
    
    # Find the third marker in our results
    marker_idx <- which(third_marker_results$`Marker Name` == marker3)
    
    # Process each pupil
    for (row_idx in 1:nrow(data1)) {
      # Process each question column
      for (q_col in question_cols) {
        mark1 <- data1[row_idx, q_col][[1]]
        mark2 <- data2[row_idx, q_col][[1]]
        mark3 <- data3[row_idx, q_col][[1]]
        
        # Skip if either mark1 or mark2 is NA/empty (shouldn't happen for completed papers)
        if (is.na(mark1) || is.na(mark2) || mark1 == "" || mark2 == "") next
        
        # Check if this cell was a disagreement between marker1 and marker2 (yellow in marker3's file)
        if (mark1 != mark2) {
          # This is an assigned item for the third marker
          third_marker_results$`Assigned Items, n`[marker_idx] <- 
            third_marker_results$`Assigned Items, n`[marker_idx] + 1
          
          # Check if third marker has marked it (not NA/empty)
          if (!is.na(mark3) && mark3 != "") {
            third_marker_results$`Items Marked, n`[marker_idx] <- 
              third_marker_results$`Items Marked, n`[marker_idx] + 1
            
            # Check if third marker's mark differs from BOTH marker1 and marker2
            if (mark3 != mark1 && mark3 != mark2) {
              third_marker_results$`Items Flagged for Fourth Marking, n`[marker_idx] <- 
                third_marker_results$`Items Flagged for Fourth Marking, n`[marker_idx] + 1
            }
          }
        }
      }
    }
  }
  
  # Calculate percentages
  third_marker_results$`Items Marked, %` <- ifelse(
    third_marker_results$`Assigned Items, n` > 0,
    round(third_marker_results$`Items Marked, n` / third_marker_results$`Assigned Items, n` * 100, 1),
    0
  )
  
  third_marker_results$`Items Flagged for Fourth Marking, %` <- ifelse(
    third_marker_results$`Items Marked, n` > 0,
    round(third_marker_results$`Items Flagged for Fourth Marking, n` / third_marker_results$`Items Marked, n` * 100, 1),
    0
  )
  
  return(third_marker_results)
}

# Generate the third marker accuracy report
third_marker_accuracy <- analyse_third_marker_accuracy(pre_allocation, project_dir, question_cols)

# Add to existing Checks Dashboard workbook
today_date <- format(Sys.Date(), "%d-%m-%Y")
dashboard_file <- file.path(project_dir, paste0("Checks_Dashboard_", today_date, ".xlsx"))

# Load existing workbook or create new one if it doesn't exist
if (file_exists(dashboard_file)) {
  wb <- loadWorkbook(dashboard_file)
  message("Updating 3rd Marker Accuracy sheet in existing Checks Dashboard...")
  
  # Remove existing sheet if it exists
  if ("3rd Marker Accuracy" %in% names(wb)) {
    removeWorksheet(wb, "3rd Marker Accuracy")
  }
} else {
  wb <- createWorkbook()
  message("Creating new Checks Dashboard with 3rd Marker Accuracy sheet...")
}

# Add/update 3rd Marker Accuracy sheet
addWorksheet(wb, "3rd Marker Accuracy")
writeData(wb, "3rd Marker Accuracy", third_marker_accuracy)

# Save the workbook
saveWorkbook(wb, dashboard_file, overwrite = TRUE)

message("3rd Marker Accuracy sheet updated in: ", dashboard_file)

# Display summary
print(third_marker_accuracy)

##########################################
# CODE TO ASSIGN FOURTH MARKER IF NEEDED #
##########################################

# Define list of internal markers approved for fourth marking
internal_markers <- c("Mibeka Tan Panza")

# Function to check if third marker has completed all assigned items for a school
third_marker_complete <- function(serial, marker1, marker2, marker3) {
  file1_path <- file.path(project_dir, marker1, paste0(serial, ".xlsx"))
  file2_path <- file.path(project_dir, marker2, paste0(serial, ".xlsx"))
  file3_path <- file.path(project_dir, marker3, paste0(serial, ".xlsx"))
  
  # Check if all files exist
  if (!file_exists(file1_path) || !file_exists(file2_path) || !file_exists(file3_path)) {
    return(FALSE)
  }
  
  # Read data
  data1 <- read_excel(file1_path)
  data2 <- read_excel(file2_path)
  data3 <- read_excel(file3_path)
  
  # Check each row and question
  for (row_idx in 1:nrow(data1)) {
    for (q_col in question_cols) {
      mark1 <- data1[row_idx, q_col][[1]]
      mark2 <- data2[row_idx, q_col][[1]]
      mark3 <- data3[row_idx, q_col][[1]]
      
      # Skip if either mark1 or mark2 is NA/empty
      if (is.na(mark1) || is.na(mark2) || mark1 == "" || mark2 == "") next
      
      # If this was a disagreement between marker1 and marker2
      if (mark1 != mark2) {
        # Check if third marker has marked it
        if (is.na(mark3) || mark3 == "") {
          return(FALSE)  # Third marker hasn't completed all assigned items
        }
      }
    }
  }
  
  return(TRUE)  # All assigned items completed
}

# Function to check if there are items flagged for fourth marking
has_fourth_marking_items <- function(serial, marker1, marker2, marker3) {
  file1_path <- file.path(project_dir, marker1, paste0(serial, ".xlsx"))
  file2_path <- file.path(project_dir, marker2, paste0(serial, ".xlsx"))
  file3_path <- file.path(project_dir, marker3, paste0(serial, ".xlsx"))
  
  # Check if all files exist
  if (!file_exists(file1_path) || !file_exists(file2_path) || !file_exists(file3_path)) {
    return(FALSE)
  }
  
  # Read data
  data1 <- read_excel(file1_path)
  data2 <- read_excel(file2_path)
  data3 <- read_excel(file3_path)
  
  # Check each row and question
  for (row_idx in 1:nrow(data1)) {
    for (q_col in question_cols) {
      mark1 <- data1[row_idx, q_col][[1]]
      mark2 <- data2[row_idx, q_col][[1]]
      mark3 <- data3[row_idx, q_col][[1]]
      
      # Skip if either mark1 or mark2 is NA/empty
      if (is.na(mark1) || is.na(mark2) || mark1 == "" || mark2 == "") next
      
      # If this was a disagreement between marker1 and marker2
      if (mark1 != mark2) {
        # Check if third marker disagrees with both (flagged for fourth marking)
        if (!is.na(mark3) && mark3 != "" && mark3 != mark1 && mark3 != mark2) {
          return(TRUE)
        }
      }
    }
  }
  
  return(FALSE)
}

# Function to randomly select a fourth marker
select_fourth_marker <- function(internal_markers) {
  if (length(internal_markers) == 0) {
    return(NA)
  }
  
  # Randomly select one internal marker
  set.seed(NULL)  # Ensure randomness
  selected_marker <- sample(internal_markers, 1)
  return(selected_marker)
}

# Function to create fourth marker file
create_fourth_marker_file <- function(serial, marker1, marker2, marker3, marker4) {
  # File paths
  file1_path <- file.path(project_dir, marker1, paste0(serial, ".xlsx"))
  file2_path <- file.path(project_dir, marker2, paste0(serial, ".xlsx"))
  file3_path <- file.path(project_dir, marker3, paste0(serial, ".xlsx"))
  marker4_folder <- file.path(project_dir, marker4)
  marker4_path <- file.path(marker4_folder, paste0(serial, ".xlsx"))
  
  # Ensure marker 4 folder exists
  if (!dir_exists(marker4_folder)) {
    dir_create(marker4_folder)
    message("Created folder for ", marker4)
  }
  
  # Read data from all three markers
  data1 <- read_excel(file1_path)
  data2 <- read_excel(file2_path)
  data3 <- read_excel(file3_path)
  
  # Create workbook starting from Marker 3's file
  wb <- loadWorkbook(file3_path)
  
  # Define styles
  yellow_fill <- createStyle(fgFill = "yellow", locked = FALSE)
  locked_style <- createStyle(locked = TRUE)
  
  # Process each row and question
  for (row_idx in 1:nrow(data1)) {
    for (col_idx in 1:length(question_cols)) {
      q_col <- question_cols[col_idx]
      
      mark1 <- data1[row_idx, q_col][[1]]
      mark2 <- data2[row_idx, q_col][[1]]
      mark3 <- data3[row_idx, q_col][[1]]
      
      # Skip if either mark1 or mark2 is NA/empty
      if (is.na(mark1) || is.na(mark2) || mark1 == "" || mark2 == "") next
      
      # Find the Excel column index for this question
      excel_col_idx <- which(names(data3) == q_col)
      excel_row_idx <- row_idx + 1  # +1 because Excel is 1-indexed and has header row
      
      # If this was a disagreement between marker1 and marker2
      if (mark1 != mark2) {
        # Check if third marker disagrees with both (flagged for fourth marking)
        if (!is.na(mark3) && mark3 != "" && mark3 != mark1 && mark3 != mark2) {
          # This item is flagged for fourth marking - clear content and keep yellow
          writeData(wb, sheet = 1, x = "", 
                    startRow = excel_row_idx, startCol = excel_col_idx)
          addStyle(wb, sheet = 1, style = yellow_fill, 
                   rows = excel_row_idx, cols = excel_col_idx, stack = TRUE)
        } else {
          # This item is not flagged for fourth marking - un-yellow and lock
          # Remove yellow by applying a style without fill, then lock
          addStyle(wb, sheet = 1, style = locked_style, 
                   rows = excel_row_idx, cols = excel_col_idx, stack = FALSE)
        }
      }
    }
  }
  
  # Protect the worksheet to enforce cell locking
  protectWorksheet(wb, sheet = 1, protect = TRUE, password = NULL)
  
  # Save the modified file to Marker 4's folder
  tryCatch({
    saveWorkbook(wb, marker4_path, overwrite = TRUE)
    message("Created fourth marker file for ", marker4, ": ", paste0(serial, ".xlsx"))
    return(TRUE)
  }, error = function(e) {
    warning("Failed to create fourth marker file for ", marker4, " (", serial, "): ", e$message)
    return(FALSE)
  })
}

# Track fourth marker assignments made
fourth_marker_assignments <- data.frame(
  `School Serial No.` = character(0),
  `Marker 1` = character(0),
  `Marker 2` = character(0),
  `Marker 3` = character(0),
  `Assigned Marker 4` = character(0),
  stringsAsFactors = FALSE
)

# Process each school to check for fourth marker assignment needs
for (i in 1:nrow(pre_allocation)) {
  serial <- pre_allocation$`School Serial No.`[i]
  marker1 <- pre_allocation$`Marker 1`[i]
  marker2 <- pre_allocation$`Marker 2`[i]
  marker3 <- pre_allocation$`Marker 3`[i]
  marker4 <- pre_allocation$`Marker 4`[i]
  
  # Skip if any essential data is missing or no third marker assigned
  if (is.na(serial) || is.na(marker1) || is.na(marker2) || is.na(marker3)) {
    next
  }
  
  # Skip if a fourth marker is already assigned
  if (!is.na(marker4) && marker4 != "") {
    message("School ", serial, " already has Marker 4 assigned: ", marker4)
    next
  }
  
  # Check if third marker has completed all assigned items
  if (third_marker_complete(serial, marker1, marker2, marker3)) {
    message("Third marker completed for school ", serial, " - checking for fourth marking items...")
    
    # Check if there are items flagged for fourth marking
    if (has_fourth_marking_items(serial, marker1, marker2, marker3)) {
      message("Items flagged for fourth marking found for school ", serial)
      
      # Select a fourth marker
      fourth_marker <- select_fourth_marker(internal_markers)
      
      if (!is.na(fourth_marker)) {
        # Create the fourth marker file
        if (create_fourth_marker_file(serial, marker1, marker2, marker3, fourth_marker)) {
          # Update the pre-allocation data
          pre_allocation$`Marker 4`[i] <- fourth_marker
          
          # Track the assignment
          fourth_marker_assignments <- rbind(fourth_marker_assignments, data.frame(
            `School Serial No.` = serial,
            `Marker 1` = marker1,
            `Marker 2` = marker2,
            `Marker 3` = marker3,
            `Assigned Marker 4` = fourth_marker,
            stringsAsFactors = FALSE,
            check.names = FALSE
          ))
          
          message("Assigned ", fourth_marker, " as Marker 4 for school ", serial)
        }
      } else {
        warning("No available internal markers for fourth marking for school ", serial)
      }
    } else {
      message("No items flagged for fourth marking for school ", serial, " - no fourth marker needed")
    }
  } else {
    message("School ", serial, " - third marker has not completed all assigned items yet")
  }
}

# Save the updated pre-allocation file
write_xlsx(pre_allocation, pre_allocation_file)

# Display summary of fourth marker assignments made
if (nrow(fourth_marker_assignments) > 0) {
  message("\n=== FOURTH MARKER ASSIGNMENTS MADE ===")
  print(fourth_marker_assignments)
  message("Updated Pre_Allocation_of_Markers.xlsx with ", nrow(fourth_marker_assignments), " new fourth marker assignments")
} else {
  message("\n=== NO NEW FOURTH MARKER ASSIGNMENTS NEEDED ===")
  message("Either no schools have items flagged for fourth marking, third markers haven't completed their work, or fourth markers are already assigned")
}

#######################################################
# OPTIONAL: CODE TO LOCK FULLY-COMPLETED SCHOOL FILES #
#######################################################

# Function to check if a marker file is complete (no blanks in required cells)
is_marker_file_complete <- function(file_path, marker_type, serial, marker1 = NULL, marker2 = NULL, marker3 = NULL) {
  if (!file_exists(file_path)) {
    return(FALSE)
  }
  
  tryCatch({
    data <- read_excel(file_path)
    
    # Check each row and question column
    for (row_idx in 1:nrow(data)) {
      for (q_col in question_cols) {
        cell_value <- data[row_idx, q_col][[1]]
        
        # Determine if this cell should be filled based on marker type
        should_be_filled <- FALSE
        
        if (marker_type == "first" || marker_type == "second") {
          # First and second markers should fill all cells
          should_be_filled <- TRUE
        } else if (marker_type == "third") {
          # Third marker only fills cells where first and second markers disagreed
          if (!is.null(marker1) && !is.null(marker2)) {
            file1_path <- file.path(project_dir, marker1, paste0(serial, ".xlsx"))
            file2_path <- file.path(project_dir, marker2, paste0(serial, ".xlsx"))
            
            if (file_exists(file1_path) && file_exists(file2_path)) {
              data1 <- read_excel(file1_path)
              data2 <- read_excel(file2_path)
              
              mark1 <- data1[row_idx, q_col][[1]]
              mark2 <- data2[row_idx, q_col][[1]]
              
              # Only check if both markers marked and disagreed
              if (!is.na(mark1) && !is.na(mark2) && mark1 != "" && mark2 != "" && mark1 != mark2) {
                should_be_filled <- TRUE
              }
            }
          }
        } else if (marker_type == "fourth") {
          # Fourth marker only fills cells where third marker disagreed with both first and second
          if (!is.null(marker1) && !is.null(marker2) && !is.null(marker3)) {
            file1_path <- file.path(project_dir, marker1, paste0(serial, ".xlsx"))
            file2_path <- file.path(project_dir, marker2, paste0(serial, ".xlsx"))
            file3_path <- file.path(project_dir, marker3, paste0(serial, ".xlsx"))
            
            if (file_exists(file1_path) && file_exists(file2_path) && file_exists(file3_path)) {
              data1 <- read_excel(file1_path)
              data2 <- read_excel(file2_path)
              data3 <- read_excel(file3_path)
              
              mark1 <- data1[row_idx, q_col][[1]]
              mark2 <- data2[row_idx, q_col][[1]]
              mark3 <- data3[row_idx, q_col][[1]]
              
              # Only check if first two disagreed and third disagreed with both
              if (!is.na(mark1) && !is.na(mark2) && mark1 != "" && mark2 != "" && 
                  mark1 != mark2 && !is.na(mark3) && mark3 != "" && 
                  mark3 != mark1 && mark3 != mark2) {
                should_be_filled <- TRUE
              }
            }
          }
        }
        
        # If this cell should be filled but is empty, file is not complete
        if (should_be_filled && (is.na(cell_value) || cell_value == "")) {
          return(FALSE)
        }
      }
    }
    
    return(TRUE)
  }, error = function(e) {
    warning("Error checking completeness of ", file_path, ": ", e$message)
    return(FALSE)
  })
}

# Function to lock an Excel file with minimal formatting impact
lock_excel_file <- function(file_path) {
  tryCatch({
    # Load the workbook
    wb <- loadWorkbook(file_path)
    
    # Simply protect the worksheet without modifying cell styles
    # This preserves all existing formatting while preventing editing 
    protectWorksheet(wb, sheet = 1, protect = TRUE, password = NULL)
    
    # Save the workbook
    saveWorkbook(wb, file_path, overwrite = TRUE)
    
    return(TRUE)
  }, error = function(e) {
    warning("Failed to lock file ", file_path, ": ", e$message)
    return(FALSE)
  })
}

# Track locked files
locked_files_summary <- data.frame(
  `School Serial No.` = character(0),
  `Marker` = character(0),
  `Marker Type` = character(0),
  `File Path` = character(0),
  `Locked Successfully` = logical(0),
  stringsAsFactors = FALSE
)

message("\n=== CHECKING AND LOCKING COMPLETE MARKER FILES ===")

# Process each school in the allocation
for (i in 1:nrow(pre_allocation)) {
  serial <- pre_allocation$`School Serial No.`[i]
  marker1 <- pre_allocation$`Marker 1`[i]
  marker2 <- pre_allocation$`Marker 2`[i]
  marker3 <- pre_allocation$`Marker 3`[i]
  marker4 <- pre_allocation$`Marker 4`[i]
  
  # Skip if essential data is missing
  if (is.na(serial)) next
  
  # Check and lock first marker file
  if (!is.na(marker1) && marker1 != "") {
    file1_path <- file.path(project_dir, marker1, paste0(serial, ".xlsx"))
    if (file_exists(file1_path)) {
      if (is_marker_file_complete(file1_path, "first", serial)) {
        locked_success <- lock_excel_file(file1_path)
        locked_files_summary <- rbind(locked_files_summary, data.frame(
          `School Serial No.` = serial,
          `Marker` = marker1,
          `Marker Type` = "First",
          `File Path` = file1_path,
          `Locked Successfully` = locked_success,
          stringsAsFactors = FALSE,
          check.names = FALSE
        ))
        if (locked_success) {
          message("Locked complete first marker file: ", marker1, " - ", serial, ".xlsx")
        }
      }
    }
  }
  
  # Check and lock second marker file
  if (!is.na(marker2) && marker2 != "") {
    file2_path <- file.path(project_dir, marker2, paste0(serial, ".xlsx"))
    if (file_exists(file2_path)) {
      if (is_marker_file_complete(file2_path, "second", serial)) {
        locked_success <- lock_excel_file(file2_path)
        locked_files_summary <- rbind(locked_files_summary, data.frame(
          `School Serial No.` = serial,
          `Marker` = marker2,
          `Marker Type` = "Second",
          `File Path` = file2_path,
          `Locked Successfully` = locked_success,
          stringsAsFactors = FALSE,
          check.names = FALSE
        ))
        if (locked_success) {
          message("Locked complete second marker file: ", marker2, " - ", serial, ".xlsx")
        }
      }
    }
  }
  
  # Check and lock third marker file
  if (!is.na(marker3) && marker3 != "") {
    file3_path <- file.path(project_dir, marker3, paste0(serial, ".xlsx"))
    if (file_exists(file3_path)) {
      if (is_marker_file_complete(file3_path, "third", serial, marker1, marker2)) {
        locked_success <- lock_excel_file(file3_path)
        locked_files_summary <- rbind(locked_files_summary, data.frame(
          `School Serial No.` = serial,
          `Marker` = marker3,
          `Marker Type` = "Third",
          `File Path` = file3_path,
          `Locked Successfully` = locked_success,
          stringsAsFactors = FALSE,
          check.names = FALSE
        ))
        if (locked_success) {
          message("Locked complete third marker file: ", marker3, " - ", serial, ".xlsx")
        }
      }
    }
  }
  
  # Check and lock fourth marker file (if assigned)
  if (!is.na(marker4) && marker4 != "") {
    file4_path <- file.path(project_dir, marker4, paste0(serial, ".xlsx"))
    if (file_exists(file4_path)) {
      if (is_marker_file_complete(file4_path, "fourth", serial, marker1, marker2, marker3)) {
        locked_success <- lock_excel_file(file4_path)
        locked_files_summary <- rbind(locked_files_summary, data.frame(
          `School Serial No.` = serial,
          `Marker` = marker4,
          `Marker Type` = "Fourth",
          `File Path` = file4_path,
          `Locked Successfully` = locked_success,
          stringsAsFactors = FALSE,
          check.names = FALSE
        ))
        if (locked_success) {
          message("Locked complete fourth marker file: ", marker4, " - ", serial, ".xlsx")
        }
      }
    }
  }
}

# Display summary of locked files
if (nrow(locked_files_summary) > 0) {
  message("\n=== SUMMARY OF LOCKED FILES ===")
  print(locked_files_summary)
  
  successful_locks <- sum(locked_files_summary$`Locked Successfully`)
  total_attempts <- nrow(locked_files_summary)
  
  message("Successfully locked ", successful_locks, " out of ", total_attempts, " complete marker files")
  
  # Show any failures
  failed_locks <- locked_files_summary[!locked_files_summary$`Locked Successfully`, ]
  if (nrow(failed_locks) > 0) {
    message("\nFailed to lock the following files:")
    print(failed_locks[, c("School Serial No.", "Marker", "Marker Type")])
  }
} else {
  message("No complete marker files found to lock")
}

message("\n=== FILE LOCKING PROCESS COMPLETED ===")

##################################################
# CODE TO ARCHIVE SCHOOLS THAT ARE 100% COMPLETE #
##################################################

# Function to determine the final version marker for a school
get_final_version_marker <- function(serial, pre_allocation_row) {
  marker1 <- pre_allocation_row$`Marker 1`
  marker2 <- pre_allocation_row$`Marker 2`
  marker3 <- pre_allocation_row$`Marker 3`
  marker4 <- pre_allocation_row$`Marker 4`
  
  # Return the highest level marker that exists (in reverse order of completion)
  if (!is.na(marker4) && marker4 != "") {
    return(marker4)
  } else if (!is.na(marker3) && marker3 != "") {
    return(marker3)
  } else {
    # For schools with only marker 1 and 2, we need to pick one
    # We'll use marker 1 as the source since both should be identical
    return(marker1)
  }
}

# Function to create clean archived file (minimal formatting changes)
create_clean_archive_file <- function(source_file_path, archive_file_path) {
  tryCatch({
    # Load the source workbook
    wb <- loadWorkbook(source_file_path)
    
    # Simply protect the worksheet without modifying any cell styles
    # This preserves all existing formatting while preventing editing
    protectWorksheet(wb, sheet = 1, protect = TRUE, password = NULL)
    
    # Save to archive location
    saveWorkbook(wb, archive_file_path, overwrite = TRUE)
    return(TRUE)
  }, error = function(e) {
    warning("Failed to create archive file: ", e$message)
    return(FALSE)
  })
}

# Create archive folder if it doesn't exist
archive_folder <- file.path(project_dir, "Archive")
if (!dir_exists(archive_folder)) {
  dir_create(archive_folder)
  message("Created Archive folder")
}

# Read the current Checks Dashboard to identify 100% complete schools
today_date <- format(Sys.Date(), "%d-%m-%Y")
dashboard_file <- file.path(project_dir, paste0("Checks_Dashboard_", today_date, ".xlsx"))

if (!file_exists(dashboard_file)) {
  warning("Checks Dashboard file not found: ", dashboard_file)
  message("Cannot proceed with archiving without the dashboard")
} else {
  # Read the Overall Progress sheet
  overall_progress_current <- read_excel(dashboard_file, sheet = "Overall Progress")
  
  # Identify schools that are 100% complete
  complete_schools <- overall_progress_current %>%
    filter(`Papers Marked, %` == 100) %>%
    pull(`School Serial No.`)
  
  message("Found ", length(complete_schools), " schools at 100% completion")
  
  if (length(complete_schools) > 0) {
    # Track archived schools
    archived_schools <- data.frame(
      School_Serial_No = character(0),
      Final_Version_Marker = character(0),
      Source_File = character(0),
      Archive_File = character(0),
      Archive_Status = character(0),
      stringsAsFactors = FALSE
    )
    
    # Process each complete school
    for (serial in complete_schools) {
      # Get the pre-allocation row for this school
      pre_allocation_row <- pre_allocation %>% 
        filter(`School Serial No.` == serial) %>% 
        slice(1)
      
      if (nrow(pre_allocation_row) == 0) {
        warning("School ", serial, " not found in pre-allocation data")
        next
      }
      
      # Determine which marker has the final version
      final_marker <- get_final_version_marker(serial, pre_allocation_row)
      
      # Define file paths
      source_file <- paste0(serial, ".xlsx")
      source_file_path <- file.path(project_dir, final_marker, source_file)
      archive_file_path <- file.path(archive_folder, source_file)
      
      # Check if source file exists
      if (!file_exists(source_file_path)) {
        warning("Source file not found for school ", serial, ": ", source_file_path)
        archived_schools <- rbind(archived_schools, data.frame(
          School_Serial_No = serial,
          Final_Version_Marker = final_marker,
          Source_File = source_file_path,
          Archive_File = archive_file_path,
          Archive_Status = "FAILED - Source file not found",
          stringsAsFactors = FALSE
        ))
        next
      }
      
      # Create the clean archive file
      if (create_clean_archive_file(source_file_path, archive_file_path)) {
        archived_schools <- rbind(archived_schools, data.frame(
          School_Serial_No = serial,
          Final_Version_Marker = final_marker,
          Source_File = source_file_path,
          Archive_File = archive_file_path,
          Archive_Status = "SUCCESS",
          stringsAsFactors = FALSE
        ))
        message("Successfully archived school ", serial, " from ", final_marker, "'s folder")
      } else {
        archived_schools <- rbind(archived_schools, data.frame(
          School_Serial_No = serial,
          Final_Version_Marker = final_marker,
          Source_File = source_file_path,
          Archive_File = archive_file_path,
          Archive_Status = "FAILED - Could not create archive file",
          stringsAsFactors = FALSE
        ))
      }
    }
    
    # Display summary of archiving results
    if (nrow(archived_schools) > 0) {
      message("\n=== SCHOOL ARCHIVING SUMMARY ===")
      print(archived_schools)
      
      # Count successful archives
      successful_archives <- sum(archived_schools$Archive_Status == "SUCCESS")
      message("\nSuccessfully archived ", successful_archives, " out of ", nrow(archived_schools), " complete schools")
      
      # Show any failures
      failed_archives <- archived_schools %>% filter(Archive_Status != "SUCCESS")
      if (nrow(failed_archives) > 0) {
        message("\nFailed to archive the following schools:")
        print(failed_archives[, c("School_Serial_No", "Archive_Status")])
      }
    }
    
  } else {
    message("No schools are at 100% completion - nothing to archive")
  }
}

message("\nArchiving process complete!")