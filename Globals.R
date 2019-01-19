# Globals.R

# Boolean which changes the application bnehavior if executed from "Knit" or from "Run" within the RStudio IDE.
run_RMarkdown <- FALSE

# Working Directory
Antoine_Working_Directory <- 'C:/Users/antoi/Desktop/MBA/Mes cours/Data Mining'
Nicolas_Working_Directory <- '/Users/nrobin/Documents/GitHub/projet_datamining_1'
Teachers_Working_Directory <- 'TBD'
Selected_Working_Directory <- Nicolas_Working_Directory
File_To_Analyze <- "base_credit.xls"

# Libraries Loading.
# Function "loadLibraries".
# This function loads all the libraries needed for this project.
# They are automatically selected and installed in case they can not be found on the user IDE.
loadLibraries <- function() {
  
  libraries_utilies <- c('data.table', 'dplyr', 'DT', 'rlist', 'stringr', 'svDialogs', 'tidyr', 'utils', 'xlsx')
  
  for (package in libraries_utilies) {
    if (!require(package, character.only=T, quietly=T)) {
      install.packages(package, repos = "https://cran.rstudio.com")
      library(package, character.only=T)
    }
  } 
}

# Load all the libraries needed for the project.
loadLibraries()

# Index used for monitoring the row index when writing in a worksheet.
worksheetWriteRowIndex <- 1

