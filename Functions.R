source("Globals.R")

 # Functions

# Function "getDataFilePath".
# This function returns the path of the file that the user wants to analyze.
# If no file is selected, then the function return "NA".
getDataFilePath <- function(default_file_path = "") {
  
  return_value <- ""
  
  if (run_RMarkdown == TRUE) {
    
    return_value <- paste0(Selected_Working_Directory, "/", File_To_Analyze)
    
  }else{
    
    # If the default_file_path is blank, it means we want the user dialog for manual selection.
    if (default_file_path == "") {
      
      #file_path <- dlg_open(title = "Select the data file to analyze", filters = dlg_filters["ALL", ])$res
      file_path <- dlg_open(default=Selected_Working_Directory, title= "Select the data file to analyze", multiple = FALSE, filters = dlg_filters["All", ], gui = .GUI)$res
      
      if (identical(file_path, character(0))) {
        
        return_value <- default_file_path
        
      }else{
        
        # Ajout d'un "/" pour le dossier sélectionné par l'utilisateur.
        return_value <- file_path
      }
      
    }else{
      return_value <- default_file_path
    }
    
    # An empty return value means that the user press "Cancel". He did not select a file.
    if (return_value == "") {
      dlg_message("Le dossier des fichiers d'entree n'a pas ete choisi. Les tables ne peuvent pas être chargees.", "ok", gui = .GUI)
      dlg_message("No file was selected. The execution of the script will stop.", "ok", gui = .GUI)
      return_value = NA
    }
  }
  
  return(return_value)
}

strEndsWith <- function(haystack, needle)
{
  hl <- nchar(haystack)
  nl <- nchar(needle)
  if(nl>hl)
  {
    return(F)
  } else
  {
    return(substr(haystack, hl-nl+1, hl) == needle)
  }
}

# descriptives_stats_list
# This function returns the descriptive stats (as list)
# Iput : a column with a continuous variables.
# The matrix hols the following stats :
# Moyenne, Ecart type, Minimum, Maximum, Min2, Max2, Variance",
# Coefficient de variation, Médiane, Valeurs manquantes, Taux de valeurs manquantes")
descriptives_stats_list <- function(varname, x) {
  
  # Define the decimal precision
  precision = 3
  
  created_list <- list( 
    Variable=varname,
    # Calcul du nombre de lignes, sans données manquantes
    Effectif=length(x[!is.na(x)]),
    # Calcul de la moyenne
    Moyenne=round(mean(x, na.rm=TRUE),precision),
    # Calcul de l'ecart-type
    'Ecart Type'=round(sd(x, na.rm=TRUE),precision),
    # Minimum
    Minimum=min(x,na.rm = TRUE),
    # Maximum
    Maximum=max(x,na.rm = TRUE),
    # Min2
    'Min 2'=sort(unique(x,na.rm = TRUE))[2],
    # Max2
    'Max 2'=sort(unique(x,na.rm = TRUE),decreasing = TRUE)[2],
    # Calcul du Coefficiant de Variation
    'Coefficient de Variation'=round(sd(x, na.rm=TRUE) / mean(x, na.rm=TRUE),precision),
    # Calcul de la médiane
    Mediane=median(x, na.rm=TRUE),
    # Calcul de la variance
    Variance=round(var(x, na.rm=TRUE),precision),
    # Nombre de données manquantes
    'Valeurs manquantes'=length(x[is.na(x)]),
    # Pourcentage de valeurs manquantes
    'Taux de valeurs manquantes'=length(x[is.na(x)])
    # Calcul des quantiles
    # Quantile=quantile(x, na.rm=TRUE)
  )
  return(created_list)
}

# descriptives_stats_list
# This function returns the descriptive stats (as data.frame)
# Iput : a column with a continuous variables.
# The matrix hols the following stats :
# Moyenne, Ecart type, Minimum, Maximum, Min2, Max2, Variance",
# Coefficient de variation, Médiane, Valeurs manquantes, Taux de valeurs manquantes")
descriptives_stats_dataframe <- function(varname, x) {
  
  # Define the decimal precision
  precision = 3
  
  created_data_frame <- data.frame( 
    Variable=varname,
    # Calcul du nombre de lignes, sans données manquantes
    Effectif=length(x[!is.na(x)]),
    # Calcul de la moyenne
    Moyenne=round(mean(x, na.rm=TRUE),precision),
    # Calcul de l'ecart-type
    'Ecart Type'=round(sd(x, na.rm=TRUE),precision),
    # Minimum
    Minimum=min(x,na.rm = TRUE),
    # Maximum
    Maximum=max(x,na.rm = TRUE),
    # Min2
    'Min 2'=sort(unique(x,na.rm = TRUE))[2],
    # Max2
    'Max 2'=sort(unique(x,na.rm = TRUE),decreasing = TRUE)[2],
    # Calcul du Coefficiant de Variation
    'Coefficient de Variation'=round(sd(x, na.rm=TRUE) / mean(x, na.rm=TRUE), precision),
    # Calcul de la médiane
    Mediane=median(x, na.rm=TRUE),
    # Calcul de la variance
    Variance=round(var(x, na.rm=TRUE), precision),
    # Nombre de données manquantes
    'Valeurs manquantes'=length(x[is.na(x)]),
    # Pourcentage de valeurs manquantes
    'Taux de valeurs manquantes'=length(x[is.na(x)])
    # Calcul des quantiles
    # Quantile=quantile(x, na.rm=TRUE)
  )
  return(created_data_frame)
}

get_variable_set <- function(work_table) {
  
  column_count = ncol(work_table)
  row_count = nrow(work_table)
  
  variable_set <- c()
  worklist_colnames <- colnames(work_table)
  
  for (index in 1:column_count) {
    
    worklist_colname <- worklist_colnames[index]
    
    identifiantFound <- FALSE
    identifiantFound <- str_detect(tolower(worklist_colname), "identifiant")
    
    IDFound <- FALSE
    IDFound <- str_detect(worklist_colname, "ID")
    
    worklist <- work_table[,index]
    worklist_with_unique_items <- unique(worklist)
    result1 <- length(worklist)
    result2 <- length(worklist_with_unique_items)
    
    if (((result1 == result2) && (result1 == row_count)) || identifiantFound==TRUE || IDFound==TRUE) {
      variable_set[[index]] <- "I"
    }else{
      if (result2 < 20 ) {
        variable_set[[index]] <- "D"
      }else{
        variable_set[[index]] <- "C"
      }
    }
  }
  
  return(variable_set)
}

index_first_column_with_continous_value <- function(work_table) {
  
  variable_set <- get_variable_set(work_table)
  variable_set_length <- length(variable_set)
  indexFound <- FALSE
  
  index <- 1
  while((index <= variable_set_length) && (indexFound==FALSE)) {
    if(variable_set[index] != "C") {
      index <- index + 1
    }else{
      indexFound = TRUE
    }
  }

  return (index)
}
  
get_tris_a_plat <- function(selected_sheet, work_table) {
  
  # Add header - "Tris à plat"
  worksheetWriteRowIndex <- 1
  xlsx.writeToCell(selected_sheet, rowIndex=worksheetWriteRowIndex, columnIndex=1, cellValue="Tris à plat", cellStyle = SPAD_WORKSHEET_TITLE_STYLE)
  
  # Start by adding the statistics of the first discrete value - eg."Type de Client"
  worksheetWriteRowIndex <- 4
  
  # Define the max value for the loop.
  table_column_max <- ncol(work_table)

  # Goes through the entire work_table and populate only the worksheet with discreet values.
  variable_set <- get_variable_set(work_table)
  
  for (table_column_index in 1:table_column_max) {
    
    if (variable_set[table_column_index] == "D") {
      
      # Add Table Elements
      worksheetWriteRowIndex <- xlsx.displayModality(sheet=selected_sheet, rowIndex=worksheetWriteRowIndex, working_table=work_table, working_table_column_num=table_column_index)
      worksheetWriteRowIndex <- worksheetWriteRowIndex + 1 
    }
  }
  
  xlsx::autoSizeColumn(selected_sheet, colIndex=1:4)
  
}


get_variables_continues <- function(selected_sheet, table_a_trier) {
  
  # Add header
  sheet_title <- "Statistiques sur variables continues"
  xlsx.writeToCell(selected_sheet, rowIndex=1, columnIndex=1, cellValue=sheet_title, cellStyle = SPAD_WORKSHEET_TITLE_STYLE)
  worksheetWriteRowIndex <- 3
  
  # Add Table Title
  table_title <- "Tableau"
  xlsx.writeToCell(selected_sheet, rowIndex=worksheetWriteRowIndex, columnIndex=1, cellValue=table_title, cellStyle = SPAD_TABLE_HEADER_STYLE)
  worksheetWriteRowIndex <- worksheetWriteRowIndex + 1
  
  # Add Decriptives Stats for each continuous variables
  table_column_max <- ncol(table_a_trier)
  table_column_names <- colnames(table_a_trier)

  first_column_with_continous_value_index <- index_first_column_with_continous_value(table_a_trier)
  
  varName <- table_column_names[first_column_with_continous_value_index]
  stat_score <- descriptives_stats_dataframe(varName, table_a_trier[,first_column_with_continous_value_index])
  stat_score_col_max <- ncol(stat_score)
  stat_score_colnames <- colnames(stat_score)

  # # Write a line with the table header (i.e. "Moyenne", "Ecart type", "Minimum", "Maximum", "Min2", "Max2", "Variance", etc...)
  xlsx.writeToCells(sheet=selected_sheet, rowIndex=worksheetWriteRowIndex, columnIndex=13, stat_score_colnames, SPAD_TABLE_TITLE_STYLE, SPAD_TABLE_TITLE_STYLE)
  worksheetWriteRowIndex <- worksheetWriteRowIndex + 1
  
  # # Goes through the entire work_table and populate only the worksheet with discreet values.
  variable_set <- get_variable_set(table_a_trier)

  for (table_column_index in 1:table_column_max) {

    if (variable_set[table_column_index] == "C") {

      varName <- table_column_names[table_column_index]
      stat_score <- descriptives_stats_list(varName, table_a_trier[,table_column_index])
      stat_score <- as.data.frame(stat_score)
      xlsx.writeToCells(sheet=selected_sheet, rowIndex=worksheetWriteRowIndex, columnIndex=13, stat_score, SPAD_TABLE_CELL_STYLE_ALIGN_RIGHT, SPAD_TABLE_CELL_STYLE_ALIGN_LEFT)
      worksheetWriteRowIndex <- worksheetWriteRowIndex + 1
    }
  }
  
  xlsx::autoSizeColumn(selected_sheet, colIndex = 1:13)
  
  #return(variable_set)
  return(NULL)
}

replace_modality_in_table <- function(a_table, column_num, pattern_for_madality, new_value) {
  
  if (column_num == 1) {
    a_table[, 1][a_table[, 1] == pattern_for_madality] <- new_value
  }
  else if (column_num == 2) {
    a_table[, 2][a_table[, 2] == pattern_for_madality] <- new_value
  }
  else if (column_num == 3) {
    a_table[, 3][a_table[, 3] == pattern_for_madality] <- new_value
  }
  else if (column_num == 4){
    a_table[, 4][a_table[, 4] == pattern_for_madality] <- new_value
  }
  else if (column_num == 5){
    a_table[, 5][a_table[, 5] == pattern_for_madality] <- new_value
  }
  else if (column_num == 6){
    a_table[, 6][a_table[, 6] == pattern_for_madality] <- new_value
  }
  else if (column_num == 7){
    a_table[, 7][a_table[, 7] == pattern_for_madality] <- new_value
  }
  else if (column_num == 8){
    a_table[, 8][a_table[, 8] == pattern_for_madality] <- new_value
  }
  else if (column_num == 9){
    a_table[, 9][a_table[, 9] == pattern_for_madality] <- new_value
  }
  else if (column_num == 10){
    a_table[, 10][a_table[, 10] == pattern_for_madality] <- new_value
  }
  else if (column_num == 11){
    a_table[, 11][a_table[, 11] == pattern_for_madality] <- new_value
  }
  else if (column_num == 12){
    a_table[, 12][a_table[, 12] == pattern_for_madality] <- new_value
  }
  else if (column_num == 13){
    a_table[, 13][a_table[, 13] == pattern_for_madality] <- new_value
  }
  else if (column_num == 14){
    a_table[, 14][a_table[, 14] == pattern_for_madality] <- new_value
  }
  else if (column_num == 15){
    a_table[, 15][a_table[, 15] == pattern_for_madality] <- new_value
  }
  else if (column_num == 16){
    a_table[, 16][a_table[, 16] == pattern_for_madality] <- new_value
  }
  else if (column_num == 17){
    a_table[, 17][a_table[, 17] == pattern_for_madality] <- new_value
  }
  else if (column_num == 18){
    a_table[, 18][a_table[, 18] == pattern_for_madality] <- new_value
  }
  else if (column_num == 19){
    a_table[, 19][a_table[, 19] == pattern_for_madality] <- new_value
  }  
  else if (column_num == 20){
    a_table[, 20][a_table[, 20] == pattern_for_madality] <- new_value
  }
  else {
    cat("Error this file has more than 20 columns !!!")
  }
  
  return(a_table)
}

