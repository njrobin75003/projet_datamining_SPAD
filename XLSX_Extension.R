source("Globals.R")

# XLSLX_Extension.R

#------------------------------------------------------------------------------------------------
# Default font definition
#------------------------------------------------------------------------------------------------
SPAD_WORKSHEET_FONT <- "Tahoma"

#------------------------------------------------------------------------------------------------
# Create an Excel Workbook
#------------------------------------------------------------------------------------------------
workbook <- createWorkbook(type="xlsx")

#------------------------------------------------------------------------------------------------
# Worksheet Cell Styles
#------------------------------------------------------------------------------------------------

# SPAD-like Worksheet Title Style - TAHOMA 16 bold
SPAD_WORKSHEET_TITLE_STYLE <- CellStyle(workbook) + Font(workbook, name = SPAD_WORKSHEET_FONT, heightInPoints=16)

# SPAD-like Worksheet Header Style - TAHOMA 11 bold
SPAD_TABLE_HEADER_STYLE <- CellStyle(workbook) + Font(workbook,  name = SPAD_WORKSHEET_FONT, heightInPoints=11, isBold=TRUE)

# SPAD-like Table Header Style - TAHOMA 10 bold
SPAD_TABLE_TITLE_STYLE <- CellStyle(workbook) + Font(workbook,  name = SPAD_WORKSHEET_FONT, heightInPoints=10, isBold=TRUE) +
  Fill(foregroundColor="lightgray", backgroundColor="white", pattern="SOLID_FOREGROUND") +
  Alignment(wrapText=TRUE, horizontal="ALIGN_CENTER") +
  Border(color="black", position=c("TOP", "BOTTOM", "LEFT", "RIGHT"), 
         pen=c("BORDER_THIN", "BORDER_THIN","BORDER_THIN", "BORDER_THIN"))

# SPAD-like Table Cell Style - TAHOMA 10 bold - ALIGN_LEFT
SPAD_TABLE_CELL_STYLE_ALIGN_LEFT <- CellStyle(workbook) + Font(workbook,  name = SPAD_WORKSHEET_FONT, heightInPoints=10) +
  Alignment(wrapText=TRUE, horizontal="ALIGN_LEFT") +
  Border(color="black", position=c("TOP", "BOTTOM", "LEFT", "RIGHT"), 
         pen=c("BORDER_THIN", "BORDER_THIN","BORDER_THIN", "BORDER_THIN")) 

# SPAD-like Table Cell Style - TAHOMA 10 bold - ALIGN_RIGHT
SPAD_TABLE_CELL_STYLE_ALIGN_RIGHT <- CellStyle(workbook) + Font(workbook,  name = SPAD_WORKSHEET_FONT, heightInPoints=10) +
  Alignment(wrapText=TRUE, horizontal="ALIGN_RIGHT") +
  Border(color="black", position=c("TOP", "BOTTOM", "LEFT", "RIGHT"), 
         pen=c("BORDER_THIN", "BORDER_THIN","BORDER_THIN", "BORDER_THIN")) 

#------------------------------------------------------------------------------------------------
# Fonctions
#------------------------------------------------------------------------------------------------

# Function that writes the content of a single cell into the selected sheet
xlsx.writeToCell<-function(sheet, rowIndex, columnIndex=1, cellValue, cellStyle){
  rows <-createRow(sheet, rowIndex)
  cell <-createCell(rows, columnIndex)
  setCellValue(cell[[1,1]], cellValue)
  xlsx::setCellStyle(cell[[1,1]], cellStyle)
}

# Function that writes the content of several cells into the selected sheet
# on the selected row (row index)
# from col 1 to col n (n = colIndex)
xlsx.writeToCells<-function(sheet, rowIndex, columnIndex=1, listValue, cellStyle1=NULL, cellStyle2=NULL){
  rows <-createRow(sheet, rowIndex)
  cells <- createCell(rows, colIndex=1:columnIndex)
  
  columnLoopIndex <- 1
  for (element in listValue) {
    setCellValue(cells[[1,columnLoopIndex]], element)
    if (columnLoopIndex == 1) {
        xlsx::setCellStyle(cells[[1,columnLoopIndex]], cellStyle2)
    }else{
      xlsx::setCellStyle(cells[[1,columnLoopIndex]], cellStyle1)
    }
    columnLoopIndex <- columnLoopIndex + 1
  }
}

# Function that writes the content of several cells into the selected sheet
# on the selected row (row index)
# from col 1 to col n (n = colIndex)
xlsx.writeToCellsFromDF<-function(sheet, rowIndex, startingPoint=1, columnIndex=1, dataFrame, dataFrameRowIndex, cellStyle) {
  
  rows <-createRow(sheet, rowIndex)
  cells <- createCell(rows, colIndex=startingPoint:columnIndex)
  
  # Pre-allocate a list and fill it with a loop
  dataFrameRow <- dataFrame[dataFrameRowIndex,]
  myListe <- c()
  for (i in 1:ncol(dataFrameRow)) {
    myListe[i] <- as.character(dataFrameRow[i])
  }
  
  # Write the content of the list in the worksheet
  columnLoopIndex <- 1
  for (element in myListe) {
    setCellValue(cells[[1,columnLoopIndex]], element)
    if (columnLoopIndex == 1) {
      xlsx::setCellStyle(cells[[1,columnLoopIndex]], SPAD_TABLE_CELL_STYLE_ALIGN_LEFT)
    }else{
      xlsx::setCellStyle(cells[[1,columnLoopIndex]], cellStyle)
    }
    columnLoopIndex <- columnLoopIndex + 1
  }
}

# This fiunction retrives the modalities for one given column
xlsx.displayModality <- function(sheet, rowIndex, working_table, working_table_column_num) {
  
  # Add  Modality Title - e.g. Type de Client
  modality_Title <- colnames(working_table[working_table_column_num])
  worksheetWriteRowIndex <- rowIndex
  xlsx.writeToCell(sheet, rowIndex=worksheetWriteRowIndex, columnIndex=1, cellValue=modality_Title, cellStyle = SPAD_TABLE_HEADER_STYLE)
  worksheetWriteRowIndex <- worksheetWriteRowIndex + 1
  
  # Add Table Title  - e.g. "Modalités" - "Effectifs" - "Pourcentage" - "% sur exprimés"
  tableRowTitleElements <- c("Modalités", "Effectifs", "Pourcentage", "% sur exprimés")
  xlsx.writeToCells(sheet=sheet, rowIndex=worksheetWriteRowIndex, columnIndex=4, listValue=tableRowTitleElements, cellStyle1=SPAD_TABLE_TITLE_STYLE, cellStyle2=SPAD_TABLE_TITLE_STYLE)
  worksheetWriteRowIndex <- worksheetWriteRowIndex + 1
  
  # Retrieve the column number of the table where the discrete values are.
  column_num <- working_table_column_num

  # Select the modalities of that column
  modalities <- working_table %>%
    group_by(working_table[ , column_num]) %>%
    summarise(Effectifs = n()) %>%
    mutate(Pourcentages = Effectifs/sum(Effectifs)*100, '% sur exprimés' = Effectifs/sum(Effectifs)*100)
  
  rounded_modalities <- working_table %>%
    group_by(working_table[ , column_num]) %>%
    summarise(Effectifs = n()) %>%
    mutate(Pourcentages = round(Effectifs/sum(Effectifs)*100,1), '% sur exprimés' = round(Effectifs/sum(Effectifs)*100,1))

  # Get the number of inividuals and frequence for each modality
  ensemble <- modalities %>%
    summarise(Ensemble = "Ensemble",
              Effectifs = sum(modalities$Effectifs),
              Pourcentages = format(round(sum(modalities$Pourcentages),1), nsmall = 1),
              '% sur exprimés' = format(round(sum(modalities$`% sur exprimés`),1), nsmall = 1))
  
  # Returns the number of modalities
  modalitiesRowCount <- nrow(modalities)
  # Returns the number of columns (ie. 4) we need in order to write the modality for one given variable.
  modalitiesColCount <- ncol(modalities)

  # Write the rounded result values (i.e. "Effectifs", "Pourcentages", "% sur exprimés") for each modaliy.
  for (rowIndex in 1:modalitiesRowCount) {
    
    xlsx.writeToCellsFromDF(sheet, rowIndex=worksheetWriteRowIndex, columnIndex=4, dataFrame=rounded_modalities, dataFrameRowIndex=rowIndex, cellStyle=SPAD_TABLE_CELL_STYLE_ALIGN_RIGHT)
    worksheetWriteRowIndex <- worksheetWriteRowIndex + 1
  }
  
  # Write an empty line with borders.
  empty_line <- c("", "", "", "")
  xlsx.writeToCells(sheet, rowIndex=worksheetWriteRowIndex, columnIndex=4, listValue=empty_line, cellStyle1=SPAD_TABLE_CELL_STYLE_ALIGN_RIGHT, cellStyle2=SPAD_TABLE_CELL_STYLE_ALIGN_RIGHT)
  worksheetWriteRowIndex <- worksheetWriteRowIndex + 1
  
  # Write the line Ensemble (total for each column).
  xlsx.writeToCellsFromDF(sheet, rowIndex=worksheetWriteRowIndex, startingPoint=1, columnIndex=4, dataFrame=ensemble, dataFrameRowIndex=1, cellStyle=SPAD_TABLE_CELL_STYLE_ALIGN_RIGHT)
  worksheetWriteRowIndex <- worksheetWriteRowIndex + 1
  
  return(worksheetWriteRowIndex)
}
