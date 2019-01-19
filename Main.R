
source("Globals.R")
source("XLSX_Extension.R")
source("Functions.R")

# Assign the working directory.
setwd(Selected_Working_Directory)

# User selects the file.
fileToAnalyze <- getDataFilePath()
  
# Import the data file into a table.
if (strEndsWith(fileToAnalyze, "xls") == TRUE) {
  table_to_analyze <- read.xlsx(fileToAnalyze, sheetIndex=1, header=TRUE, stringsAsFactors=FALSE, encoding="UTF-8")
}else{
  table_to_analyze <- fread(fileToAnalyze, sep="|", header = TRUE, dec = ",", stringsAsFactors = FALSE)
  table_to_analyze <- as.data.frame(table_to_analyze)
}

# --------------------------------------------------------------------------------------------
# Step 1 - Write descriptives statistisques
# --------------------------------------------------------------------------------------------

# Add the first Excel spreadsheet in order to export the statistics of the discrete variables.
sheet1 <- xlsx::createSheet(workbook, sheetName = "Tris à plats")

# Export the statistics of the discrete variables.
get_tris_a_plat(sheet1, table_to_analyze)

# Add the second Excel spreadsheet in order to export the statistics of the continous variables.
sheet2 <- xlsx::createSheet(workbook, sheetName = "Variables continues")

# Export the statistics of the continous variables.
score <- get_variables_continues(sheet2, table_to_analyze)

xlsx::saveWorkbook(workbook, "Statistiques descriptives.xlsx")

# DESCRIPTION DES VARIABLES QUALITATIVES

# Calcul des effectifs
#effectifs=table(table_to_analyze$Situation.familiale,useNA = "always")
# Calcul des fréqueces
#frequences=prop.table(effectifs)
# Concaténation des resultats
#View(cbind(effectifs, frequences))
#il suffit donc de changer le nom de la variable pour en obtenir les stats

# An interesting library : DT.
#datatable(table_to_analyze_dt)
# x[i] = x[[i]] = x$nom
# exemple : my.df$V2[my.df$V2 == "-sh2"] <- -100
#hhhhhh <- table_to_analyze_dt[,2]




# --------------------------------------------------------------------------------------------
# Step 2 - Consolidate modalities
# --------------------------------------------------------------------------------------------

table_to_analyze_2 <- copy(table_to_analyze)

# First the user has to select the columns where he/she wants to see the modalities grouped.

# Collect the names of all the columns
table_column_names <- colnames(table_to_analyze_2)

# Let the user select the columns in the console.
selected_colums <- dlg_list(table_column_names, multiple = TRUE)$res

# Columns are selected now.
number_of_selected_columns <- length(selected_colums)

for (column_index in 1:number_of_selected_columns) {
  
  # The program will now look for the modalities of the selected column.
  column_element <- selected_colums[column_index]
  column_value <- which(colnames(table_to_analyze)==column_element )
  
  new_group <- table_to_analyze_2 %>% 
    group_by_(column_element) %>% 
    summarise(n())
  
  possible_options_list <- as.list(new_group[1])
  possible_options <- possible_options_list[[1]]
  
  # User can now select the modalities he/she wants to group.
  selected_modalites <- dlg_list(possible_options, multiple = TRUE)$res
  
  number_of_selected_modalities <- length(selected_modalites)
  
  grouped_modality <- ""
  for (index in 1:number_of_selected_modalities) {
    if (index == 1) {
      grouped_modality <- paste0(selected_modalites[index])
    }else{
      grouped_modality <- paste0(grouped_modality, "/", selected_modalites[index])
    }
  }
  
  for (modality_index in 1:number_of_selected_modalities) {
    selected_pattern <- selected_modalites[modality_index]
    table_to_analyze_2 <- replace_modality_in_table(table_to_analyze_2, column_value, selected_pattern, grouped_modality)
  }
}

# #Regroupement de modalites sur les variables situation_familiale et domiciliation_de_lepargne
# table_to_analyze$Situation.familiale<-recode(table_to_analyze$Situation.familiale, 
#                                         "célibataire" = "célibataire",
#                                         "divorcé" = "divorcé/veuf",
#                                         "marié" = "marié",
#                                         "veuf" = "divorcé/veuf")
# 
# base_credit$Domiciliation.de.l.épargne<-recode(table_to_analyze$Domiciliation.de.l.épargne,
#                                               "moins de 10K épargne"="moins de 10K épargne", "pas d'épargne"="pas d'épargne",
#                                               "de 10 à 100K épargne"="plus de 10K épargne", "plus de 100K épargne"="plus de 10K épargne")



# Add the first Excel spreadsheet in order to export the statistics of the discrete variables.
sheet1 <- xlsx::createSheet(workbook, sheetName = "Tris à plats")

# Export the statistics of the discrete variables.
get_tris_a_plat(sheet1, table_to_analyze)

xlsx::saveWorkbook(workbook, "Regroupement des Modalités.xlsx")

# --------------------------------------------------------------------------------------------
# Step 3 - Logistique Regression
# --------------------------------------------------------------------------------------------

#REGRESSION LOGISTIQUE SUR LA NOUVELLE BASE
#Discrétisation de la variable à expliquer

base_credit$Type.de.client<- recode(base_credit$Type.de.client, "Bon client"= 1,
                                    "Mauvais client"= 0)
#On cree nos echantillons de test et d'apprentissage
ind <- sample(2, nrow(base_credit), replace=T, prob=c(0.75,0.25)) tdata<- base_credit[ind==1,] # training = 75%
vdata<- base_credit[ind==2,] #validation = 25%
#choix des modalités de reference
base_credit$Age.du.client <- relevel(base_credit$Age.du.client, ref = "moins de 23 ans")
base_credit$Situation.familiale <- relevel(base_credit$Situation.familiale, ref = "divorcé/veuf")
base_credit$Ancienneté <- relevel(base_credit$Ancienneté, ref = "anc. 1 an ou moins")
base_credit$Domiciliation.du.salaire <- relevel(base_credit$Domiciliation.du.salaire, ref = "Non domicilié")
base_credit$Domiciliation.de.l.épargne <- relevel(base_credit$Domiciliation.de.l.épargne, ref = "pas d'épargne")
base_credit$Profession <- relevel(base_credit$Profession, ref = "cadre") base_credit$Moyenne.encours <- relevel(base_credit$Moyenne.encours, ref = "plus de 5 K encours")
base_credit$Moyenne.des.mouvements <- relevel(base_credit$Moyenne.des.mouvements, ref = "de 10 à 30K mouvt")
base_credit$Cumul.des.débits <- relevel(base_credit$Cumul.des.débits, ref = "plus de 100 débits")
base_credit$Autorisation.de.découvert <- relevel(base_credit$Autorisation.de.découvert, ref = "découvert autorisé")
base_credit$Interdiction.de.chéquier <- relevel(base_credit$Interdiction.de.chéquier, ref = "chéquier interdit")
#estimation du modèle avec les données d'entrainement fit.glm = glm(tdata[,2]~.,data=tdata[,3:15],family=binomial) summary(fit.glm)
#prédiction avec les données de test

score.glm = predict(fit.glm, vdata[,3:15],type="response") sum(as.numeric(predict.glm(fit.glm,vdata[,3:15],type="response")>=0.5)) class.glm=as.numeric(predict.glm(fit.glm,vdata[,3:15],type="response")>=0.5) table(class.glm,vdata[,2])
sum( score.glm >= 0.5 & vdata[,2]==1)/sum(vdata[,2]==1)#proportion de bien classés (bon clients)
sum( score.glm < 0.5 & vdata[,2]==0)/sum(vdata[,2]==0)#Proportion de bien classés (mauvais clients)
sum( score.glm2 >= 0.5 & tdata[,2]==1)/sum(tdata[,2]==1)#proportion de bien classés (bon clients)
sum( score.glm2 < 0.5 & tdata[,2]==0)/sum(tdata[,2]==0)#Proportion de bien classés (mauvais clients)
s=quantile(score.glm,probs=seq(0,1,0.01)) PVP = rep(0,length(s))
PFP = rep(0,length(s))
for (i in 1:length(s)){
  PVP[i]=sum(score.glm>=s[i]& vdata[,2]==1)/sum(vdata[,2]==1)
  PFP[i]=sum( score.glm >=s[i] & vdata[,2]==0)/sum(vdata[,2]==0) }
plot(PFP,PVP,type="l",col="red")
