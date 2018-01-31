
filePath <- paste0("./Data/Stat-v3.xlsx")

DataList <- list()
for(i in getSheetNames(filePath)){
  DataList[[i]] <- read.xlsx(filePath, sheet = i, startRow = 2)
  legend_cols <- c(which(colnames(DataList[[i]]) == "Легенда:"):ncol(DataList[[i]]))
  DataList[[i]] <- DataList[[i]][,-legend_cols]
}




