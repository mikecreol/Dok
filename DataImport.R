
filePath <- paste0("./Data/Stat v1.xlsx")

DataList <- list()
for(i in getSheetNames(filePath)){
  DataList[[i]] <- read.xlsx(filePath, sheet = i, startRow = 2)
}




