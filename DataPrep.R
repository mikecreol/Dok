
Temp <- fillNAsWithUpperCell(DataList[[1]])


A <- mapply(summaryDoc, Temp, colnames(Temp))
listToExcel(A, filename = "./Output/SummaryStats v1.xlsx", rowNames = FALSE)

colnames(Data) <- stri_trans_general(colnames(Data), 'cyrillic-latin')

