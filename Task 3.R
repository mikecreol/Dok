
# 3. Описателна статистика. За всички колони. Доверителни интервали (референтни стойности) за ЕОД
# и ПО по диагноза(колона диагноза). Средни стойности на ЕОД и ПО по диагнози сравнени с определените стойности от 2.2.
#-------------------------------------------------------------------------------

Temp <- DataList[[4]]
Temp2 <- describe(Temp,
                 na.rm = TRUE,
                 skew = FALSE,
                 quant=c(.25,.75),
                 IQR = TRUE)
write.xlsx(Temp2, "Test.xlsx", rowNames = TRUE)

A <- mapply(summaryDoc, Data, colnames(Data))
listToExcel(A, filename = "./Output/SummaryStats v1.xlsx", rowNames = FALSE)

