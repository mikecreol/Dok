
# 3. Описателна статистика.(Процентно разпределение. Frequency tables.) За всички колони. 
#-------------------------------------------------------------------------------

# Temp <- DataList[[4]]
# Temp2 <- describe(Temp,
#                  na.rm = TRUE,
#                  skew = FALSE,
#                  quant=c(.25, .50, .75),
#                  IQR = TRUE)
# write.xlsx(Temp2, "Test.xlsx", rowNames = TRUE)

Data <- DataList[[4]]

Data$Пол <- factor(Data$Пол, labels = c("Ж", "М"))
Data$`Зъб.№` <- as.factor(Data$`Зъб.№`)
Data$Зъбна.група <- factor(Data$Зъбна.група, labels = c("фронт ГЧ", "премолари ГЧ", "молари ГЧ", "фронт ДЧ", "премолари ДЧ", "молари ДЧ"))
Data$Диагноза <- factor(Data$Диагноза, labels = c("хиперемия", "обратим пулпит", "необратим пулпит"))
# Data$спонтанна.болка <- replaceValues(Data$спонтанна.болка, NA, 0)
Data$спонтанна.болка <- as.factor(Data$спонтанна.болка)
Data$`продължителност.на.болката,.група` <- factor(Data$`продължителност.на.болката,.група`, 
                                                   labels = c("до 1 мин", "2-5 мин", "6-10 мин", "11-20 мин", "21-60 мин", ">61 мин"))
Data$честота.на.болката <- factor(Data$честота.на.болката, labels = c("1", "2-5", "6-10", ">=11"))
Data$Давност.на.болката <- factor(Data$Давност.на.болката, labels = c("1", "2-5", "6-10"))

Data$апроксимален.дефект <- factor(Data$апроксимален.дефект, labels = c("до 1/2", "> 1/2"))
Data$оклузален.дефект <- factor(Data$оклузален.дефект, labels = c("> 2/3")) # there are no cases with 1s (<2/3), so only 1 factor label setted
Data$цервикален.дефект <- factor(Data$цервикален.дефект, labels = c(" <3 мм", ">3 мм"))
Data$кариозна.маса <- factor(Data$кариозна.маса, labels = c("светла, мека", "тъмна, суха"))
Data$комуникация <- as.factor(Data$комуникация)
Data$Перкусия <- factor(Data$Перкусия, labels = c("болка"))
Data$Рентген <- factor(Data$Рентген, labels = c("б.о.", "разш.пероиод", "конкремент", "кариес в контакт с пулпата"))


A <- mapply(summaryDoc, Data[-1], colnames(Data)[-1])
listToExcel(A, filename = "./Output/SummaryStats v1.xlsx", sheetname = "3")


# 3.1. Доверителни интервали (референтни стойности) за ЕОД и ПО по диагноза(колона диагноза).
#-------------------------------------------------------------------------------

diagnoses <- split(Data, Data$Диагноза)

startrow <- 2

for (i in 1:3){
  data1 <- data.frame(diagnoses[[i]])
  listCI <- apply(data1[,c("ПО", "ЕОД")], 2, FUN = CI)
  CIdf <- rbind(listCI[[1]], listCI[[2]])
  CIdf$Label <- c("ПО", "ЕОД")
  CIdf <- CIdf[, c(3,1,2)]
  
  writeWorksheetToFile(file = "./Output/SummaryStats v1.xlsx", CIdf, sheet = "3.1", startRow = startrow)
  startrow = 4 + nrow(CIdf) + startrow
}


# 3.2. Средни стойности на ЕОД и ПО по диагнози сравнени с определените стойности от 2.2. T test
#-------------------------------------------------------------------------------

# read data 2.2
Data22 <- fillNAsWithUpperCell(DataList[[2]])
Data22$out_of_range <- ifelse(Data22$ПО < 75 | Data22$ПО > 96, 1, 0)
Intaktni <- Data22[Data22$out_of_range == 0,]

startrow <- 2

for (i in 1:3){
  data1 <- data.frame(diagnoses[[i]])

  t1 <- t.test(Intaktni$ПО, data1$ПО, paired = FALSE, alt = "two.sided", conf.level = 0.95)
  t2 <- t.test(Intaktni$ЕОД, data1$ЕОД, paired = FALSE, alt = "two.sided", conf.level = 0.95)
  
  PO_ttest <- ttest_res(t1)
  PO_ttest$comparison <- "ПО"
  cols <- colnames(PO_ttest)
  PO_ttest <- cbind(PO_ttest[, ncol(PO_ttest)], PO_ttest[,-ncol(PO_ttest)])
  colnames(PO_ttest) <- c(cols[length(cols)], cols[-length(cols)])
  
  EOD_ttest <- ttest_res(t2)
  EOD_ttest$comparison <- "ЕОД"
  cols <- colnames(EOD_ttest)
  EOD_ttest <- cbind(EOD_ttest[, ncol(EOD_ttest)], EOD_ttest[,-ncol(EOD_ttest)])
  colnames(EOD_ttest) <- c(cols[length(cols)], cols[-length(cols)])
  
  writeWorksheetToFile(file = "./Output/SummaryStats v1.xlsx", PO_ttest, sheet = "3.2", startRow = startrow)
  startrow = 3 + nrow(PO_ttest) + startrow
  writeWorksheetToFile(file = "./Output/SummaryStats v1.xlsx", EOD_ttest, sheet = "3.2", startRow = startrow)
  startrow <- 4 + nrow(EOD_ttest) + startrow
  
}



