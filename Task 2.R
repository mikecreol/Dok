
# Определяне на доверителен интервал за здрави зъби на резултатите от ПО(Пулсова оксиметрия). 
# Min и Max, Квантили, Хистограма. От таблица 2.2.
#-------------------------------------------------------------------------------

Data <- fillNAsWithUpperCell(DataList[[2]])
# View(describe(Data))
Data$Пол <- factor(Data$Пол, labels = c("Ж", "М"))
Data$Пушач <- as.factor(Data$Пушач)
Data$`Зъбен.№` <- as.factor(Data$`Зъбен.№`)
Data$Зъбна.група <- factor(Data$Зъбна.група, labels = c("фронт ГЧ", "премолари ГЧ", "молари ГЧ", "фронт ДЧ", "премолари ДЧ", "молари ДЧ"))

# дескриптивна статистика, здрави зъби (всички пром.)
A <- mapply(summaryDoc, Data[-1], colnames(Data)[-1])
listToExcel(A, filename = "./Output/SummaryStats v1.xlsx", sheetname = "2")

# ПО - дескриптивна статистика, хистограма
POsummary <- summaryDoc(Data$ПО, varName="ПО")
writeWorksheetToFile(file = "./Output/SummaryStats v1.xlsx", POsummary, sheet = "2_ПО", startRow = 1)

writeWorksheetToFile(file = "./Output/SummaryStats v1.xlsx", CIquantile(Data$ПО, 5), sheet = "2_ПО", startRow = 20)

plot_obj <- ggplot(data=Data, aes(Data$ПО)) + 
  geom_histogram(breaks=seq(8, 101, by=3), 
                 col="green", 
                 fill="blue", 
                 alpha = .8) + 
  labs(title="Хистограма на пулсова оксимерия") +
  labs(x="Пулсова оксиметрия", y="Честота")

# wb <- loadWorkbook("./Output/SummaryStats v1.xlsx")
# xlsArea <- "!$E$3"
# imageArea <- paste0("2_ПО", xlsArea)
# createName(wb, name = "hist", formula = imageArea, overwrite = TRUE)# Create a named region
ggsave(file="./Output/histogram_PO.PNG") # Create R plot
# addImage(wb, filename = "./Output/histogram_PO.PNG", name = "hist", originalSize = TRUE) # Write image to the named region created above


# 2.1. ПО и ЕОД изследване, преди и след изолиране със силикон. T-Test. Ще установим, че няма разлика.
#-------------------------------------------------------------------------------

Data <- fillNAsWithUpperCell(DataList[[1]])

par(mfrow=c(2,2))
boxplot(Data$ПО.преди)
boxplot(Data$ПО.след)
boxplot(Data$ЕОД.преди)
boxplot(Data$ЕОД.след)


boxplot.stats(Data$ПО.преди)
boxplot.stats(Data$ПО.след)
boxplot.stats(Data$ЕОД.преди)
boxplot.stats(Data$ЕОД.след)

# ПО преди и след
t1 <- t.test(Data$ПО.преди, Data$ПО.след, paired = FALSE, alt = "two.sided", conf.level = 0.95) # p-value = 0.02013
# t2 <- t.test(Data$ПО.преди, Data$ПО.след, paired = FALSE, alt = "greater", conf.level = 0.95) # p-value = 0.01007
PO_two.sided <- ttest_res(t1)
# PO_greater <- ttest_res(t2)
PO_two.sided <- MyHelperFunctions::myRename(PO_two.sided, c("mean.x", "mean.y"), c("MeanPO_before", "MeanPO_after"))

# ЕОД преди и след
t3 <- t.test(Data$ЕОД.преди, Data$ЕОД.след, paired = FALSE, alt = "two.sided", conf.level = 0.95) # p-value = 0.01682
# t4 <- t.test(Data$ЕОД.преди, Data$ЕОД.след, paired = FALSE, alt = "greater", conf.level = 0.95) # p-value = 0.008408
EOD_two.sided <- ttest_res(t3)
# EOD_greater <- ttest_res(t4)
EOD_two.sided <- MyHelperFunctions::myRename(EOD_two.sided, c("mean.x", "mean.y"), c("MeanEOD_before", "MeanEOD_after"))

startrow <- 2
writeWorksheetToFile(file = "./Output/SummaryStats v1.xlsx", PO_two.sided, sheet = "2.1", startRow = startrow)
startrow = 4 + nrow(PO_two.sided) + startrow
# writeWorksheetToFile(file = "./Output/SummaryStats v1.xlsx", PO_greater, sheet = "2.1", startRow = startrow)
# startrow = 4 + nrow(PO_greater) + startrow
writeWorksheetToFile(file = "./Output/SummaryStats v1.xlsx", EOD_two.sided, sheet = "2.1", startRow = startrow)
# startrow = 4 + nrow(EOD_two.sided) + startrow
# writeWorksheetToFile(file = "./Output/SummaryStats v1.xlsx", EOD_greater, sheet = "2.1", startRow = startrow)


# 2.2. Mean, SD, SE, по групи зъби. Корелация между ЕОД и ПО по групи (в колона зъбна група). 
# Под 75 и над 96 на ПО да ги махна. И Tukey HSD пак по групи.
#-------------------------------------------------------------------------------

Data <- fillNAsWithUpperCell(DataList[[2]])

Data[,"Зъбна.група"] <- factor(Data[,"Зъбна.група"], 
                                      labels = c("фронт ГЧ", "премолари ГЧ", "молари ГЧ", "фронт ДЧ", "премолари ДЧ", "молари ДЧ"))

# boxplot(Data$ПО)
# sum(Data$ПО < 75) # 299
# sum(Data$ПО > 96) # 88

Data$out_of_range <- ifelse(Data$ПО < 75 | Data$ПО > 96, 1, 0)
ReducedData <- Data[Data$out_of_range == 0,]

# descr_all <- describe(Data$ПО, na.rm=T, skew=F, quant = c(.25, .50, .75), IQR=T) # All data, before removing <75 and >96

# Mean, sd, se, min, max, q25, median, q75 - Total and by groups
descr_total <- describe(ReducedData$ПО, na.rm=T, skew=F, quant = c(.25, .50, .75), IQR=T)
row.names(descr_total) <- c("Total")
descr_groups <- do.call(rbind, lapply(split(ReducedData, ReducedData$Зъбна.група),
                                      function(x) describe(x$ПО, na.rm=T, skew=F, quant = c(.25, .50, .75), IQR=T)))
descrData <- rbind(descr_groups, descr_total)
descrData$vars <- row.names(descrData)
descrData <- descrData[, c("vars", "n", "mean", "sd", "se", "min", "Q0.25", "Q0.5", "Q0.75", "max", "range", "IQR")]
colnames(descrData) <- c("Labels", colnames(descrData)[-1])


# Correlations - Total and by groups
cor_total <- cor(ReducedData[, c("ПО", "ЕОД")])
cor_total <- data.frame(Зъбна.група = c("Total"), Correlation = cor_total[1,2])
cor_groups <- do.call(rbind, lapply(split(ReducedData, ReducedData$Зъбна.група),
                                   function(x) data.frame(Зъбна.група=x$Зъбна.група[1], Correlation=cor(x$ПО, x$ЕОД))))
corData <- rbind(cor_groups, cor_total)


# Tukey HSD  - by groups

TukeyPO <- tukeyHSD.ByGroup(var = "ПО", groups = "Зъбна.група", ReducedData)

startrow <- 1
writeWorksheetToFile(file = "./Output/SummaryStats v1.xlsx", descrData, sheet = "2.2", startRow = startrow)
startrow <- startrow + nrow(descrData) + 4
writeWorksheetToFile(file = "./Output/SummaryStats v1.xlsx", "Корелация между ПО и ЕОД (по зъбна група)", sheet = "2.2", startRow = startrow)
startrow <- startrow + 2
writeWorksheetToFile(file = "./Output/SummaryStats v1.xlsx", corData, sheet = "2.2", startRow = startrow)
startrow <- startrow + nrow(corData) + 4
writeWorksheetToFile(file = "./Output/SummaryStats v1.xlsx", TukeyPO, sheet = "2.2", 
                     startRow = startrow)

