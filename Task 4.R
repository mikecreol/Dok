
# 4. Описателна статистика. (Процентно разпределение. Frequency tables.) 
#-------------------------------------------------------------------------------

# 4.травми -> Data1
Data1 <- read.xlsx(filePath, sheet = "4 - травми", startRow = 5)
Data1 <- Data1[, -c(37:38)]

Data1[,c(1:5)] <- fillNAsWithUpperCell(Data1[,c(1:5)])
# Do columns (4,5) (Посока, Време след) need to be filled with upper cell???

Data1$Пол <- factor(Data1$Пол, labels = c("Ж", "М"))
Data1$Посока <- factor(Data1$Посока, labels = c("отдолу", "фронтално", "отгоре", "странично"))
Data1$Цвят.на.пулпата <- factor(Data1$Цвят.на.пулпата, labels = c("розов"))
Data1$Кореново.развитие <- factor(Data1$Кореново.развитие, labels = c("незавършено", "завършено"))
Data1$Стадий.на.корен.разв. <- factor(Data1$Стадий.на.корен.разв., labels = c("до 2/3", "изграден корен с отворен апекс", "завършено"))
Data1$flag <- replaceValues(Data1$flag, NA, 0)

num.columns <- c("Пациент.номер", "Възраст", "Пулпно.разкритие", 
                 "ЕОД.1д", "ЕОД.15д", "ЕОД.30д", "ЕОД.3м", "ЕОД.6м", 
                 "ПО.1д", "ПО.15д", "ПО.30д", "ПО.3м", "ПО.6м", "flag")
cols.with.labels <- c("Пол", "Посока", "Цвят.на.пулпата", "Кореново.развитие", "Стадий.на.корен.разв.")
Data1[, which(!colnames(Data1) %in% c(num.columns, cols.with.labels))] <- apply(Data1[, which(!colnames(Data1) %in% c(num.columns, cols.with.labels))], 
                                                                              2, function(x) {as.factor(x)})

# 4.референтни -> Data2
Data2 <- DataList[[6]]
Data2[,c(1:3)] <- fillNAsWithUpperCell(Data2[,c(1:3)])

Data2$Пол <- factor(Data2$Пол, labels = c("Ж", "М"))
Data2$Зъбен.номер <- as.factor(Data2$Зъбен.номер)
Data2$корен.развитие <- factor(Data2$корен.развитие, labels = c("незавършено", "завършено"))
Data2$стадий.на.кор.разв. <- factor(Data2$стадий.на.кор.разв., labels = c("до 2/3", "изграден корен с отворен апекс", "завършено"))

# descriptive statistics (Травми, Референтни)
A <- mapply(summaryDoc, Data1[,-c(1,ncol(Data1))], colnames(Data1[, -c(1,ncol(Data1))]))
listToExcel(A, filename = "./Output/SummaryStats v1.xlsx", sheetname = "4.травми-табл")

A <- mapply(summaryDoc, Data2[-1], colnames(Data2[-1]))
listToExcel(A, filename = "./Output/SummaryStats v1.xlsx", sheetname = "4.референтни-табл")


# 4.1. Графика на средните по периоди (В динамика) 
#-------------------------------------------------------------------------------
Data1 <- Data1[!(Data1$flag == 1), ]

plotData <- Data1[, c("ЕОД.1д", "ЕОД.15д", "ЕОД.30д", "ЕОД.3м", "ЕОД.6м", "ПО.1д", "ПО.15д", "ПО.30д", "ПО.3м", "ПО.6м")]
EODdata <- melt(data = plotData[,c(1:5)], measure.vars = c(1:5), variable.name = "Period", value.name = "EOD")
EODdata[,"Period"] <- gsub("[[:punct:]]", "", EODdata[,"Period"])
EODdata[,"Period"] <- gsub("[ЕОД]", "", EODdata[,"Period"])
EODdata[,"Period"] <- replaceValues(EODdata$Period, c("1д", "15д", "30д", "3м", "6м"), c(1,2,3,4,5))
# EODdata$Period <- as.numeric(EODdata$Period)
EODdata$Period <- factor(EODdata$Period, labels = c("1д", "15д", "30д", "3м", "6м"))

POdata <- melt(data = plotData[,c(6:10)], measure.vars = c(1:5), variable.name = "Period", value.name = "PO")
POdata[,"Period"] <- gsub("[[:punct:]]", "", POdata[,"Period"])
POdata[,"Period"] <- gsub("[ПО]", "", POdata[,"Period"])
POdata[,"Period"] <- replaceValues(POdata$Period, c("1д", "15д", "30д", "3м", "6м"), c(1,2,3,4,5))
# POdata$Period <- as.numeric(POdata$Period)
POdata$Period <- factor(POdata$Period, labels = c("1д", "15д", "30д", "3м", "6м"))

EODmeans <- plotmeans(EOD ~ Period, data = EODdata, 
                      mean.labels = TRUE, 
                      # frame = FALSE, 
                      main = "Динамика на средните на ЕОД по периоди", 
                      xlab = "Период", 
                      ylab = "ЕОД")
# png(filename = "./Output/EODmeans.png")
# plot(EODmeans, type = "simple")
# dev.off()

POmeans <- plotmeans(PO ~ Period, data = POdata, 
                      mean.labels = TRUE, 
                      # frame = FALSE, 
                      main = "Динамика на средните на ПО по периоди", 
                      xlab = "Период", 
                      ylab = "ПО")
# png(filename = "./Output/POmeans.png")
# plot(POmeans, type = "simple")
# dev.off()


# @ DIDO: Vij gi tiq grafiki, neshto ne sa ok.
# Also - How to save plots, it doesn't work, I have exported them from the menu.


# 4.2. Сравнение на средни по периоди и групирани по диагноза пак по периоди (1д и тн.). 
# (може да махнем тези които не реагират; 5 оранжеви реда)
#-------------------------------------------------------------------------------

##### Comparisons by Periods with TUKEY
TukeyEODbyPeriods <- tukeyHSD.ByGroup(var = "EOD", groups = "Period", EODdata)
TukeyPObyPeriods <- tukeyHSD.ByGroup(var = "PO", groups = "Period", POdata)


##### Comparisons by Periods with Ttests

t_EOD1 <- t.test(plotData$ЕОД.1д, plotData$ЕОД.15д, paired = TRUE, 
                           alt = "two.sided", conf.level = 0.95)
t_EOD2 <- t.test(plotData$ЕОД.15д, plotData$ЕОД.30д, paired = TRUE, 
                 alt = "two.sided", conf.level = 0.95)
t_EOD3 <- t.test(plotData$ЕОД.30д, plotData$ЕОД.3м, paired = TRUE, 
                 alt = "two.sided", conf.level = 0.95)
t_EOD4 <- t.test(plotData$ЕОД.3м, plotData$ЕОД.6м, paired = TRUE, 
                 alt = "two.sided", conf.level = 0.95)

t_PO1 <- t.test(plotData$ПО.1д, plotData$ПО.15д, paired = TRUE, 
                 alt = "two.sided", conf.level = 0.95)
t_PO2 <- t.test(plotData$ПО.15д, plotData$ПО.30д, paired = TRUE, 
                 alt = "two.sided", conf.level = 0.95)
t_PO3 <- t.test(plotData$ПО.30д, plotData$ПО.3м, paired = TRUE, 
                 alt = "two.sided", conf.level = 0.95)
t_PO4 <- t.test(plotData$ПО.3м, plotData$ПО.6м, paired = TRUE, 
                 alt = "two.sided", conf.level = 0.95)


# output results
# EOD comparisons by period
t_EOD1_res <- ttest_res(t_EOD1)
t_EOD1_res$Label <- "Сравняване средното на ЕОД (1д - 15д)"
cols <- colnames(t_EOD1_res)
t_EOD1_res <- cbind(t_EOD1_res[, ncol(t_EOD1_res)], t_EOD1_res[,-ncol(t_EOD1_res)])
colnames(t_EOD1_res) <- c(cols[length(cols)], cols[-length(cols)])

t_EOD2_res <- ttest_res(t_EOD2)
t_EOD2_res$Label <- "Сравняване средното на ЕОД (15д - 30д)"
cols <- colnames(t_EOD2_res)
t_EOD2_res <- cbind(t_EOD2_res[, ncol(t_EOD2_res)], t_EOD2_res[,-ncol(t_EOD2_res)])
colnames(t_EOD2_res) <- c(cols[length(cols)], cols[-length(cols)])

t_EOD3_res <- ttest_res(t_EOD3)
t_EOD3_res$Label <- "Сравняване средното на ЕОД (30д - 3м)"
cols <- colnames(t_EOD3_res)
t_EOD3_res <- cbind(t_EOD3_res[, ncol(t_EOD3_res)], t_EOD3_res[,-ncol(t_EOD3_res)])
colnames(t_EOD3_res) <- c(cols[length(cols)], cols[-length(cols)])

t_EOD4_res <- ttest_res(t_EOD4)
t_EOD4_res$Label <- "Сравняване средното на ЕОД (3м - 6м)"
cols <- colnames(t_EOD4_res)
t_EOD4_res <- cbind(t_EOD4_res[, ncol(t_EOD4_res)], t_EOD4_res[,-ncol(t_EOD4_res)])
colnames(t_EOD4_res) <- c(cols[length(cols)], cols[-length(cols)])


# PO comparisons by period
t_PO1_res <- ttest_res(t_PO1)
t_PO1_res$Label <- "Сравняване средното на ПО (1д - 15д)"
cols <- colnames(t_PO1_res)
t_PO1_res <- cbind(t_PO1_res[, ncol(t_PO1_res)], t_PO1_res[,-ncol(t_PO1_res)])
colnames(t_PO1_res) <- c(cols[length(cols)], cols[-length(cols)])

t_PO2_res <- ttest_res(t_PO2)
t_PO2_res$Label <- "Сравняване средното на ПО (15д - 30д)"
cols <- colnames(t_PO2_res)
t_PO2_res <- cbind(t_PO2_res[, ncol(t_PO2_res)], t_PO2_res[,-ncol(t_PO2_res)])
colnames(t_PO2_res) <- c(cols[length(cols)], cols[-length(cols)])

t_PO3_res <- ttest_res(t_PO3)
t_PO3_res$Label <- "Сравняване средното на ПО (30д - 3м)"
cols <- colnames(t_PO3_res)
t_PO3_res <- cbind(t_PO3_res[, ncol(t_PO3_res)], t_PO3_res[,-ncol(t_PO3_res)])
colnames(t_PO3_res) <- c(cols[length(cols)], cols[-length(cols)])

t_PO4_res <- ttest_res(t_PO4)
t_PO4_res$Label <- "Сравняване средното на ПО (3м - 6м)"
cols <- colnames(t_PO4_res)
t_PO4_res <- cbind(t_PO4_res[, ncol(t_PO4_res)], t_PO4_res[,-ncol(t_PO4_res)])
colnames(t_PO4_res) <- c(cols[length(cols)], cols[-length(cols)])


startrow <- 2

writeWorksheetToFile(file = "./Output/SummaryStats v1.xlsx", TukeyEODbyPeriods, sheet = "4.2", startRow = startrow)
startrow = 4 + nrow(TukeyEODbyPeriods) + startrow
writeWorksheetToFile(file = "./Output/SummaryStats v1.xlsx", TukeyPObyPeriods, sheet = "4.2", startRow = startrow)
startrow = 4 + nrow(TukeyPObyPeriods) + startrow

writeWorksheetToFile(file = "./Output/SummaryStats v1.xlsx", t_EOD1_res, sheet = "4.2", startRow = startrow)
startrow = 4 + nrow(t_EOD1_res) + startrow
writeWorksheetToFile(file = "./Output/SummaryStats v1.xlsx", t_EOD2_res, sheet = "4.2", startRow = startrow)
startrow = 4 + nrow(t_EOD2_res) + startrow
writeWorksheetToFile(file = "./Output/SummaryStats v1.xlsx", t_EOD3_res, sheet = "4.2", startRow = startrow)
startrow = 4 + nrow(t_EOD3_res) + startrow
writeWorksheetToFile(file = "./Output/SummaryStats v1.xlsx", t_EOD4_res, sheet = "4.2", startRow = startrow)
startrow = 4 + nrow(t_EOD4_res) + startrow

writeWorksheetToFile(file = "./Output/SummaryStats v1.xlsx", t_PO1_res, sheet = "4.2", startRow = startrow)
startrow = 4 + nrow(t_PO1_res) + startrow
writeWorksheetToFile(file = "./Output/SummaryStats v1.xlsx", t_PO2_res, sheet = "4.2", startRow = startrow)
startrow = 4 + nrow(t_PO2_res) + startrow
writeWorksheetToFile(file = "./Output/SummaryStats v1.xlsx", t_PO3_res, sheet = "4.2", startRow = startrow)
startrow = 4 + nrow(t_PO3_res) + startrow
writeWorksheetToFile(file = "./Output/SummaryStats v1.xlsx", t_PO4_res, sheet = "4.2", startRow = startrow)
startrow = 4 + nrow(t_PO4_res) + startrow



# 4.3. Корелация ЕОД и ПО по периоди.
#-------------------------------------------------------------------------------
# Correlations - by groups
cor_by_periods <- data.frame(cor(plotData, use = "pairwise.complete.obs", method = "pearson"))
cor_by_periods$Labels <- row.names(cor_by_periods)
cor_by_periods <- cor_by_periods[, c(11,1:10)]


startrow <- 1
writeWorksheetToFile(file = "./Output/SummaryStats v1.xlsx", "Корелации по периоди (ЕОД)", sheet = "4.3", startRow = startrow)
startrow <- startrow + 2
writeWorksheetToFile(file = "./Output/SummaryStats v1.xlsx", cor_by_periods, sheet = "4.3", startRow = startrow, rownames = TRUE)



# 4.4. Т тест. От "4 референтни" да сравня има ли разлика между кореново развитие 0 и 1 на ЕОД и ПО. 
# И оттам взимаме референтните стойности, които сравняваме с 6тия месец на 4 травми, което е в точка 4.4. Евентуално разбити по диагнози.
#-------------------------------------------------------------------------------

t1 <- t.test(ЕОД ~ корен.развитие, data = Data2, paired = FALSE, alt = "two.sided", conf.level = 0.95)
t2 <- t.test(ПО ~ корен.развитие, data = Data2, paired = FALSE, alt = "two.sided", conf.level = 0.95)

t1res <- ttest_res(t1)
t1res$Label <- "Сравняване средното на ЕОД при незавършено и завършено кор.развитие"
cols <- colnames(t1res)
t1res <- cbind(t1res[, ncol(t1res)], t1res[,-ncol(t1res)])
colnames(t1res) <- c(cols[length(cols)], cols[-length(cols)])
t1res <- MyHelperFunctions::myRename(t1res, c("mean.x", "mean.y"), c("MeanEOD_nezav.kor.razv", "MeanEOD_zav.kor.razv"))

t2res <- ttest_res(t2)
t2res$Label <- "Сравняване средното на ПО при незавършено и завършено кор.развитие"
cols <- colnames(t2res)
t2res <- cbind(t2res[, ncol(t2res)], t2res[,-ncol(t2res)])
colnames(t2res) <- c(cols[length(cols)], cols[-length(cols)])
t2res <- MyHelperFunctions::myRename(t2res, c("mean.x", "mean.y"), c("MeanPO_nezav.kor.razv", "MeanPO_zav.kor.razv"))

startrow <- 2
writeWorksheetToFile(file = "./Output/SummaryStats v1.xlsx", t1res, sheet = "4.4", startRow = startrow)
startrow = 4 + nrow(t1res) + startrow
writeWorksheetToFile(file = "./Output/SummaryStats v1.xlsx", t2res, sheet = "4.4", startRow = startrow)



# 4.5. Т тест. Определяне на референтни стойностни за деца ПО и ЕОД (таблица 4. референтни) и сравнение на 6тия месец с тях. 
# По групи зъби (??? какви са тези групи) и диагнози (може и да не може да стане).
#-------------------------------------------------------------------------------

t1_all <- t.test(subset(Data1, !(flag == 1), ПО.6м), Data2$ПО, paired = FALSE, alt = "two.sided", conf.level = 0.95)
t1_kids <- t.test(subset(Data1, !(flag == 1) & Възраст %in% c(7:12), ПО.6м), 
                  subset(Data2, Възраст %in% c(7:12), ПО), 
                  paired = FALSE, alt = "two.sided", conf.level = 0.95)
# t1_adults <- t.test(subset(Data1, !(flag == 1) & Възраст > 12, ПО.6м), 
#                     subset(Data2, Възраст > 12, ПО), 
#                     paired = FALSE, alt = "two.sided", conf.level = 0.95)       # data1=19, data2=0


t2_all <- t.test(subset(Data1, !(flag == 1), ЕОД.6м), Data2$ЕОД, paired = FALSE, alt = "two.sided", conf.level = 0.95)
t2_kids <- t.test(subset(Data1, !(flag == 1) & Възраст %in% c(7:12), ЕОД.6м), 
                  subset(Data2, Възраст %in% c(7:12), ЕОД), 
                  paired = FALSE, alt = "two.sided", conf.level = 0.95)
# t2_adults <- t.test(subset(Data1, !(flag == 1) & Възраст > 12, ЕОД.6м), 
#                     subset(Data2, Възраст > 12, ЕОД), 
#                     paired = FALSE, alt = "two.sided", conf.level = 0.95)        # data1=19, data2=0


# output ttest results
t1_all_res <- ttest_res(t1_all)
t1_all_res$Label <- "Сравняване средното на ПО 6м (травми) и ПО (референтни) - ALL"
cols <- colnames(t1_all_res)
t1_all_res <- cbind(t1_all_res[, ncol(t1_all_res)], t1_all_res[,-ncol(t1_all_res)])
colnames(t1_all_res) <- c(cols[length(cols)], cols[-length(cols)])
t1_all_res <- MyHelperFunctions::myRename(t1_all_res, c("mean.x", "mean.y"), c("Mean ПО 6м травми", "Mean ПО референтни"))

t1_kids_res <- ttest_res(t1_kids)
t1_kids_res$Label <- "Сравняване средното на ПО 6м (травми) и ПО (референтни) - KIDS(7-12)"
cols <- colnames(t1_kids_res)
t1_kids_res <- cbind(t1_kids_res[, ncol(t1_kids_res)], t1_kids_res[,-ncol(t1_kids_res)])
colnames(t1_kids_res) <- c(cols[length(cols)], cols[-length(cols)])
t1_kids_res <- MyHelperFunctions::myRename(t1_kids_res, c("mean.x", "mean.y"), c("Mean ПО 6м травми", "Mean ПО референтни"))

# t1_adults_res <- ttest_res(t1_adults)
# t1_adults_res$Label <- "Сравняване средното на ПО 6м (травми) и ПО (референтни) - ADULTS(13+)"
# cols <- colnames(t1_adults_res)
# t1_adults_res <- cbind(t1_adults_res[, ncol(t1_adults_res)], t1_adults_res[,-ncol(t1_adults_res)])
# colnames(t1_adults_res) <- c(cols[length(cols)], cols[-length(cols)])
# t1_adults_res <- MyHelperFunctions::myRename(t1_adults_res, c("mean.x", "mean.y"), c("Mean ПО 6м травми", "Mean ПО референтни"))



t2_all_res <- ttest_res(t2_all)
t2_all_res$Label <- "Сравняване средното на ЕОД 6м (травми) и ЕОД (референтни) - ALL"
cols <- colnames(t2_all_res)
t2_all_res <- cbind(t2_all_res[, ncol(t2_all_res)], t2_all_res[,-ncol(t2_all_res)])
colnames(t2_all_res) <- c(cols[length(cols)], cols[-length(cols)])
t2_all_res <- MyHelperFunctions::myRename(t2_all_res, c("mean.x", "mean.y"), c("Mean ЕОД 6м травми", "Mean ЕОД референтни"))

t2_kids_res <- ttest_res(t2_kids)
t2_kids_res$Label <- "Сравняване средното на ЕОД 6м (травми) и ЕОД (референтни) - KIDS(7-12)"
cols <- colnames(t2_kids_res)
t2_kids_res <- cbind(t2_kids_res[, ncol(t2_kids_res)], t2_kids_res[,-ncol(t2_kids_res)])
colnames(t2_kids_res) <- c(cols[length(cols)], cols[-length(cols)])
t2_kids_res <- MyHelperFunctions::myRename(t2_kids_res, c("mean.x", "mean.y"), c("Mean ЕОД 6м травми", "Mean ЕОД референтни"))

# t2_adults_res <- ttest_res(t2_adults)
# t2_adults_res$Label <- "Сравняване средното на ЕОД 6м (травми) и ЕОД (референтни) - ADULTS(13+)"
# cols <- colnames(t2_adults_res)
# t2_adults_res <- cbind(t2_adults_res[, ncol(t2_adults_res)], t2_adults_res[,-ncol(t2_adults_res)])
# colnames(t2_adults_res) <- c(cols[length(cols)], cols[-length(cols)])
# t2_adults_res <- MyHelperFunctions::myRename(t2_adults_res, c("mean.x", "mean.y"), c("Mean ЕОД 6м травми", "Mean ЕОД референтни"))


startrow <- 2
writeWorksheetToFile(file = "./Output/SummaryStats v1.xlsx", t1_all_res, sheet = "4.5", startRow = startrow)
startrow = 4 + nrow(t1_all_res) + startrow
writeWorksheetToFile(file = "./Output/SummaryStats v1.xlsx", t1_kids_res, sheet = "4.5", startRow = startrow)
startrow = 4 + nrow(t1_kids_res) + startrow
# writeWorksheetToFile(file = "./Output/SummaryStats v1.xlsx", t1_adults_res, sheet = "4.5", startRow = startrow)
# startrow = 4 + nrow(t1_adults_res) + startrow

writeWorksheetToFile(file = "./Output/SummaryStats v1.xlsx", t2_all_res, sheet = "4.5", startRow = startrow)
startrow = 4 + nrow(t2_all_res) + startrow
writeWorksheetToFile(file = "./Output/SummaryStats v1.xlsx", t2_kids_res, sheet = "4.5", startRow = startrow)
# startrow = 4 + nrow(t2_kids_res) + startrow
# writeWorksheetToFile(file = "./Output/SummaryStats v1.xlsx", t2_adults_res, sheet = "4.5", startRow = startrow)



file.rename("./Output/SummaryStats v1.xlsx", "./Output/SummaryStats v4.xlsx")

