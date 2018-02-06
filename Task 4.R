
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

plotData <- Data1[, c("ЕОД.1д", "ЕОД.15д", "ЕОД.30д", "ЕОД.3м", "ЕОД.6м", "ПО.1д", "ПО.15д", "ПО.30д", "ПО.3м", "ПО.6м")]
EODdata <- melt(data = plotData[,c(1:5)], measure.vars = c(1:5), variable.name = "Period", value.name = "EOD")
EODdata[,"Period"] <- gsub("[[:punct:]]", "", EODdata[,"Period"])
EODdata[,"Period"] <- gsub("[ЕОД]", "", EODdata[,"Period"])
EODdata[,"Period"] <- replaceValues(EODdata$Period, c("1д", "15д", "30д", "3м", "6м"), c(1,2,3,4,5))
EODdata$Period <- as.numeric(EODdata$Period)

POdata <- melt(data = plotData[,c(6:10)], measure.vars = c(1:5), variable.name = "Period", value.name = "PO")
POdata[,"Period"] <- gsub("[[:punct:]]", "", POdata[,"Period"])
POdata[,"Period"] <- gsub("[ПО]", "", POdata[,"Period"])
POdata[,"Period"] <- replaceValues(POdata$Period, c("1д", "15д", "30д", "3м", "6м"), c(1,2,3,4,5))
POdata$Period <- as.numeric(POdata$Period)


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

# Data42 <- Data1[!(Data1$flag == 1), ]






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


