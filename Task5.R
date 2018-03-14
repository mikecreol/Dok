
# 2.2.
#------------------------------------------------------------------------

Data <- fillNAsWithUpperCell(DataList[[2]])

Data[,"Зъбна.група"] <- factor(Data[,"Зъбна.група"], 
                               labels = c("фронт ГЧ", "премолари ГЧ", "молари ГЧ", "фронт ДЧ", "премолари ДЧ", "молари ДЧ"))

Data$out_of_range <- ifelse(Data$ПО < 75 | Data$ПО > 96, 1, 0)
ReducedData <- Data[Data$out_of_range == 0,]

startRow <- 3

writeWorksheetToFile(file = "./Output/Second Wave Stats v1.xlsx", 
                     tukeyHSD2.ByGroup(var = "ПО", groups = "Зъбна.група", ReducedData), 
                     sheet = "2.2", 
                     startRow = startRow)
startRow <- startRow + 20

writeWorksheetToFile(file = "./Output/Second Wave Stats v1.xlsx", 
                     aggregate(ПО ~ Зъбна.група, data = ReducedData, mean),
                     sheet = "2.2", 
                     startRow = startRow)
startRow <- startRow + 10

writeWorksheetToFile(file = "./Output/Second Wave Stats v1.xlsx", 
                     aggregate(ПО ~ Зъбна.група, data = ReducedData, sd),
                     sheet = "2.2", 
                     startRow = startRow)
startRow <- startRow + 10

writeWorksheetToFile(file = "./Output/Second Wave Stats v1.xlsx", 
                     tukeyHSD2.ByGroup(var = "ЕОД", groups = "Зъбна.група", ReducedData), 
                     sheet = "2.2", 
                     startRow = startRow)
startRow <- startRow + 20

writeWorksheetToFile(file = "./Output/Second Wave Stats v1.xlsx", 
                     aggregate(ЕОД ~ Зъбна.група, data = ReducedData, mean),
                     sheet = "2.2", 
                     startRow = startRow)
startRow <- startRow + 10

writeWorksheetToFile(file = "./Output/Second Wave Stats v1.xlsx", 
                     aggregate(ЕОД ~ Зъбна.група, data = ReducedData, sd),
                     sheet = "2.2", 
                     startRow = startRow)
startRow <- startRow + 10

# 4.4.  
#------------------------------------------------------------------------

Data <- fillNAsWithUpperCell(DataList[[6]])
table(Data$стадий.на.кор.разв.)

Data[, "стадий.на.кор.разв."] <- factor(Data[,"стадий.на.кор.разв."], 
                               labels = c("до 2/3", 
                                          "изграден корен с отворен апекс",
                                          "завършено"))

startRow <- 1

writeWorksheetToFile(file = "./Output/Second Wave Stats v1.xlsx", 
                     tukeyHSD2.ByGroup(var = "ПО", groups = "стадий.на.кор.разв.", Data), 
                     sheet = "4.4.", 
                     startRow = startRow)
startRow <- startRow + 6

writeWorksheetToFile(file = "./Output/Second Wave Stats v1.xlsx", 
                     tukeyHSD2.ByGroup(var = "ЕОД", groups = "стадий.на.кор.разв.", Data), 
                     sheet = "4.4.", 
                     startRow = startRow)
startRow <- startRow + 6

writeWorksheetToFile(file = "./Output/Second Wave Stats v1.xlsx", 
                     aggregate(ПО ~ стадий.на.кор.разв., data = Data, mean), 
                     sheet = "4.4.", 
                     startRow = startRow)
startRow <- startRow + 6

writeWorksheetToFile(file = "./Output/Second Wave Stats v1.xlsx", 
                     aggregate(ПО ~ стадий.на.кор.разв., data = Data, sd), 
                     sheet = "4.4.", 
                     startRow = startRow)
startRow <- startRow + 6

writeWorksheetToFile(file = "./Output/Second Wave Stats v1.xlsx", 
                     aggregate(ЕОД ~ стадий.на.кор.разв., data = Data, mean), 
                     sheet = "4.4.", 
                     startRow = startRow)
startRow <- startRow + 6

writeWorksheetToFile(file = "./Output/Second Wave Stats v1.xlsx", 
                     aggregate(ЕОД ~ стадий.на.кор.разв., data = Data, sd), 
                     sheet = "4.4.", 
                     startRow = startRow)
startRow <- startRow + 6


# 3 пулпити по колона диагноза да се обедини с ПО и ЕОД от 2.2. интактни - 2766; и да се сложат като контрола. Средно SD и Т тест по групи
#------------------------------------------------------------------------

Data <- DataList[[4]]

Data3 <- fillNAsWithUpperCell(DataList[[2]])
Data3$out_of_range <- ifelse(Data3$ПО < 75 | Data3$ПО > 96, 1, 0)
ReducedData <- Data3[Data3$out_of_range == 0,]
ReducedData$Диагноза <- factor("интактни")
ReducedData <- ReducedData[, c("Диагноза", "ЕОД", "ПО")]

table(Data$Диагноза)
Data[, "Диагноза"] <- factor(Data[,"Диагноза"], 
                                    labels = c("хиперемия", 
                                               "обратим пулпит",
                                               "необратим пулп."))

Data2 <- Data[, c("Диагноза", "ЕОД", "ПО")]
Data4 <- rbind(Data2, ReducedData)
Data <- Data4

startRow <- 3

writeWorksheetToFile(file = "./Output/Second Wave Stats v1.xlsx", 
                     tukeyHSD2.ByGroup(var = "ПО", groups = "Диагноза", Data), 
                     sheet = "3", 
                     startRow = startRow)
startRow <- startRow + 12

writeWorksheetToFile(file = "./Output/Second Wave Stats v1.xlsx", 
                     tukeyHSD2.ByGroup(var = "ЕОД", groups = "Диагноза", Data), 
                     sheet = "3", 
                     startRow = startRow)
startRow <- startRow + 12

writeWorksheetToFile(file = "./Output/Second Wave Stats v1.xlsx", 
                     aggregate(ПО ~ Диагноза, data = Data, mean), 
                     sheet = "3", 
                     startRow = startRow)
startRow <- startRow + 6

writeWorksheetToFile(file = "./Output/Second Wave Stats v1.xlsx", 
                     aggregate(ПО ~ Диагноза, data = Data, sd), 
                     sheet = "3", 
                     startRow = startRow)
startRow <- startRow + 6

writeWorksheetToFile(file = "./Output/Second Wave Stats v1.xlsx", 
                     aggregate(ЕОД ~ Диагноза, data = Data, mean), 
                     sheet = "3", 
                     startRow = startRow)
startRow <- startRow + 6

writeWorksheetToFile(file = "./Output/Second Wave Stats v1.xlsx", 
                     aggregate(ЕОД ~ Диагноза, data = Data, sd), 
                     sheet = "3", 
                     startRow = startRow)
startRow <- startRow + 6


# диагноза с "продължителност на болката"
#------------------------------------------------------------------------

Data <- DataList[[4]]

Data[, "Диагноза"] <- factor(Data[,"Диагноза"],
                             labels = c("хиперемия",
                                        "обратим пулпит",
                                        "необратим пулп."))

# Data$A <- Data[, "Диагноза"]
# Data$B <- Data$`продължителност.на.болката,.група`
# 
# tbl <- table(Data$`продължителност.на.болката,.група`, Data$Диагноза)
# as.data.frame.matrix(tbl)
# as.data.frame.matrix(prop.table(tbl) * 100)
# 
# chisq.test(tbl[1:4,1:2]) # 1 to 2
# chisq.test(tbl[3:6,2:3]) # 2 to 3


# mytable <- xtabs(~ A + B, data=Data)
# loglm1 <- loglm(~ A + B, mytable)
# summary(loglm1)

# prop.table(a) * 100

sink("./Output/cross table 1.txt")
CrossTable(Data$`продължителност.на.болката,.група`, Data$Диагноза,
           chisq = T,
           format = "SPSS")
           # prop.r = F, prop.c = F)
sink()

# диагноза с "кариозна маса"
#------------------------------------------------------------------------

Data <- DataList[[4]]
sink("./Output/cross table 3.txt")
CrossTable(Data$кариозна.маса, Data$Диагноза,
           chisq = T,
           format = "SPSS")
# prop.r = F, prop.c = F)
sink()

sink("./Output/cross table 3.txt")
CrossTable(Data$Диагноза, Data$кариозна.маса,
           chisq = T,
           format = "SPSS")
# prop.r = F, prop.c = F)
sink()

# 2.3. Всички сравнение на преди ВО след ВО и контролата да се залепи тука пак 2766, за ПО и ЕОД
#------------------------------------------------------------------------

Data <- DataList[[3]]
colnames(Data)

Data1 <- Data[, c("ЕОД.(μА)", "ПО.(%)")]
colnames(Data1) <- c("ЕОД", "ПО")
Data1$Class <- "преди ВЕ"
Data2 <- Data[, c("ЕОД.след.ВЕ.(μА)", "ПО.след.ВЕ.(%)")]
colnames(Data2) <- c("ЕОД", "ПО")
Data2$Class <- "след ВЕ"
Data4 <- Data[, c("ЕОД.след.обт..(μА)", "ПО.след.обт..(%)")]
colnames(Data4) <- c("ЕОД", "ПО")
Data4$Class <- "след обт"

Data3 <- fillNAsWithUpperCell(DataList[[2]])
Data3$out_of_range <- ifelse(Data3$ПО < 75 | Data3$ПО > 96, 1, 0)
ReducedData <- Data3[Data3$out_of_range == 0,]
ReducedData$Диагноза <- factor("интактни")
ReducedData <- ReducedData[, c("Диагноза", "ЕОД", "ПО")]
ReducedData <- ReducedData[,c("ЕОД", "ПО")]
ReducedData$Class <- "интактни"

Data <- rbind(Data1, Data2, Data4, ReducedData)

startRow <- 3
writeWorksheetToFile(file = "./Output/Second Wave Stats v1.xlsx", 
                     tukeyHSD2.ByGroup(var = "ПО", groups = "Class", Data), 
                     sheet = "2.3.", 
                     startRow = startRow)
startRow <- startRow + 10

writeWorksheetToFile(file = "./Output/Second Wave Stats v1.xlsx", 
                     tukeyHSD2.ByGroup(var = "ЕОД", groups = "Class", Data), 
                     sheet = "2.3.", 
                     startRow = startRow)


