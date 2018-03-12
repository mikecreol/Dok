
# 2.2.
#------------------------------------------------------------------------

Data <- fillNAsWithUpperCell(DataList[[2]])

Data[,"Зъбна.група"] <- factor(Data[,"Зъбна.група"], 
                               labels = c("фронт ГЧ", "премолари ГЧ", "молари ГЧ", "фронт ДЧ", "премолари ДЧ", "молари ДЧ"))

Data$out_of_range <- ifelse(Data$ПО < 75 | Data$ПО > 96, 1, 0)
ReducedData <- Data[Data$out_of_range == 0,]


writeWorksheetToFile(file = "./Output/Second Wave Stats v1.xlsx", 
                     tukeyHSD2.ByGroup(var = "ПО", groups = "Зъбна.група", ReducedData), 
                     sheet = "2.2", 
                     startRow = 3)

writeWorksheetToFile(file = "./Output/Second Wave Stats v1.xlsx", 
                     aggregate(ПО ~ Зъбна.група, data = ReducedData, mean),
                     sheet = "2.2", 
                     startRow = 20)

writeWorksheetToFile(file = "./Output/Second Wave Stats v1.xlsx", 
                     aggregate(ПО ~ Зъбна.група, data = ReducedData, sd),
                     sheet = "2.2", 
                     startRow = 35)



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