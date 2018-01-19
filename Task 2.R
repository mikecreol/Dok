
# Определяне на доверителен интервал за здрави зъби на резултатите от ПО(Пулсова оксиметрия). 
# Min и Max, Квантили, Хистограма. От таблица 2.2.
#-------------------------------------------------------------------------------

Data <- fillNAsWithUpperCell(DataList[[2]])
# View(describe(Data))
A <- mapply(summaryDoc, Data, colnames(Data))
listToExcel(A, filename = "./Output/SummaryStats v1.xlsx", rowNames = FALSE)


# 2.1. ПО и ЕОД изследване, преди и след изолиране със силикон. T-Test. Ще установим, че няма разлика.
#-------------------------------------------------------------------------------

Data <- fillNAsWithUpperCell(DataList[[1]])


# 2.2. Mean, SD, SE, по групи зъби. Корелация между ЕОД и ПО по групи (в колона зъбна група). 
# Под 75 и над 96 на ПО да ги махна. И Tukey HSD пак по групи.
#-------------------------------------------------------------------------------

Data <- fillNAsWithUpperCell(DataList[[2]])