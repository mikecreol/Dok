

summaryDoc <- function(x, varName = deparse(substitute(x))){
  # Usage
  # a <- mapply(summaryDoc, Data, colnames(Data))
  # listToExcel(a, filename = "SummaryStats v2.xlsx", rowNames = FALSE)
  
  if(class(x) == "numeric" | class(x) == "integer"){
    n <- length(na.omit(x))
    mean <- mean(x, na.rm = T)
    sd <- sd(x, na.rm = T)
    se <- sd / sqrt(n)
    me <- qnorm(.975) * se
    low95ci <- mean - me
    high95ci <- mean + me
    quantiles <- t(t(quantile(x, na.rm = T)))
    row.names(quantiles) <- paste0("quantiles ", row.names(quantiles))
    
    Df <- as.data.frame(rbind(mean, sd, se, low95ci, high95ci, quantiles))
    Df[[2]] <- rownames(Df)
    colnames(Df) <- c(varName, "Labels")
    Df <- Df[,c("Labels", varName)]
  } else {
    Df <- as.data.frame(table(x, useNA = "ifany")) # FREQs
    Df[,3] <- percent(Df[,2] / sum(Df[,2])) # Percentages
    colnames(Df) <- c(varName, "Count", "Percent")
  }
  
  return(Df)
}

listToExcel <- function(List, filename, ...){
  # Outputs a list of dataframese to Excel
  wb <- createWorkbook()
  addWorksheet(wb, "Sheet 1")
  
  startRow = 1
  for(i in 1:length(List)){
    if(class(List) == "data.frame"){
      Df <- List
    } else {
      Df <- as.data.frame(List[[i]])
    }
    
    writeData(wb,
              Df,
              sheet = "Sheet 1",
              startRow = startRow,
              ...)
    startRow = 4 + nrow(Df) + startRow
  }
  
  saveWorkbook(wb, filename, overwrite = TRUE)
}

percent <- function(x, digits = 2, format = "f", ...) {
  paste0(formatC(100 * x, format = format, digits = digits, ...), "%")
}

fillNAsWithUpperCell <- function(Df){
  for(i in 1:dim(Df)[1]){
    for(j in 1:dim(Df)[2])
      if(is.na(Df[i, j]) & i != 1){
        Df[i, j] <- Df[i-1, j]
      }
  }
  
  return(Df)
}