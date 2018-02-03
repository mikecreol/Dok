

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

listToExcel <- function(List, filename, sheetname, ...){
  # Outputs a list of dataframese to Excel
  # wb <- createWorkbook()
  # addWorksheet(wb, "Sheet 1")
  
  if(file.exists(filename)){
    wb <- loadWorkbook(filename)
  } else {
    wb <- loadWorkbook(filename, create = TRUE)
  }
  
  startRow = 1
  for(i in 1:length(List)){
    if(class(List) == "data.frame"){
      Df <- List
    } else {
      Df <- as.data.frame(List[[i]])
    }
    
    # writeData(wb,
    #           Df,
    #           sheet = "Sheet 1",
    #           startRow = startRow,
    #           ...)
    
    writeWorksheetToFile(filename, Df, sheet = sheetname, startRow = startRow)
    startRow = 4 + nrow(Df) + startRow
  }
  
  # saveWorkbook(wb, filename, overwrite = TRUE)
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


tukeyHSD.ByGroup <- function(var, groups, data1, ...){
  
  f1 <- formula(paste(var, "~", groups))
  a1 <- aov(f1, data = data1)
  a <- TukeyHSD(a1, ...)
  
  p_adj <- a[[1]][ , "p adj"]
  signif <- ifelse(p_adj <= 0.05, "Significant", "Non Significant")
  res <- data.frame(P.Adj = p_adj, Significance = signif)
  res$Comparisons <- row.names(res)
  res <- res[, c("Comparisons", "P.Adj", "Significance")]
  
  return(res)
}


CI <- function(x) {
  
  n <- length(na.omit(x))
  mean <- mean(x, na.rm = T)
  sd <- sd(x, na.rm = T)
  se <- sd / sqrt(n)
  me <- qnorm(.975) * se
  low95ci <- mean - me
  high95ci <- mean + me
  
  Df <- as.data.frame(cbind(low95ci, high95ci))
  colnames(Df) <- c("low 95 CI", "high 95 CI")
  
  return(Df)
}


ttest_res <- function(tt) {
  
  h0 <- c("difference in means is equal to 0")
  alt <- tt$alternative
  t.value <- round(tt$statistic, 4)
  df <- round(tt$parameter, 4)
  p.value <- round(tt$p.value, 4)
  ci_low95 <- round(tt$conf.int[1], 4)
  ci_high95 <- round(tt$conf.int[2], 4)
  if(length(tt$estimate) == 1){
    mean_of_the_differences <- tt$estimate
  } else {
    mean.x <- round(tt$estimate[[1]], 4)
    mean.y <- round(tt$estimate[[2]], 4)
  }
  
  if(length(tt$estimate) == 1){
    res <- as.data.frame(cbind(h0, alt, t.value, df, p.value, ci_low95, ci_high95, mean_of_the_differences))
  } else {
    res <- as.data.frame(cbind(h0, alt, t.value, df, p.value, ci_low95, ci_high95, mean.x, mean.y))
    }

  return(res)
  
}

CIquantile <- function(vec, perc){
  # the probability for the confidence interval
  q <- quantile(vec, seq(0, 1, perc / (2 * 100))) # by a hundred because it takes percentage value
  returnVec <- c(q[2], q[length(q) - 1])
  names(returnVec) <- c(paste0("CI_", perc, "%_Low"), paste0("CI_", perc, "%_High"))
  returnDf <- as.data.frame(t(returnVec))
  return(returnDf)
}


