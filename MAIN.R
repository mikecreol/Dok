
try({
  lapply(paste0('package:',names(sessionInfo()$otherPkgs)),
         detach,
         character.only=TRUE,
         unload=TRUE,
         force = TRUE)
})

rm(list=ls())
options(scipen = 999)

Sys.setlocale("LC_ALL", "bulgarian")

library(openxlsx)
library(MyHelperFunctions)
library(gplots)
library(gtools)
library(stringi)
library(psych)
library(xtable)
library(XLConnect)
library(ggplot2)
library(reshape2)
library(gmodels)
library(MASS)

#-----------------------------------------------------------------------------------------------

source("HelperFunctions.R")

#-----------------------------------------------------------------------------------------------

# source("DataImport.R")
# source("DataPrep.R")


