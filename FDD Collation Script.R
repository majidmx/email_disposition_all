rm(list=ls())
rm(list=ls(all=TRUE)) 
Sys.setenv(JAVA_HOME='C:/Program Files/Java/jre1.8.0_251')
requiredPackages = c('rio','readxl','rJava','xlsx','mailR','RDCOMClient','dplyr','plyr', 'stringi', 'readxl', 'stringr', 'openxlsx', 'xtable')
for(p in requiredPackages){
  if(!require(p,character.only = TRUE)) install.packages(p)
  library(p,character.only = TRUE)
}
library("readxl")
library('rio')
library("rJava")
library("xlsx")
library("plyr")
library("dplyr")
library("stringr")
library("stringi")
library("readxl")
library("mailR")
library("RDCOMClient")
library("openxlsx")
library("xtable")

invalid = read_excel("C:/Users/majidmm/Documents/FDD/new_script_templates/Invalid_receipts_data.xlsx", sheet = 'Sheet1')
invalid["code"] <- 3


valid = read_excel("C:/Users/majidmm/Documents/FDD/new_script_templates/Valid_receipts_data.xlsx", sheet = 'Sheet1')
valid["code"] <- 2


registration = read_excel("C:/Users/majidmm/Documents/FDD/new_script_templates/Registration_data.xlsx", sheet = 'Sheet1')
registration ["code"] <- 1


new <- rbind.fill(registration,valid,invalid)

print(new)

write.xlsx(new, 'C:/Users/majidmm/Documents/FDD/new_script_templates/concat_data.xlsx', sheetName = "Sheet1")