rm(list=ls())
rm(list=ls(all=TRUE)) 
requiredPackages = c('readxl','dplyr','plyr','readxl','openxlsx','xtable','tidyverse')
for(p in requiredPackages){
  if(!require(p,character.only = TRUE)) install.packages(p)
  library(p,character.only = TRUE)
}
library("readxl")
library("plyr")
library("dplyr")
library("readxl")
library("tidyverse")
library("openxlsx")
library("xtable")

input = read_excel("C:/Users/majidmm/Documents/FDD/temp.xlsx", sheet = 'temp')
mailID = input %>% select(1,8,16)
mailID$Valid = "-"

isValidEmail <- function(x) {
  grepl("\\<[A-Z0-9._%+-]+@[A-Z0-9.-]+\\.[A-Z]{2,}\\>", as.character(x), ignore.case=TRUE)
}

for(i in 1:length(mailID$`dsp_shortcode`))
{
  m = unlist(strsplit(mailID$`email_id`[i],","))
  if(any(isValidEmail(m) == FALSE)==TRUE)
  {
    mailID$Valid[i] = "Invalid"
  } else
  {
    mailID$Valid[i] = "Ok"
  }
}

InvalidMails = mailID%>%filter(mailID$Valid == "Invalid")
write.xlsx(InvalidMails,"C:/Users/majidmm/Documents/FDD/InvalidEmails_r.xlsx")