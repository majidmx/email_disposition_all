library(readxl)
library(openxlsx)
library(dplyr)

#input = read.xlsx("C:/Users/majidmm/Documents/FDD/temp.xlsx", cols = c(1,3,5,6))
input = read_excel("C:/Users/majidmm/Documents/FDD/temp.xlsx")

input2 = input %>% select(1,3,5,6,7,11,12,13,15,16)
input2$`Out Date`<-as.Date(as.POSIXct(input2$`Out Date`, origin="1970-01-01"))
input2$`In Date`<-as.Date(as.POSIXct(input2$`In Date`, origin="1970-01-01"))

print(input2)
