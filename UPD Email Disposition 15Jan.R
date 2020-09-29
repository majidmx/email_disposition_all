rm(list=ls())
rm(list=ls(all=TRUE)) 
Sys.setenv(JAVA_HOME='C:/Program Files/Java/jre1.8.0_241')
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

input1 = read_excel("C:/Users/barthwa/Downloads/temp.xlsx", sheet = 'UPD Portal - North America (NA)')
input2 = read_excel("C:/Users/barthwa/Downloads/temp.xlsx", sheet = 'Response level')

for(a in 1:length(input2$`Response ID (UPD Portal)`))
{
  y = input1%>%filter(input1$ResponseId == input2$`Response ID (UPD Portal)`[a])
  if(is.na(y$Q9[1])==TRUE)
  {
    input2$`DSP Comments`[a] <- "-"
  } else {
    input2$`DSP Comments`[a] <- y$Q9[1]
  }
}

for(b in 1:length(input2$`Response ID (UPD Portal)`))
{
  y = input1%>%filter(input1$ResponseId == input2$`Response ID (UPD Portal)`[b])
  if(input2$`Reason Code`[b]== "Amazon Authorized")
  { 
    if(is.na(y$Q16[1])==TRUE)
    {
      input2$`Reviewer Comments`[b] <- "-"
    } else {
      input2$`Reviewer Comments`[b] <- y$Q16[1]
    }
  } else {
    if(is.na(y$Q17[1])==TRUE)
    {
      input2$`Reviewer Comments`[b] <- "-"
    } else {
      input2$`Reviewer Comments`[b] <- y$Q17[1]
    }
  }
}

route<- input1
route<-route[-c(1, 2), ]
route$route<-paste(route$Q8_1_1,"-",route$Q8_1_2,",",route$Q8_2_1,"-",route$Q8_2_2,",",route$Q8_3_1,"-",route$Q8_3_2,",",route$Q8_4_1,"-",route$Q8_4_2,",",route$Q8_5_1,"-"
                   ,route$Q8_5_2,",",route$Q8_6_1,"-",route$Q8_6_2,",",route$Q8_7_1,"-",route$Q8_7_2,",",route$Q8_8_1,"-",route$Q8_8_2,",",route$Q8_9_1,"-",route$Q8_9_2,","
                   ,route$Q8_10_1,"-",route$Q8_10_2,",",route$Q8_11_1,"-",route$Q8_11_2,",",route$Q8_12_1,"-",route$Q8_12_2,",",route$Q8_13_1,"-",route$Q8_13_2,","
                   ,route$Q8_14_1,"-",route$Q8_14_2,",",route$Q8_15_1,"-",route$Q8_15_2,",",route$Q8_16_1,"-",route$Q8_16_2,",",route$Q8_17_1,"-",route$Q8_17_2,","
                   ,route$Q8_18_1,"-",route$Q8_18_2,",",route$Q8_19_1,"-",route$Q8_19_2,",",route$Q8_20_1,"-",route$Q8_20_2)

i<-1
for(i in 1:length(route$route))
{
  route$route[i]<-gsub(", NA - NA", "", route$route[i])
}
j<-1
for(j in 1:length(route$route))
{
  route$route[j]<-gsub("NA - NA", "", route$route[j])
}

raw<-input2
#audit<-read_excel("C:/Users/anktkum/Downloads/temp.xlsx", sheet = 'Route level')
#audit<-audit%>%select(`Response ID`,`Route Code`,`UPD Minutes Requested / Route`,`UPD Status`,`Approved Minutes`,`Overall Route Time ACTUALS`,`Overall Route Time PLANNED`,`Overall Route Time DELTA* (min)`)
#email<-read_excel("H:/Others/UPD/DSP_Emails_UPD.xlsx")

dsp<-unique(raw$`DSP Short Code`)
raw$`Qualtrics Submission Date`<-as.Date(as.POSIXct(raw$`Qualtrics Submission Date`, origin="1970-01-01"))
raw$`Route Execution Date`<-as.Date(as.POSIXct(raw$`Route Execution Date`, origin="1970-01-01"))
#audit$`Overall Route Time ACTUALS`<-format(as.POSIXlt(audit$`Overall Route Time ACTUALS`) , format="%H:%M:%S")
#audit$`Overall Route Time PLANNED`<-format(as.POSIXlt(audit$`Overall Route Time PLANNED`) , format="%H:%M:%S")

raw$`Requested Routes`<-route$route[match(raw$`Response ID (UPD Portal)`,route$ResponseId)]

types<-c("Late Departure","On-Road Delay","Amazon Authorized")
k<-1
l<-1
for(k in 1:length(raw$`Response ID (UPD Portal)`))
{
  raw$`Qualtrics submission email`[k]<-gsub(";", ",", raw$`Qualtrics submission email`[k])
  raw$emaillength[k]<-length(unlist(strsplit(raw$`Qualtrics submission email`[k], ",")))
  for(l in 1:length(types))
  {
    if(grepl(types[l],raw$`Reason Code`[k]))
    {
      raw$`Reason Code`[k]<-types[l]
    }
  }
}

m<-1

for(m in m:length(dsp))
{
  temp<-raw%>%filter(raw$`DSP Short Code`== dsp[m])
  for(x in 1:length(temp$`DSP Short Code`))
  {
    temp$`Qualtrics submission email`[1] = paste(temp$`Qualtrics submission email`[1],",",temp$`Qualtrics submission email`[x])
    temp$`Qualtrics submission email`[1] = gsub(" ", "",temp$`Qualtrics submission email`[1])
  }
  z1 = gsub(" ", "",temp$`Qualtrics submission email`[1])
  z2 = unlist(strsplit(z1, ","))
  z3 = lapply(z2, tolower)
  z4 = unique(z3)
  z5 = do.call(paste, c(as.list(z4), sep = ","))
  temp$`Qualtrics submission email`[1] = z5
  Sys.setenv("R_ZIPCMD" = "C:/Rtools/bin/zip.exe")
  library(RDCOMClient)
  library(openxlsx)
  library(xtable)
  library(rio)
  OutApp <- COMCreate("Outlook.Application")
  outMail = OutApp$CreateItem(0)
  list = unlist(strsplit(temp$`Qualtrics submission email`[1], ","))
  #outMail[["Bcc"]] <-paste("jawlap@amazon.com","vmatrupa@amazon.com", sep = ";", collapse = NULL)
  #outMail[["SentOnBehalfOfName"]] <- paste(" amzl-dspinvoice@amazon.com", sep = ";", collapse = NULL)
  #outMail[["subject"]] = paste("UPD Report " , "Week-",temp$`Submission Week`[1], dsp[m])
  body<-"
  
Thanks,
Amazon Logistics" 
  msg = paste(sprintf(paste("Hi,","\n",                              "
Thank you for your recent UPD submissions. The attached report contains the status on all of your recent UPD submissions.
Please review and ensure all 'approved minutes' in the report have been input into the Work Summary Tool (WST).

If you would like an additional review of the requests pertaining to delays at the station (Under the roof), please reach out to the station management and email amzl-dspinvoice@amazon.com.                                      
                                              ")),body)
  to = list
  temp$`Qualtrics submission email`<-NULL
  temp$emaillength<-NULL
  temp<-arrange(temp,desc(`Submission Week`))
  wb <- createWorkbook()
  addWorksheet(wb, paste("Request Level Week-",temp$`Submission Week`[1]))
  writeDataTable(wb, paste("Request Level Week-",temp$`Submission Week`[1]), x = temp)
  saveWorkbook(wb, tf <- tempfile(fileext = ".xlsx"))
  from <- "dsp-upd@amazon.com"
  subject <- paste("UPD Report -",dsp[m],"-","Week",temp$`Submission Week`[1])
  #outMail[["Attachments"]]$Add(tf)
  bcc <- c("ayushiso@amazon.com","barthwa@amazon.com")
  tryCatch( { send.mail(from=from,
                        to=to,
                        bcc=bcc,
                        subject = subject,
                        body= msg,
                        smtp = list(host.name = "smtp.amazon.com"), 
                        authenticate = FALSE,
                        send = TRUE,
                        attach.files = c(tf),
                        file.names = c(paste("UPD Report -",dsp[m],"-","Week",temp$`Submission Week`[1],".xlsx")))}, error = function(e) {print(dsp[m])})
  #outMail$Send()
}