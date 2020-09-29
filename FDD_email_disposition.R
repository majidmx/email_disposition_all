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

input2 = read_excel("C:/Users/majidmm/Documents/FDD/temp.xlsx", sheet = 'temp')
input2$`RENTAL OUT DATE`<-as.Date(as.POSIXct(input2$`RENTAL OUT DATE`, origin="1970-01-01"))
input2$`RENTAL IN DATE`<-as.Date(as.POSIXct(input2$`RENTAL IN DATE`, origin="1970-01-01"))

raw<-input2

dsp<-unique(raw$`DSP CODE`)
raw$`SUBMISSION DATE`<-as.Date(as.POSIXct(raw$`SUBMISSION DATE`, origin="1970-01-01"))

m<-1

for(m in m:length(dsp))
{
  temp<-raw%>%filter(raw$`DSP CODE`== dsp[m])
  for(x in 1:length(temp$`DSP CODE`))
  {
    temp$`email_id`[1] = paste(temp$`email_id`[1],",",temp$`email_id`[x])
    temp$`email_id`[1] = gsub(" ", "",temp$`email_id`[1])
  }
  z1 = gsub(" ", "",temp$`email_id`[1])
  z2 = unlist(strsplit(z1, ","))
  z3 = lapply(z2, tolower)
  z4 = unique(z3)
  z5 = do.call(paste, c(as.list(z4), sep = ","))
  temp$`email_id`[1] = z5

  Sys.setenv("R_ZIPCMD" = "C:/Rtools/bin/zip.exe")
  library(RDCOMClient)
  library(openxlsx)
  library(xtable)
  library(rio)
  OutApp <- COMCreate("Outlook.Application")
  outMail = OutApp$CreateItem(0)
  list = unlist(strsplit(temp$`email_id`[1], ","))
  dsp_name <- temp$`DSP NAME`[1]

  body<-"
  
Thanks,
Amazon Logistics" 
  msg = paste(sprintf(paste("*** This is an unmonitored email account. Please do not reply ***","\n","\n","Dear",dsp_name,",","\n","\n",
"This is to inform you that some of the rental receipts you uploaded between",min(temp$`SUBMISSION DATE`) ,"to",max(temp$`SUBMISSION DATE`),"
in the Fleet Tool have been rejected. The attached sheet provides detailed information of all the rejected receipts. The columns A to H are entries submitted in the Fleet Tool by you and the receipt validation details are given in the subsequent columns I & J. Please check the ‘Rejection Reason’ column to understand why the entry has been rejected and re-upload valid receipts accordingly. 

Note: A valid rental receipt should include a header with vendor name/information, rental in and out dates, and vehicle information (i.e. license, Vehicle Identification Number (VIN). In and Out dates, license number given only in the attached receipt is taken into consideration for validation.

Kindly refer to the below reasons your receipt may have been rejected:

[1] Not a Rental Receipt/Bank, Credit Card Statement/Registration Card/Reservation Receipt
– The above mentioned categories are not considered a valid rental receipt as they do not include many of the details required to validate the vehicle. 

[2] Invalid Header/Vendor Information Missing from the Receipt
– An official receipt issued by the vendor is mandatory, and should contain a valid header with vendor information, or the name of the vendor, within the receipt. 

[3] Unreadable Document
–Details in the uploaded receipt are unreadable (i.e. dark, blurry). Please re-upload a clear, readable receipt.

[4] Unacceptable Document Type
–The receipt should be a printed document issued directly by the vendor. Hand-written receipts are not considered valid. Receipts should be uploaded as a .pdf or image (i.e. .jpg .png).

[5] License Number, In date or/and Due Date is Missing
 –If any or all of the above information is missing from the receipt, it is considered invalid. If a VIN/vehicle number is listed on the receipt, it is acceptable in place of a license number.
 
Please reach out to your Business Coach if you have any questions.
                                              ")),body)


  #to = list
  bcc = list 

  temp$`email_id`<-NULL
  temp$`DSP NAME`<-NULL
  temp$emaillength<-NULL
  temp<-arrange(temp,desc(`SUBMISSION WEEK`))
  
  wb <- createWorkbook()
  addWorksheet(wb, paste(format(as.Date(min(temp$`SUBMISSION DATE`)), "%m-%d"),"to",format(as.Date(max(temp$`SUBMISSION DATE`)), "%m-%d")))
  writeDataTable(wb, paste(format(as.Date(min(temp$`SUBMISSION DATE`)), "%m-%d"),"to",format(as.Date(max(temp$`SUBMISSION DATE`)), "%m-%d")), x = temp)
  saveWorkbook(wb, tf <- tempfile(fileext = ".xlsx"))
  p <- gsub("^c\\(|\\)$", "",unique(temp["STATION CODE"]))
  q <- gsub('"', "",p)

  from <- "dsp-invoice-comm@amazon.com"
  subject <- paste("Rejected Fleet Receipts -",dsp[m],"(",q,")",format(as.Date(min(temp$`SUBMISSION DATE`)), "%m-%d"),"to",format(as.Date(max(temp$`SUBMISSION DATE`)), "%m-%d"))
  
  #bcc <- c(x)
  to <- c("majidmm@amazon.com")
 #cc <- c("sisathee@amazon.com")
  tryCatch( { send.mail(from=from,
				to=to,
				bcc=bcc,
				subject = subject,
                        body= msg,
				inline = TRUE,
				encoding = "utf-8",
                        smtp = list(host.name = "smtp.amazon.com"), 
                        authenticate = FALSE,
                        send = TRUE,
                        attach.files = c(tf),
                        file.names = c(paste("Rejected Fleet Receipts -",dsp[m],"-",format(as.Date(min(temp$`SUBMISSION DATE`)), "%m-%d"),"to",format(as.Date(max(temp$`SUBMISSION DATE`)), "%m-%d"),".xlsx")))}, error = function(e) {print(dsp[m])})
  #outMail$Send()
}