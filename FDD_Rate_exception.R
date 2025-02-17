rm(list=ls())
rm(list=ls(all=TRUE)) 
Sys.setenv(JAVA_HOME='C:/Program Files/Java/jre1.8.0_261')
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


valid = read_excel("C:/Users/majidmm/Documents/FDD/new_script_templates/Valid_receipts_data.xlsx", sheet = 'Sheet1')
valid["code"] <- 2

registration = read_excel("C:/Users/majidmm/Documents/FDD/new_script_templates/Registration_data.xlsx", sheet = 'Sheet1')
registration ["code"] <- 1


new <- rbind.fill(registration,valid)

#print(new)
#write.xlsx(new, 'C:/Users/majidmm/Documents/FDD/new_script_templates/concat_data.xlsx', sheetName = "Sheet1")


raw<-new

dsp<-unique(raw$`DSP CODE`)

m<-1

for(m in m:length(dsp))
  {
  registration_data<-raw%>%filter(raw$`DSP CODE`==dsp[m]&raw$'code'==1)
  for(x in 1:length(registration_data$`DSP CODE`))
  {
    registration_data$`email_id`[1] = paste(registration_data$`email_id`[1],",",registration_data$`email_id`[x])
    registration_data$`email_id`[1] = gsub(" ", "",registration_data$`email_id`[1])
  }
  valid_data<-raw%>%filter(raw$`DSP CODE`== dsp[m]&raw$'code'==2)

  z1 = gsub(" ", "",registration_data$`email_id`[1])
  z2 = unlist(strsplit(z1, ","))
  z3 = lapply(z2, tolower)
  z4 = unique(z3)
  z5 = do.call(paste, c(as.list(z4), sep = ","))
  registration_data$`email_id`[1] = z5

  Sys.setenv("R_ZIPCMD" = "C:/Rtools/bin/zip.exe")
  library(RDCOMClient)
  library(openxlsx)
  library(xtable)
  library(rio)
  OutApp <- COMCreate("Outlook.Application")
  outMail = OutApp$CreateItem(0)
  list = unlist(strsplit(registration_data$`email_id`[1], ","))
  dsp_name <- registration_data$`DSP NAME`[1]

  registration_data$'Week'<-str_sub(registration_data$'Week', - 2, - 1)  
  registration_data<- arrange(registration_data,`Week`)
  week_min <- min(registration_data$'Week')
  week_max <- max(registration_data$'Week')


  impact_sum <- sum(registration_data$'Impact',na.rm=TRUE)

  valid_data$`Date`<-as.Date(as.POSIXct(valid_data$`Date`, origin="1970-01-01"))
  #min_month<-format(min(valid_data$`Date`),"%B")
  #max_month<-format(max(valid_data$`Date`),"%B")

  html_link ='<a href="https://logistics.amazon.com/resources/file/b9bf2b8d-b790-4116-b7f1-c0693d2f0d73?version=8BX_r9LTsMQdAMlBVujTs9xfBnkOZz_3">Fleet Tool</a>'

  date_range <- 'September - October'
  week_range <- 'Week 36 - Week 41'


  body<-"
  
Amazon Logistics" 
  msg = paste(sprintf(paste("<font color=\"#FF0000\">*** This is an unmonitored email account. Please do not reply ***</font>","<br>","<br>",

"Dear","<b>",dsp_name,"</b>",",","<br>","<br>",

"We are reaching out to you with an update regarding your vehicle registrations and the associated invoice submitted in the Fleet Tool. Please refer to the attachment for the details. Guidelines to read the attachment as follows.","<br>","<br>",

"<B>[1] Sheet 1(Registration status):</B>","<br>",
"This will give you a summary of the registrations completed in the Fleet Tool.","<br>","<br>",

"�	Column D contains the total number of authorized vehicles from the route commitment tool.","<br>",
"�	Column E represents the total active vehicle count registered on the fleet tool.","<br>",
"�	Column F is the approximate amount that may be delayed due to unregistered vehicles in your monthly pre-payment for the corresponding week.","<br>","<br>", 

"<B>Please note, $</B>","<b>",impact_sum,"</b>","<B>is the approximate amount that may be delayed due to unregistered vehicles between</B>","<b>",week_range,"</b>","<B>(Refer column F for breakdown).</B>","<br>","<br>", 

"<B>[2] Sheet 2 (Valid receipts daily count):</B>","<br>",
"The daily breakdown of valid vehicle count based on the receipts accepted by Amazon for the month of ",date_range,"is set forth in Column E.","<br>","<br>", 

"Please refer to this link to access registration instruction manual for",html_link,".","<br>","<br>",

"Please reach out to your Fleet Manager if you have any questions.","<br>","<br>","Thanks,","<br>","
                                              ")),body)



  #to = list
  bcc = list 

  registration_data$'DSP NAME'<-NULL
  registration_data$'email_id'<-NULL
  registration_data$'code'<-NULL
  registration_data$'Date'<-NULL
  registration_data$'Valid Receipt Count'<-NULL
  registration_data$'$ Impact'<- paste0('$', registration_data$'Impact')
  registration_data$'Impact'<-NULL
  #registration_data$'Week Number'<-NULL



  valid_data$'DSP NAME'<-NULL
  valid_data$'email_id'<-NULL
  valid_data$'code'<-NULL
  valid_data$'Authorized Vehicle count'<-NULL
  valid_data$'Registered Vehicle Count'<-NULL
  valid_data$'$ Impact'<-NULL
  valid_data$'Impact'<-NULL
  valid_data$'code'<-NULL
  #valid_data$`Date`<-as.Date(as.POSIXct(valid_data$`Date`, origin="1970-01-01"))
  valid_data<-arrange(valid_data,`Date`)
  #valid_data$'Week Number'<-NULL

  wb <- createWorkbook()
  
  addWorksheet(wb, paste("Registration Status"))
  writeDataTable(wb, paste("Registration Status"), x = registration_data)

  addWorksheet(wb, paste("Valid receipts daily count"))
  writeDataTable(wb, paste("Valid receipts daily count"), x = valid_data)

  saveWorkbook(wb, tf <- tempfile(fileext = ".xlsx"))

  p <- gsub("^c\\(|\\)$", "",unique(registration_data["STATION CODE"]))
  q <- gsub('"', "",p)


  from <- "dsp-invoice-comm@amazon.com"
  #from <- "majidmm@amazon.com"

  subject <- paste("Fleet registration status -",dsp[m],"(",q,")",date_range)
  
  #bcc <- c(x)
  #to <- c("majidmm@amazon.com")
  to <- c("dsp-invoice-comm@amazon.com")
  
 #cc <- c("sisathee@amazon.com")
  tryCatch( { send.mail(from=from,
				to=to,
				bcc=bcc,
				subject = subject,
                        body= msg,
				inline = TRUE,
				html = TRUE,
				encoding = "utf-8",
                        smtp = list(host.name = "smtp.amazon.com"), 
                        authenticate = FALSE,
                        send = TRUE,
                        attach.files = c(tf),
                        file.names = c(paste("Fleet registration status -",dsp[m],"(",q,")",date_range,".xlsx")))}, 
				error = function(e) {print(dsp[m])})
  #outMail$Send()
}
