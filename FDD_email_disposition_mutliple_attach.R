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

#input = read_excel("C:/Users/majidmm/Documents/FDD/new_script_templates/new_temp.xlsx", sheet = 'Sheet1')
#input$`RENTAL OUT DATE`<-as.Date(as.POSIXct(input$`RENTAL OUT DATE`, origin="1970-01-01"))
#input$`RENTAL IN DATE`<-as.Date(as.POSIXct(input$`RENTAL IN DATE`, origin="1970-01-01"))

invalid = read_excel("C:/Users/majidmm/Documents/FDD/new_script_templates/Invalid_receipts_data.xlsx", sheet = 'Sheet1')
invalid["code"] <- 3

valid = read_excel("C:/Users/majidmm/Documents/FDD/new_script_templates/Valid_receipts_data.xlsx", sheet = 'Sheet1')
valid["code"] <- 2

registration = read_excel("C:/Users/majidmm/Documents/FDD/new_script_templates/Registration_data.xlsx", sheet = 'Sheet1')
registration ["code"] <- 1


new <- rbind.fill(registration,valid,invalid)

#print(new)
#write.xlsx(new, 'C:/Users/majidmm/Documents/FDD/new_script_templates/concat_data.xlsx', sheetName = "Sheet1")


input2 <- read_excel("C:/Users/majidmm/Documents/FDD/new_script_templates/Guidelines.xlsx", sheet = 'Sheet1')

raw<-new

dsp<-unique(raw$`DSP CODE`)
#raw$`SUBMISSION DATE`<-as.Date(as.POSIXct(raw$`SUBMISSION DATE`, origin="1970-01-01"))

m<-1

for(m in m:length(dsp))
  {
  invalid_receipts<-raw%>%filter(raw$`DSP CODE`==dsp[m]&raw$'code'==3)
  for(x in 1:length(invalid_receipts$`DSP CODE`))
  {
    invalid_receipts$`email_id`[1] = paste(invalid_receipts$`email_id`[1],",",invalid_receipts$`email_id`[x])
    invalid_receipts$`email_id`[1] = gsub(" ", "",invalid_receipts$`email_id`[1])
  }
  registration_data<-raw%>%filter(raw$`DSP CODE`== dsp[m]&raw$'code'==1)
  valid_data<-raw%>%filter(raw$`DSP CODE`== dsp[m]&raw$'code'==2)
  invalid_receipts$`SUBMISSION DATE`<-as.Date(as.POSIXct(invalid_receipts$`SUBMISSION DATE`, origin="1970-01-01"))

  z1 = gsub(" ", "",invalid_receipts$`email_id`[1])
  z2 = unlist(strsplit(z1, ","))
  z3 = lapply(z2, tolower)
  z4 = unique(z3)
  z5 = do.call(paste, c(as.list(z4), sep = ","))
  invalid_receipts$`email_id`[1] = z5

  Sys.setenv("R_ZIPCMD" = "C:/Rtools/bin/zip.exe")
  library(RDCOMClient)
  library(openxlsx)
  library(xtable)
  library(rio)
  OutApp <- COMCreate("Outlook.Application")
  outMail = OutApp$CreateItem(0)
  list = unlist(strsplit(invalid_receipts$`email_id`[1], ","))
  dsp_name <- invalid_receipts$`DSP NAME`[1]

  registration_data$'Week'<-str_sub(registration_data$'Week', - 2, - 1)  
  registration_data<- arrange(registration_data,`Week`)
  #week_min <- min(registration_data$'Week')
  #week_max <- max(registration_data$'Week')

  #registration_data<- arrange(registration_data,`Week Number`)
  #week_min <- min(registration_data$'Week Number')
  #week_max <- max(registration_data$'Week Number')


  impact_sum <- sum(registration_data$'Impact',na.rm=TRUE)

  valid_data$`Date`<-as.Date(as.POSIXct(valid_data$`Date`, origin="1970-01-01"))
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

"•	Column D contains the total number of authorized vehicles from the route commitment tool.","<br>",
"•	Column E represents the total active vehicle count registered on the fleet tool.","<br>",
"•	Column F is the approximate amount that may be delayed due to unregistered vehicles in your monthly pre-payment for the corresponding week.","<br>","<br>", 

"<B>Please note, $</B>","<b>",impact_sum,"</b>","<B>is the approximate amount that may be delayed due to unregistered vehicles between </B>","<b>",week_range,"</b>","<B>(Refer column F for breakdown).</B>","<br>","<br>", 

"<B>[2] Sheet 2 (Valid receipts daily count):</B>","<br>",
"The daily breakdown of valid vehicle count based on the receipts accepted by Amazon for the month of ",date_range,"is set forth in Column E.","<br>","<br>", 

"<B>[3] Sheet 3 (Rejected receipt details):</B>","<br>",
"This will give you details about the receipts rejected between submission dates",min(invalid_receipts$'SUBMISSION DATE'),"to",max(invalid_receipts$'SUBMISSION DATE'),". Columns A through H are entries submitted in the Fleet Tool by you and the receipt validation details from Amazon are given in the subsequent columns I & J (Highlighted in red color). Please check the ‘Rejection Reason’ column to understand why the entry has been rejected. Please re-upload valid receipts accordingly.","<br>","<br>",

"<B>[4] Sheet 4 (Guidelines- reason for rejection):</B>","<br>",
"These guidelines can help you understand why the receipt was rejected.","<br>","<br>",

"Please refer to this link to access registration instruction manual for",html_link,".","<br>","<br>",

"Please reach out to your Fleet Manager if you have any questions.","<br>","<br>","Thanks,","<br>","
                                              ")),body)



  #to = list
  bcc = list 

  registration_data$'DSP NAME'<-NULL
  registration_data$'LICENSE NUMBER'<-NULL
  registration_data$'RENTAL OUT DATE'<-NULL
  registration_data$'RENTAL IN DATE'<-NULL
  registration_data$'SUBMISSION DATE'<-NULL
  registration_data$'VIN'<-NULL
  registration_data$'FILE TITLE'<-NULL
  registration_data$'STATUS'<-NULL
  registration_data$'REJECTION REASON'<-NULL
  registration_data$'SUBMISSION WEEK'<-NULL
  registration_data$'email_id'<-NULL
  registration_data$'code'<-NULL
  registration_data$'Date'<-NULL
  registration_data$'Valid Receipt Count'<-NULL
  registration_data$'$ Impact'<- paste0('$', registration_data$'Impact')
  registration_data$'Impact'<-NULL
  #registration_data$'Week'<-NULL



  valid_data$'DSP NAME'<-NULL
  valid_data$'LICENSE NUMBER'<-NULL
  valid_data$'RENTAL OUT DATE'<-NULL
  valid_data$'RENTAL IN DATE'<-NULL
  valid_data$'SUBMISSION DATE'<-NULL
  valid_data$'VIN'<-NULL
  valid_data$'FILE TITLE'<-NULL
  valid_data$'STATUS'<-NULL
  valid_data$'REJECTION REASON'<-NULL
  valid_data$'SUBMISSION WEEK'<-NULL
  valid_data$'email_id'<-NULL
  valid_data$'code'<-NULL
  valid_data$'Authorized Vehicle count'<-NULL
  valid_data$'Registered Vehicle Count'<-NULL
  valid_data$'$ Impact'<-NULL
  valid_data$'Impact'<-NULL
  valid_data$'code'<-NULL
  #valid_data$`Date`<-as.Date(as.POSIXct(valid_data$`Date`, origin="1970-01-01"))
  valid_data<-arrange(valid_data,`Date`)
  #valid_data$'Week'<-NULL


  invalid_receipts$`email_id`<-NULL
  invalid_receipts$`DSP NAME`<-NULL
  invalid_receipts$emaillength<-NULL
  invalid_receipts$'code'<-NULL
  invalid_receipts$'Week'<-NULL
  invalid_receipts$'Authorized Vehicle count'<-NULL
  invalid_receipts$'Registered Vehicle Count'<-NULL
  invalid_receipts$'$ Impact'<-NULL
  invalid_receipts$'Impact'<-NULL
  invalid_receipts$'Date'<-NULL
  invalid_receipts$'Valid Receipt Count'<-NULL
  invalid_receipts<-arrange(invalid_receipts,desc(`SUBMISSION WEEK`))
  invalid_receipts$`RENTAL OUT DATE`<-as.Date(as.POSIXct(invalid_receipts$`RENTAL OUT DATE`, origin="1970-01-01"))
  invalid_receipts$`RENTAL IN DATE`<-as.Date(as.POSIXct(invalid_receipts$`RENTAL IN DATE`, origin="1970-01-01"))
  invalid_receipts$'Week'<-NULL


  wb <- createWorkbook()
  
  addWorksheet(wb, paste("Registration Status"))
  writeDataTable(wb, paste("Registration Status"), x = registration_data)

  addWorksheet(wb, paste("Valid receipts daily count"))
  writeDataTable(wb, paste("Valid receipts daily count"), x = valid_data)

  addWorksheet(wb, paste("Invalid receipt details"))
  writeDataTable(wb, paste("Invalid receipt details"), x = invalid_receipts)
  highlightstyle <- createStyle(textDecoration = "bold",bgFill = "#FFC7CE")
  conditionalFormatting(wb, "Invalid receipt details", cols=9:10, rows=1,rule="!=0", style = highlightstyle)


  addWorksheet(wb, paste("Guideline Reason for rejection"))

  headingstyle <- createStyle(fgFill = "#DCE6F1", halign = "CENTER", textDecoration = "bold")
  boldstyle <- createStyle(textDecoration = "bold")

  writeData(wb, paste("Guideline Reason for rejection"), x = input2, headerStyle = headingstyle )
  conditionalFormatting(wb, "Guideline Reason for rejection", cols=1, rows=3,rule="!=0", style = boldstyle)


  saveWorkbook(wb, tf <- tempfile(fileext = ".xlsx"))

  p <- gsub("^c\\(|\\)$", "",unique(invalid_receipts["STATION CODE"]))
  q <- gsub('"', "",p)


  from <- "dsp-invoice-comm@amazon.com"
  #from <- "majidmm@amazon.com"
  subject <- paste("Fleet registration status -",dsp[m],"(",q,")",date_range)
  
  #bcc <- c(x)
  to <- c("dsp-invoice-comm@amazon.com")
  #to <- c("majidmm@amazon.com")

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
