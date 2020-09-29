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


dsp_payments = read_excel("C:/Users/username-xyz/Documents/DSP_PAYMENTS/template.xlsx", sheet = 'Sheet1')

raw<-dsp_payments

station<-unique(raw$`Station`)

m<-1

for(m in m:length(station))
  {
  payments_data<-raw%>%filter(raw$`Station`==station[m])
  for(x in 1:length(payments_data$`Station`))
  {
    payments_data$`Send To (ORAMs)`[1] = paste(payments_data$`Send To (ORAMs)`[1],";",payments_data$`Send To (ORAMs)`[x])
    payments_data$`Send To (ORAMs)`[1] = gsub(" ", "",payments_data$`Send To (ORAMs)`[1])
  }

  z1 = gsub(" ", "",payments_data$`Send To (ORAMs)`[1])
  z2 = unlist(strsplit(z1, ";"))
  z2_1 = paste(z2, "@amazon.com",sep="")
  z3 = lapply(z2_1, tolower)
  z4 = unique(z3)
  z5 = do.call(paste, c(as.list(z4), sep = ","))
  payments_data$`Send To (ORAMs)`[1] = z5

  Sys.setenv("R_ZIPCMD" = "C:/Rtools/bin/zip.exe")
  library(RDCOMClient)
  library(openxlsx)
  library(xtable)
  library(rio)
  OutApp <- COMCreate("Outlook.Application")
  outMail = OutApp$CreateItem(0)
  list = unlist(strsplit(payments_data$`Send To (ORAMs)`[1], ","))

  #payments_data$`Time Stamp`<-as.Date(as.POSIXct(payments_data$`Time Stamp`, origin="1970-01-01"))
  #min_month<-format(min(valid_data$`Date`),"%B")
  #max_month<-format(max(valid_data$`Date`),"%B")
  time_stamp<-payments_data$`Time Stamp`[1]
  html_link ='<a href="mailto:amzl-dspinvoice@amazon.com">amzl-dspinvoice@amazon.com</a>'



  body<-"
  
" 

  msg = paste(sprintf(paste("Hello Team,","<br>","<br>",
"If you need permissions to the Invoice Portal – please “reply all” and notify our team","<b>IMMEDIATELY</b>",".","<br>","<br>",
"NA DSP Payments Team is sending you this communication to inform Station","[",station[m],"]","that there are open disputes pending for Ops action.","<br>",
"These disputes will need to be actioned by the","<span style=\"background: yellow;\"><b>EOD today</b></span>",". IF the dispute is eligible for an adjustment, the dispute will need to go through","<b>BOTH Ops Tier 1 AND 2</b>","for complete resolution. Please","<b>“reply all”</b>","to notify the sender that this dispute has been action or is work-in-progress.","<br>","<br>",
"Please also refer to and review the attached pptx for the action plan and for a quick walk-through on how to effectively utilize the Flex and SIM portal.","<br>","<br>",
"Data provided on","<b>",time_stamp,"</b>","<br>","<br>",
"Disputes that need immediate action in the Invoice / Flex Portal:","Open Ops Tier Invoices for","[",station[m],"]",".xlsx","<br>","<br>",
"<span style=\"background: yellow;\"><b>NOTE:</b></span>","<font color=\"#FF0000\"><i><u>Do not</u></i></font>","<i>push adjustments of</i>","<b><i>zero dollar amount</i></b>","<i>to Finance in the Flex/ Invoice / Cortex Payment portal.</i>","<b><i>The adjustments will need to applied by OTR team, IF the dispute is valid.</i></b>","<i>The adjustment can be added by pressing “adjust invoice” – selecting the related service type.</i>","<br>","<br>",
"<font color=\"#1874cd\"><b>Thanks,</b></font>","<br>",
"<font color=\"#1874cd\"><b>Tanner Reynolds |</b></font>","<font color=\"#ee7600\"><b>NA DSP PNP (Pricing, Incentives, & Payments)</b></font>","<br>",
"<font size=2 color=\"#36648b\"><b>Email:</b></font>","<font size=2 color=\"#36648b\">",html_link,"</font>","
                                              ")),body)



  to = list
  #bcc = list 

  payments_data$'last_dispute_date'<-NULL
  payments_data$'Time Stamp'<-NULL
  payments_data$'Program Type'<-NULL
  payments_data$'Assignees'<-NULL
  payments_data$'Send To (ORAMs)'<-NULL
  payments_data$'Escalation'<-NULL
  payments_data$'Escalation POC'<-NULL
  payments_data$'BC POC'<-NULL
  payments_data$'Send To (Escalation)'<-NULL
  payments_data$'Station has SIM'<-NULL
  payments_data$'Note'<-NULL
  

  wb <- createWorkbook()
  
  addWorksheet(wb, paste("Disputes Resolution"))
  writeDataTable(wb, paste("Disputes Resolution"), x = payments_data)

  saveWorkbook(wb, tf <- tempfile(fileext = ".xlsx"))

  from <- "xyz@amazon.com"
  subject <- paste("**Action Required** Dispute Resolution for ","[",station[m],"]")
  
  #bcc <- c(x)
  #to <- c("xyz@amazon.com")
  
  #cc <- c("xyz@amazon.com")
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
                        file.names = c(paste("Open Ops Tier Invoices for","[",station[m],"]",".xlsx")))}, 
				error = function(e) {print(station[m])})
  #outMail$Send()
}