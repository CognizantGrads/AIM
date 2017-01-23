##this script imports Unilever SAP data in R from excel tables##
library(openxlsx)
library(dplyr)
library(plyr)
library(data.table)
library(chron)
options(scipen=999)

####LOAD DATA####
pname <- "D://D//Cognizant//Projects//Internal//Unilever//" # change if necessary
fname <- "tabs.xlsx"#change if necessary
fname2<- "WebUI.xlsx"
fname3<-"CRM.xlsx"

fpname <- paste(pname,fname,sep='')
fpname2<-paste(pname,fname2,sep='')
fpname3<-paste(pname,fname3,sep='')

## EVENT MESSAGE
EVM_PAR<-read.xlsx(fpname, sheet = 9)###event parameters
EVM_HDR<-read.xlsx(fpname, sheet = 10) ### event message header details
YNOTE_EVM <-read.xlsx(fpname, sheet = 11) ### how the information received is stored, from evm_par and evm_hdr
## EVENT MESSAGE/EVENT HANDLER
EH_EVMSG<-read.xlsx(fpname, sheet = 7) ## This table links ECC and event management
## EVENT HANDLER
EH_CNTRL <- read.xlsx(fpname, sheet = 5) ## EH control parameters
EH_INFO <-  read.xlsx(fpname, sheet = 6) ## EH info parameters
EH_HDR <- read.xlsx(fpname, sheet = 4) ## not relevant eh header info, from EH_EVMSG
YNOTE_EH<- read.xlsx(fpname,sheet = 3) ## final
EH_STAT <- read.xlsx(fpname,sheet = 8)##status of each order..looks isolated

WEBUI<-read.xlsx(fpname2)###webui
CRM<-read.xlsx(fpname3) ##crm

###Delete Columns with all NA's or ALL 0
EH_EVMSG$EARLIEST_MSG_DTE <- NULL
EH_EVMSG$LATEST_MSG_DATE <- NULL
EH_EVMSG$ADD_DATA <- NULL
EH_EVMSG$BUILT_EH_HIER <- NULL
EH_EVMSG$EVENT_DATE <- NULL #data also provides GMT
EH_EVMSG$EVENT_TZONE <-NULL
EH_EVMSG$MSG_EXP_TZONE <-NULL

###Function convert to date
numToDate <- function(x){
    x <- sub('(.{4})(.{2})(.{2})(.{2})(.{2})(.{2})', "\\1-\\2-\\3 \\4:\\5:\\6",x)
    dfx = t(as.data.frame(strsplit(x,' ')))
    row.names(dfx) = NULL
    dfx <- as.data.frame(dfx)
    dfx[["DateTime"]] <- paste(dfx$V1, dfx$V2, sep=" ")
    dfx[["DateTime"]] <- strptime(as.character(dfx$DateTime), "%Y-%m-%d %H:%M:%S")
    x <- dfx[["DateTime"]]
}
# columns to convert to dates
# PROC_DATE, MSG_RCVD_DATE, EARLIEST_EV_DATE, LATEST_EV_DATE ,MSG_DATE_UTC ,EVENT_DATE_UTC
EH_EVMSG$PROC_DATE <-numToDate(EH_EVMSG$PROC_DATE) #Change for each date column
EH_EVMSG$MSG_RCVD_DATE <-numToDate(EH_EVMSG$MSG_RCVD_DATE)
EH_EVMSG$EARLIEST_EV_DATE <-numToDate(EH_EVMSG$EARLIEST_EV_DATE)
EH_EVMSG$LATEST_EV_DATE <-numToDate(EH_EVMSG$LATEST_EV_DATE)
EH_EVMSG$MSG_DATE_UTC <-numToDate(EH_EVMSG$MSG_DATE_UTC)
EH_EVMSG$EVENT_DATE_UTC <-numToDate(EH_EVMSG$EVENT_DATE_UTC)

TimezoneUTC <- function(x){ #Minus hour(CET to GMT)
    x <- x-1
}

# columns to convert from CET to UTC/GMT
#EARLIEST_EV_DATE,LATEST_EV_DATE
EH_EVMSG$LATEST_EV_DATE$hour <- TimezoneUTC(EH_EVMSG$LATEST_EV_DATE$hour) #Change for each CET column
EH_EVMSG$EARLIEST_EV_DATE$hour <- TimezoneUTC(EH_EVMSG$EARLIEST_EV_DATE$hour) #Change for each CET column


####add material group to EH_EVMSG based on EH_GUID
EH_EVMSG_NEW<-merge(x = EH_EVMSG, y = YNOTE_EH[ , c("EH_GUID","YN_SO_MAT","YN_SO_NO")], by = "EH_GUID", all.x=TRUE)

names(WEBUI)[names(WEBUI) == 'SalesDoc'] <- 'YN_SO_NO' #Rename column 
dfTemp <- data.frame("YN_SO_NO" = WEBUI$YN_SO_NO, "Payer.Name" = WEBUI$Payer.Name, stringsAsFactors=FALSE) #Temp DF
EH_EVMSG_NEW$Customer <- dfTemp$Payer.Name[match(EH_EVMSG_NEW$YN_SO_NO, dfTemp$YN_SO_NO)] #Adds a new column based order number match
write.csv(EH_EVMSG_NEW, file = "D://D//Cognizant//Projects//Internal//Unilever//Master.csv",row.names=FALSE)
