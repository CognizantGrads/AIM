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










