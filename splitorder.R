##this function subsets all tables based on sales order number and write them to csv

###add number to order between tables and sales order number to distinguish


require(data.table)

splitorder<-function(){
    
    so.list<-unique(YNOTE_EH$YN_SO_NO)
    for (i in 1: length(so.list)){
        A<-YNOTE_EH[YNOTE_EH$YN_SO_NO==so.list[i],]
        list1<-c(unique(A$EH_GUID))
        B<-EH_HDR[EH_HDR$EH_GUID %in% list1,]
        C<-EH_CNTRL[EH_CNTRL$EH_GUID %in% list1,]
        list2<-c(unique(C$EH_GUID))
        D<-EH_INFO[EH_INFO$EH_GUID %in% list2,]
        list3<-c(unique(D$EH_GUID))
        E<-EH_EVMSG[EH_EVMSG$EH_GUID %in% list3,]
        list4<-c(unique(E$MSG_GUID))
        G<-EH_STAT[EH_STAT$EH_GUID %in% list3,]
        H<-EVM_PAR[EVM_PAR$EVT_GUID %in% list4,]
        I<-EVM_HDR[EVM_HDR$EVT_GUID %in% list4,]
        L<-YNOTE_EVM[YNOTE_EVM$EVT_GUID %in% list4,]
        crmtable<-CRM[CRM$Sales.order==so.list[i],]
        webui<-WEBUI[WEBUI$SalesDoc==so.list[i],]
        write.csv(A,paste("1.YNOTE_EH",so.list[i],".csv",sep = ''),row.names = F)
        write.csv(B,paste("2.EH_HDR",so.list[i],".csv",sep = ''),row.names = F)
        write.csv(C,paste("3.EH_CNTRL",so.list[i],".csv",sep = ''),row.names = F)
        write.csv(D,paste("4.EH_INFO",so.list[i],".csv",sep = ''),row.names = F)
        write.csv(E,paste("5.EH_EVMSG",so.list[i],".csv",sep = ''),row.names = F)
        write.csv(G,paste("6.EH_STAT",so.list[i],".csv",sep = ''),row.names = F)
        write.csv(H,paste("7.EVM_PAR",so.list[i],".csv",sep = ''),row.names = F)
        write.csv(I,paste("8.EVM_HDR",so.list[i],".csv",sep = ''),row.names = F)
        write.csv(L,paste("9.YNOTE_EVM",so.list[i],".csv",sep = ''),row.names = F)
        write.csv(crmtable,paste("11.CRM",so.list[i],".csv",sep=''),row.names = F)
        write.csv(webui,paste("10.webui",so.list[i],".csv",sep=''),row.names = F)
    }
    
} 
