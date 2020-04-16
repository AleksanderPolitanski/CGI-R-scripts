# Used packages
libs<-c("dplyr","httpuv","openxlsx","salesforcer","beepr","readr","zoo","stringi","lubridate")

# Install packages if needed
new.packages <- libs[!(libs %in% installed.packages()[,"Package"])]

if(length(new.packages)) {
  install.packages(new.packages)}

# Load packages
lapply(libs, FUN=library, character.only=TRUE )


######################################################################################################################################################################
######################################################################################################################################################################

# Path to your folder
path<-"C:/Users/aleksander.politansk/Desktop/RFC/RFC 200/STAGING ENV WORK/Case/"

# Which salesforce enviornment are you using?
# "Sandbox" or "Production"
env<-"Sandbox"

# How many rows you want to query in one iteration?
i_rows<-500


###################################################################################


# generate token from the salesforce environment, save it and authorise. cacheing is turned off

if (env == "Sandbox") {
  auth_details <- sf_auth(login_url = "https://test.salesforce.com", cache = FALSE)
  saveRDS(auth_details$token, "my-sandbox-token.rds")
  sf_auth(token = "my-sandbox-token.rds", cache = FALSE)
} else if (env == "Production") {
  auth_details <- sf_auth(login_url = "https://login.salesforce.com", cache = FALSE)
  saveRDS(auth_details$token, "my-production-token.rds")
  sf_auth(token = "my-production-token.rds", cache = FALSE)
} else {
  print("Please, specify which salesforce enviornment you are using!")
}

###################################################################################

ids<-read.table(paste0(path, "Case ids Walkin.txt"), stringsAsFactors = FALSE)
ids<-ids$V1



###################################################################################
###################################################################################
###################################################################################
###################################################################################


# Deleted part for Events because there are no events on cases


Task_result<-data.frame()
  j<-1
  j_max<-ceiling(length(ids)/i_rows)
  
for (i in seq(from=1,to=length(ids),by=i_rows)) {

  Task_soql<-paste0("
  SELECT WhatId, rebuild_LastModifiedDate__c, ActivityDate
  FROM Task
  WHERE WhatId in (
  "
  ,paste0("'",paste0(na.omit(ids[i:(i+(i_rows-1))]), sep = "", collapse = "','"),"'"),
  ")
  "
  )
  
Task_soql<-gsub(pattern = "\n", replacement = "", x = Task_soql)
Object_Content <- sf_query(Task_soql)
Task_result <- rbind(Task_result,Object_Content)
Task_result<-Task_result %>% distinct()
# print(paste0('Iteration number ',j," out of ",j_max," is done. Time: ",format(Sys.time(), "%X") ))
j<-j+1
}  

  
# Take max date from extracted columns
Task_resultx <- apply(Task_result[,-1],1,max,na.rm=TRUE)

# Prepare Event_result_out data frame
Task_result_out <- data.frame(Task_result$WhatId,Task_resultx)
names(Task_result_out) <- c("WhatId","Max_date")
# Task_result_out$Max_date<-stri_datetime_parse(Task_result_out$Max_date,format = "uuuu-MM-dd HH:mm:ss",
#   lenient = FALSE, tz = NULL, locale = NULL)

Task_result_out$Max_date<-as.character(Task_result_out$Max_date)  

Task_result_out <- Task_result_out %>% 
  select(WhatId, Max_date) %>%
  group_by(WhatId) %>%
  summarise(Max_date_distinct=max(Max_date, na.rm = TRUE))



###################################################################################

EmailMessage_result<-data.frame()
  j<-1
  j_max<-ceiling(length(ids)/i_rows)
  
for (i in seq(from=1,to=length(ids),by=i_rows)) {

  EmailMessage_soql<-paste0("
  SELECT ParentId, CreatedDate
  FROM EmailMessage
  WHERE ParentId in (
  "
  ,paste0("'",paste0(na.omit(ids[i:(i+(i_rows-1))]), sep = "", collapse = "','"),"'"),
  ")
  "
  )
  
EmailMessage_soql<-gsub(pattern = "\n", replacement = "", x = EmailMessage_soql)
Object_Content <- sf_query(EmailMessage_soql)
EmailMessage_result <- rbind(EmailMessage_result,Object_Content)
EmailMessage_result<-EmailMessage_result %>% distinct()
# print(paste0('Iteration number ',j," out of ",j_max," is done. Time: ",format(Sys.time(), "%X") ))
j<-j+1
}  

  
EmailMessage_result_out <- EmailMessage_result %>% 
  select(ParentId, CreatedDate) %>%
  group_by(ParentId) %>%
  summarise(Max_date=max(CreatedDate, na.rm = TRUE))


###################################################################################


LiveChatTranscript_result<-data.frame()
  j<-1
  j_max<-ceiling(length(ids)/i_rows)
  
for (i in seq(from=1,to=length(ids),by=i_rows)) {

  LiveChatTranscript_soql<-paste0("
  SELECT CaseId, CreatedDate
  FROM LiveChatTranscript
  WHERE CaseId in (
  "
  ,paste0("'",paste0(na.omit(ids[i:(i+(i_rows-1))]), sep = "", collapse = "','"),"'"),
  ")
  "
  )
  
LiveChatTranscript_soql<-gsub(pattern = "\n", replacement = "", x = LiveChatTranscript_soql)
Object_Content <- sf_query(LiveChatTranscript_soql)
LiveChatTranscript_result <- rbind(LiveChatTranscript_result,Object_Content)
LiveChatTranscript_result<-LiveChatTranscript_result %>% distinct()
# print(paste0('Iteration number ',j," out of ",j_max," is done. Time: ",format(Sys.time(), "%X") ))
j<-j+1
}

LiveChatTranscript_result_out <- LiveChatTranscript_result %>% 
  select(CaseId, CreatedDate) %>%
  group_by(CaseId) %>%
  summarise(Max_date=max(CreatedDate,na.rm = TRUE))  
  
  


###################################################################################
###################################################################################
###################################################################################
###################################################################################

# Let's make a merged dataframe

All_ids<-data.frame(ids)



  Merged_out<-All_ids %>% 
  left_join(Task_result_out, by = c("ids"="WhatId")) %>%
  left_join(EmailMessage_result_out, by = c("ids"="ParentId")) %>%  
  left_join(LiveChatTranscript_result_out, by = c("ids"="CaseId"))


  names(Merged_out)<-c("Id","Task","Email Message","Live Agent Transcript")




Merged_out$Max_date <- apply(Merged_out[,-1],1,max,na.rm=TRUE)


dates_edit<-Merged_out$Max_date

for (i in 1:length(dates_edit)){
  if(is.na(dates_edit[i])==TRUE){
    dates_edit[i]<-dates_edit[i]
  } else if(nchar(dates_edit[i])<11){
    dates_edit[i]<-paste0(dates_edit[i]," 00:00:00")
  }
}


Merged_out$Max_date<-dates_edit

# This is crucial! If not done, R changes time according to time zone difference
Merged_out<-force_tz(Merged_out,tzone="GMT")

# Randomize order to minimize number of errors
randoms<-as.data.frame(rand(n=length(Merged_out[[1]]),m=1))

Merged_out<-cbind(Merged_out,randoms)
Merged_out<-Merged_out %>% arrange(V1)
Merged_out<-Merged_out %>% select(-V1)

Merged_out_loader<-Merged_out %>% select(Id,Max_date)
names(Merged_out_loader)<-c("Id","Last_interaction_date__c")

# write.xlsx(x = Merged_out, file = paste0(path,"dates from children",Sys.Date()," 01",".xlsx"))
write.table(Merged_out_loader, file = paste0(path,"dates from children walkin.csv"), quote = TRUE, row.names = FALSE, col.names = TRUE, sep = ",", na = "")


print(paste0('Finished at ',format(Sys.time(), "%X") ))






