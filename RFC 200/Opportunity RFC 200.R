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
path<-"C:/Users/aleksander.politansk/Desktop/RFC/RFC 200/Opportunity/"

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

ids<-read.table(paste0(path, "Opportunity ids.txt"), stringsAsFactors = FALSE)
ids<-ids$V1



###################################################################################
###################################################################################
###################################################################################
###################################################################################



Event_result<-data.frame()
  j<-1
  j_max<-ceiling(length(ids)/i_rows)
  
for (i in seq(from=1,to=length(ids),by=i_rows)) {

  Event_soql<-paste0("
  SELECT WhatId, ActivityDateTime, EndDateTime, rebuild_LastModifiedDate__c
  FROM Event
  WHERE WhatId in (
  "
  ,paste0("'",paste0(na.omit(ids[i:(i+(i_rows-1))]), sep = "", collapse = "','"),"'"),
  ")
  "
  )
  
Event_soql<-gsub(pattern = "\n", replacement = "", x = Event_soql)
Object_Content <- sf_query(Event_soql)
Event_result <- rbind(Event_result,Object_Content)
Event_result<-Event_result %>% distinct()
# print(paste0('Iteration number ',j," out of ",j_max," is done. Time: ",format(Sys.time(), "%X") ))
j<-j+1
}  

  
# Take max date from extracted columns
Event_resultx <- apply(Event_result[,-1],1,max,na.rm=TRUE)

# Prepare Event_result_out data frame
Event_result_out <- data.frame(Event_result$WhatId,Event_resultx)
names(Event_result_out) <- c("WhatId","Max_date")
Event_result_out$Max_date<-stri_datetime_parse(Event_result_out$Max_date,format = "uuuu-MM-dd HH:mm:ss",
  lenient = FALSE, tz = NULL, locale = NULL)
  

Event_result_out <- Event_result_out %>% 
  select(WhatId, Max_date) %>%
  group_by(WhatId) %>%
  summarise(Max_date_distinct=max(Max_date))


###################################################################################



Linked_Document_result<-data.frame()
  j<-1
  j_max<-ceiling(length(ids)/i_rows)
  
for (i in seq(from=1,to=length(ids),by=i_rows)) {

  Linked_Document_soql<-paste0("
  SELECT Opportunity__c, rebuild_LastModifiedDate__c
  FROM Add_Organisation_Contacts_Document_Link__c
  WHERE Opportunity__c in (
  "
  ,paste0("'",paste0(na.omit(ids[i:(i+(i_rows-1))]), sep = "", collapse = "','"),"'"),
  ")
  "
  )
  
Linked_Document_soql<-gsub(pattern = "\n", replacement = "", x = Linked_Document_soql)
Object_Content <- sf_query(Linked_Document_soql)
Linked_Document_result <- rbind(Linked_Document_result,Object_Content)
Linked_Document_result<-Linked_Document_result %>% distinct()
# print(paste0('Iteration number ',j," out of ",j_max," is done. Time: ",format(Sys.time(), "%X") ))
j<-j+1
}  

  
Linked_Document_result_out <- Linked_Document_result %>% 
  select(Opportunity__c, rebuild_LastModifiedDate__c) %>%
  group_by(Opportunity__c) %>%
  summarise(Max_date=max(rebuild_LastModifiedDate__c))


###################################################################################



Lead_result<-data.frame()
  j<-1
  j_max<-ceiling(length(ids)/i_rows)
  
for (i in seq(from=1,to=length(ids),by=i_rows)) {

  Lead_soql<-paste0("
  SELECT Converted_Opportunity__c, rebuild_LastModifiedDate__c
  FROM Lead
  WHERE Converted_Opportunity__c in (
  "
  ,paste0("'",paste0(na.omit(ids[i:(i+(i_rows-1))]), sep = "", collapse = "','"),"'"),
  ")
  "
  )
  
Lead_soql<-gsub(pattern = "\n", replacement = "", x = Lead_soql)
Object_Content <- sf_query(Lead_soql)
Lead_result <- rbind(Lead_result,Object_Content)
Lead_result<-Lead_result %>% distinct()
# print(paste0('Iteration number ',j," out of ",j_max," is done. Time: ",format(Sys.time(), "%X") ))
j<-j+1
}  

  
Lead_result_out <- Lead_result %>% 
  select(Converted_Opportunity__c, rebuild_LastModifiedDate__c) %>%
  group_by(Converted_Opportunity__c) %>%
  summarise(Max_date=max(rebuild_LastModifiedDate__c))


###################################################################################


Opportunity_result<-data.frame()
  j<-1
  j_max<-ceiling(length(ids)/i_rows)
  
for (i in seq(from=1,to=length(ids),by=i_rows)) {

  Opportunity_soql<-paste0("
  SELECT Parent_Opportunity__c, rebuild_LastModifiedDate__c
  FROM Opportunity
  WHERE Parent_Opportunity__c in (
  "
  ,paste0("'",paste0(na.omit(ids[i:(i+(i_rows-1))]), sep = "", collapse = "','"),"'"),
  ")
  "
  )
  
Opportunity_soql<-gsub(pattern = "\n", replacement = "", x = Opportunity_soql)
Object_Content <- sf_query(Opportunity_soql)
Opportunity_result <- rbind(Opportunity_result,Object_Content)
Opportunity_result<-Opportunity_result %>% distinct()
# print(paste0('Iteration number ',j," out of ",j_max," is done. Time: ",format(Sys.time(), "%X") ))
j<-j+1
}  

  
Opportunity_result_out <- Opportunity_result %>% 
  select(Parent_Opportunity__c, rebuild_LastModifiedDate__c) %>%
  group_by(Parent_Opportunity__c) %>%
  summarise(Max_date=max(rebuild_LastModifiedDate__c))



###################################################################################


Opportunity_Competitor__c_result<-data.frame()
  j<-1
  j_max<-ceiling(length(ids)/i_rows)
  
for (i in seq(from=1,to=length(ids),by=i_rows)) {

  Opportunity_Competitor__c_soql<-paste0("
  SELECT Opportunity__c , rebuild_LastModifiedDate__c
  FROM Opportunity_Competitor__c
  WHERE Opportunity__c in (
  "
  ,paste0("'",paste0(na.omit(ids[i:(i+(i_rows-1))]), sep = "", collapse = "','"),"'"),
  ")
  "
  )
  
Opportunity_Competitor__c_soql<-gsub(pattern = "\n", replacement = "", x = Opportunity_Competitor__c_soql)
Object_Content <- sf_query(Opportunity_Competitor__c_soql)
Opportunity_Competitor__c_result <- rbind(Opportunity_Competitor__c_result,Object_Content)
Opportunity_Competitor__c_result<-Opportunity_Competitor__c_result %>% distinct()
# print(paste0('Iteration number ',j," out of ",j_max," is done. Time: ",format(Sys.time(), "%X") ))
j<-j+1
}  

  
Opportunity_Competitor__c_result_out <- Opportunity_Competitor__c_result %>% 
  select(Opportunity__c, rebuild_LastModifiedDate__c) %>%
  group_by(Opportunity__c) %>%
  summarise(Max_date=max(rebuild_LastModifiedDate__c))


###################################################################################




Opportunity_Partner__c_result<-data.frame()
  j<-1
  j_max<-ceiling(length(ids)/i_rows)
  
for (i in seq(from=1,to=length(ids),by=i_rows)) {

  Opportunity_Partner__c_soql<-paste0("
  SELECT Opportunity__c, rebuild_LastModifiedDate__c
  FROM Opportunity_Partner__c
  WHERE Opportunity__c in (
  "
  ,paste0("'",paste0(na.omit(ids[i:(i+(i_rows-1))]), sep = "", collapse = "','"),"'"),
  ")
  "
  )
  
Opportunity_Partner__c_soql<-gsub(pattern = "\n", replacement = "", x = Opportunity_Partner__c_soql)
Object_Content <- sf_query(Opportunity_Partner__c_soql)
Opportunity_Partner__c_result <- rbind(Opportunity_Partner__c_result,Object_Content)
Opportunity_Partner__c_result<-Opportunity_Partner__c_result %>% distinct()
# print(paste0('Iteration number ',j," out of ",j_max," is done. Time: ",format(Sys.time(), "%X") ))
j<-j+1
}  

  
Opportunity_Partner__c_result_out <- Opportunity_Partner__c_result %>% 
  select(Opportunity__c, rebuild_LastModifiedDate__c) %>%
  group_by(Opportunity__c) %>%
  summarise(Max_date=max(rebuild_LastModifiedDate__c))


###################################################################################


Placement_Test_Attendance__c_result<-data.frame()
  j<-1
  j_max<-ceiling(length(ids)/i_rows)
  
for (i in seq(from=1,to=length(ids),by=i_rows)) {

  Placement_Test_Attendance__c_soql<-paste0("
  SELECT RelatedOpportunity__c, Date_Of_Placement_Test__c, rebuild_LastModifiedDate__c
  from Placement_Test_Attendance__c
  WHERE RelatedOpportunity__c in (
  "
  ,paste0("'",paste0(na.omit(ids[i:(i+(i_rows-1))]), sep = "", collapse = "','"),"'"),
  ")
  "
  )
  
Placement_Test_Attendance__c_soql<-gsub(pattern = "\n", replacement = "", x = Placement_Test_Attendance__c_soql)
Object_Content <- sf_query(Placement_Test_Attendance__c_soql)
Placement_Test_Attendance__c_result <- rbind(Placement_Test_Attendance__c_result,Object_Content)
Placement_Test_Attendance__c_result<-Placement_Test_Attendance__c_result %>% distinct()
# print(paste0('Iteration number ',j," out of ",j_max," is done. Time: ",format(Sys.time(), "%X") ))
j<-j+1
}

  
# Take max date from extracted columns after additional work
Placement_Test_Attendance__c_result_max<-apply(Placement_Test_Attendance__c_result[,c(-1)],1,max,na.rm=TRUE)
  
  
# Prepare Opp_Won_result_out data frame
Placement_Test_Attendance__c_result_out <- data.frame(Placement_Test_Attendance__c_result$RelatedOpportunity__c,Placement_Test_Attendance__c_result_max)
names(Placement_Test_Attendance__c_result_out) <- c("RelatedOpportunity__c","Max_date")
Placement_Test_Attendance__c_result_out$Max_date<-as.Date(Placement_Test_Attendance__c_result_out$Max_date)
  
  
  
Placement_Test_Attendance__c_result_out <- Placement_Test_Attendance__c_result_out %>% 
  select(RelatedOpportunity__c, Max_date) %>%
  group_by(RelatedOpportunity__c) %>%
  summarise(Max_date_distinct=max(Max_date,na.rm = TRUE))  
  

###################################################################################


Activity__c_result<-data.frame()
  j<-1
  j_max<-ceiling(length(ids)/i_rows)
  
for (i in seq(from=1,to=length(ids),by=i_rows)) {

  Activity__c_soql<-paste0("
  SELECT What_Id__c, User_Last_Modified_Date__c, End_Date_Time__c
  from Activity__c
  WHERE What_Id__c in (
  "
  ,paste0("'",paste0(na.omit(ids[i:(i+(i_rows-1))]), sep = "", collapse = "','"),"'"),
  ")
  "
  )
  
Activity__c_soql<-gsub(pattern = "\n", replacement = "", x = Activity__c_soql)
Object_Content <- sf_query(Activity__c_soql)
Activity__c_result <- rbind(Activity__c_result,Object_Content)
Activity__c_result<-Activity__c_result %>% distinct()
# print(paste0('Iteration number ',j," out of ",j_max," is done. Time: ",format(Sys.time(), "%X") ))
j<-j+1
}

  
# Take max date from extracted columns after additional work
Activity__c_result_max<-apply(Activity__c_result[,-1],1,max,na.rm=TRUE)
  
  
# Prepare Opp_Won_result_out data frame
Activity__c_result_out <- data.frame(Activity__c_result$What_Id__c,Activity__c_result_max)
names(Activity__c_result_out) <- c("What_Id__c","Max_date")
Activity__c_result_out$Max_date<-as.Date(Activity__c_result_out$Max_date)
  
  
  
Activity__c_result_out <- Activity__c_result_out %>% 
  select(What_Id__c, Max_date) %>%
  group_by(What_Id__c) %>%
  summarise(Max_date_distinct=max(Max_date,na.rm = TRUE))  
  
  



###################################################################################
###################################################################################
###################################################################################
###################################################################################

# Let's make a merged dataframe

All_ids<-data.frame(ids)

Merged_out<-All_ids %>% 

  left_join(Event_result_out, by = c("ids"="WhatId")) %>%  
  left_join(Linked_Document_result_out, by = c("ids"="Opportunity__c")) %>%  
  left_join(Lead_result_out, by = c("ids"="Converted_Opportunity__c")) %>%  
  left_join(Opportunity_result_out, by = c("ids"="Parent_Opportunity__c")) %>%
  left_join(Opportunity_Competitor__c_result_out, by = c("ids"="Opportunity__c")) %>%
  left_join(Opportunity_Partner__c_result_out, by = c("ids"="Opportunity__c")) %>%
  left_join(Placement_Test_Attendance__c_result_out, by = c("ids"="RelatedOpportunity__c")) %>%
  left_join(Activity__c_result_out, by = c("ids"="What_Id__c"))


names(Merged_out)<-c("Ids","Event","Linked Document","Lead","Opportunity","Opportunity Competitor","Opportunity Partner","PT Attendance","Shadow Activity")


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


write.xlsx(x = Merged_out, file = paste0(path,"dates from children.xlsx"))



print(paste0('Finished at ',format(Sys.time(), "%X") ))

