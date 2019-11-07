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
path<-"C:/Users/aleksander.politansk/Desktop/RFC/RFC 200/"

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

ids<-read.table(paste0(path, "Account ids.txt"), stringsAsFactors = FALSE)
ids<-ids$V1



###################################################################################
###################################################################################
###################################################################################
###################################################################################

Account_result<-data.frame()
  j<-1
  j_max<-ceiling(length(ids)/i_rows)
  
for (i in seq(from=1,to=length(ids),by=i_rows)) {

  Account_soql<-paste0("
  SELECT Parent_Guardian__c, rebuild_LastModifiedDate__c
  FROM Account
  WHERE Parent_Guardian__c in (
  "
  ,paste0("'",paste0(na.omit(ids[i:(i+(i_rows-1))]), sep = "", collapse = "','"),"'"),
  ")
  "
  )
  
Account_soql<-gsub(pattern = "\n", replacement = "", x = Account_soql)
Object_Content <- sf_query(Account_soql)
Account_result <- rbind(Account_result,Object_Content)
Account_result<-Account_result %>% distinct()
# print(paste0('Iteration number ',j," out of ",j_max," is done. Time: ",format(Sys.time(), "%X") ))
j<-j+1
}

Account_result_out <- Account_result %>% 
  select(Parent_Guardian__c, rebuild_LastModifiedDate__c) %>%
  group_by(Parent_Guardian__c) %>%
  summarise(Max_date=max(rebuild_LastModifiedDate__c,na.rm = TRUE))  
  
  
  
###################################################################################
   


Org_Account_result<-data.frame()
  j<-1
  j_max<-ceiling(length(ids)/i_rows)
  
for (i in seq(from=1,to=length(ids),by=i_rows)) {

  Org_Account_soql<-paste0("
  SELECT AccountId, User_Last_Modified_Date__c
  FROM AccountContactRelation
  WHERE AccountId in (
  "
  ,paste0("'",paste0(na.omit(ids[i:(i+(i_rows-1))]), sep = "", collapse = "','"),"'"),
  ")
  and Account.RecordTypeId = '0121r000000plHlAAI'
  "
  )
  
Org_Account_soql<-gsub(pattern = "\n", replacement = "", x = Org_Account_soql)
Object_Content <- sf_query(Org_Account_soql)
Org_Account_result <- rbind(Org_Account_result,Object_Content)
Org_Account_result<-Org_Account_result %>% distinct()
# print(paste0('Iteration number ',j," out of ",j_max," is done. Time: ",format(Sys.time(), "%X") ))
j<-j+1
}

Org_Account_result_out <- Org_Account_result %>% 
  select(AccountId, User_Last_Modified_Date__c) %>%
  group_by(AccountId) %>%
  summarise(Max_date=max(User_Last_Modified_Date__c,na.rm = TRUE))  
  
  
  
###################################################################################


LiveChatTranscript_result<-data.frame()
  j<-1
  j_max<-ceiling(length(ids)/i_rows)
  
for (i in seq(from=1,to=length(ids),by=i_rows)) {

  LiveChatTranscript_soql<-paste0("
  SELECT AccountId, CreatedDate
  FROM LiveChatTranscript
  WHERE AccountId in (
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
  select(AccountId, CreatedDate) %>%
  group_by(AccountId) %>%
  summarise(Max_date=max(CreatedDate,na.rm = TRUE))  
  
  
  
###################################################################################
   
    
  
Person_Account_result<-data.frame()
  j<-1
  j_max<-ceiling(length(ids)/i_rows)
  
for (i in seq(from=1,to=length(ids),by=i_rows)) {

  Person_Account_soql<-paste0("
  SELECT Contact.AccountId, User_Last_Modified_Date__c
  FROM AccountContactRelation
  WHERE Contact.AccountId in (
  "
  ,paste0("'",paste0(na.omit(ids[i:(i+(i_rows-1))]), sep = "", collapse = "','"),"'"),
  ")
  and Contact.Account.RecordTypeId <> '0121r000000plHlAAI'
  "
  )
  
Person_Account_soql<-gsub(pattern = "\n", replacement = "", x = Person_Account_soql)
Object_Content <- sf_query(Person_Account_soql)
Person_Account_result <- rbind(Person_Account_result,Object_Content)
Person_Account_result<-Person_Account_result %>% distinct()
# print(paste0('Iteration number ',j," out of ",j_max," is done. Time: ",format(Sys.time(), "%X") ))
j<-j+1
}  

Person_Account_result_out <- Person_Account_result %>% 
  select(Contact.AccountId, User_Last_Modified_Date__c) %>%
  group_by(Contact.AccountId) %>%
  summarise(Max_date=max(User_Last_Modified_Date__c,na.rm = TRUE))  
  
###################################################################################
    
  
Event_result<-data.frame()
  j<-1
  j_max<-ceiling(length(ids)/i_rows)
  
for (i in seq(from=1,to=length(ids),by=i_rows)) {

  Event_soql<-paste0("
  SELECT AccountId, ActivityDateTime, EndDateTime, rebuild_LastModifiedDate__c
  FROM Event
  WHERE AccountId in (
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
Event_result_out <- data.frame(Event_result$AccountId,Event_resultx)
names(Event_result_out) <- c("AccountId","Max_date")
Event_result_out$Max_date<-stri_datetime_parse(Event_result_out$Max_date,format = "uuuu-MM-dd HH:mm:ss",
  lenient = FALSE, tz = NULL, locale = NULL)
  

Event_result_out <- Event_result_out %>% 
  select(AccountId, Max_date) %>%
  group_by(AccountId) %>%
  summarise(Max_date_distinct=max(Max_date))


###################################################################################
    
  
Key_User__c_result<-data.frame()
  j<-1
  j_max<-ceiling(length(ids)/i_rows)
  
for (i in seq(from=1,to=length(ids),by=i_rows)) {

  Key_User__c_soql<-paste0("
  SELECT Account__c, rebuild_LastModifiedDate__c
  FROM Key_User__c
  WHERE Account__c in (
  "
  ,paste0("'",paste0(na.omit(ids[i:(i+(i_rows-1))]), sep = "", collapse = "','"),"'"),
  ")
  "
  )
  
Key_User__c_soql<-gsub(pattern = "\n", replacement = "", x = Key_User__c_soql)
Object_Content <- sf_query(Key_User__c_soql)
Key_User__c_result <- rbind(Key_User__c_result,Object_Content)
Key_User__c_result<-Key_User__c_result %>% distinct()
# print(paste0('Iteration number ',j," out of ",j_max," is done. Time: ",format(Sys.time(), "%X") ))
j<-j+1
}  

  
Key_User__c_result_out <- Key_User__c_result %>% 
  select(Account__c, rebuild_LastModifiedDate__c) %>%
  group_by(Account__c) %>%
  summarise(Max_date=max(rebuild_LastModifiedDate__c))


  
  
  
###################################################################################
    
  
Case_result<-data.frame()
  j<-1
  j_max<-ceiling(length(ids)/i_rows)
  
for (i in seq(from=1,to=length(ids),by=i_rows)) {

  Case_soql<-paste0("
  SELECT AccountId, rebuild_LastModifiedDate__c
  FROM Case
  WHERE AccountId in (
  "
  ,paste0("'",paste0(na.omit(ids[i:(i+(i_rows-1))]), sep = "", collapse = "','"),"'"),
  ")
  "
  )
  
Case_soql<-gsub(pattern = "\n", replacement = "", x = Case_soql)
Object_Content <- sf_query(Case_soql)
Case_result <- rbind(Case_result,Object_Content)
Case_result<-Case_result %>% distinct()
# print(paste0('Iteration number ',j," out of ",j_max," is done. Time: ",format(Sys.time(), "%X") ))
j<-j+1
}  
  
  
  
Case_result_out <- Case_result %>% 
  select(AccountId, rebuild_LastModifiedDate__c) %>%
  group_by(AccountId) %>%
  summarise(Max_date=max(rebuild_LastModifiedDate__c))



###################################################################################
    
  
Lead_result<-data.frame()
  j<-1
  j_max<-ceiling(length(ids)/i_rows)
  
for (i in seq(from=1,to=length(ids),by=i_rows)) {

  Lead_soql<-paste0("
  SELECT ConvertedAccountId, rebuild_LastModifiedDate__c
  FROM Lead
  WHERE ConvertedAccountId in (
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
  select(ConvertedAccountId, rebuild_LastModifiedDate__c) %>%
  group_by(ConvertedAccountId) %>%
  summarise(Max_date=max(rebuild_LastModifiedDate__c))


###################################################################################
    
  
Opp_notWon_result<-data.frame()
  j<-1
  j_max<-ceiling(length(ids)/i_rows)
  
for (i in seq(from=1,to=length(ids),by=i_rows)) {

  Opp_notWon_soql<-paste0("
  SELECT AccountId, rebuild_LastModifiedDate__c
  FROM Opportunity
  WHERE 
  Current_Stage__c not in ('Closed Won', 'Closed Won (Contract Signed/Funds Received)')
  AND
  AccountId in (
  "
  ,paste0("'",paste0(na.omit(ids[i:(i+(i_rows-1))]), sep = "", collapse = "','"),"'"),
  ")
  "
  )
  
Opp_notWon_soql<-gsub(pattern = "\n", replacement = "", x = Opp_notWon_soql)
Object_Content <- sf_query(Opp_notWon_soql)
Opp_notWon_result <- rbind(Opp_notWon_result,Object_Content)
Opp_notWon_result<-Opp_notWon_result %>% distinct()
# print(paste0('Iteration number ',j," out of ",j_max," is done. Time: ",format(Sys.time(), "%X") ))
j<-j+1
}  
  
  
  
Opp_notWon_result_out <- Opp_notWon_result %>% 
  select(AccountId, rebuild_LastModifiedDate__c) %>%
  group_by(AccountId) %>%
  summarise(Max_date=max(rebuild_LastModifiedDate__c))


###################################################################################
    
  
Opp_Won_result<-data.frame()
  j<-1
  j_max<-ceiling(length(ids)/i_rows)
  
for (i in seq(from=1,to=length(ids),by=i_rows)) {

  Opp_Won_soql<-paste0("
  SELECT AccountId, rebuild_LastModifiedDate__c, Project_Start_Date__c, CloseDate, Expected_Project_Duration_Months__c
  FROM Opportunity
  WHERE Current_Stage__c in ('Closed Won', 'Closed Won (Contract Signed/Funds Received)')
  AND
  AccountId in (
  "
  ,paste0("'",paste0(na.omit(ids[i:(i+(i_rows-1))]), sep = "", collapse = "','"),"'"),
  ")
  "
  )
  
Opp_Won_soql<-gsub(pattern = "\n", replacement = "", x = Opp_Won_soql)
Object_Content <- sf_query(Opp_Won_soql)
Opp_Won_result <- rbind(Opp_Won_result,Object_Content)
Opp_Won_result<-Opp_Won_result %>% distinct()
# print(paste0('Iteration number ',j," out of ",j_max," is done. Time: ",format(Sys.time(), "%X") ))
j<-j+1
}  
  
# There was a problem with Project_Start_Date__c - automatically converted to number. Convert back.
Opp_Won_result$Project_Start_Date__c <- as.Date(Opp_Won_result$Project_Start_Date__c)

#####################################################################################################
# Take max date from extracted columns
Start_close <- apply(Opp_Won_result[,c(3,4)],1,max,na.rm=TRUE)
Start_close<-as.Date(Start_close)

Opp_Won_result[is.na(Opp_Won_result$Expected_Project_Duration_Months__c)==TRUE,"Expected_Project_Duration_Months__c"] <- 0


Opp_Won_result<-Opp_Won_result %>%
  select(AccountId,rebuild_LastModifiedDate__c,Expected_Project_Duration_Months__c)


Opp_Won_result<-data.frame(Opp_Won_result,Start_close)


for (i in 1:length(Opp_Won_result$AccountId)){
  month(Opp_Won_result$Start_close[i])<-month(Opp_Won_result$Start_close[i])+Opp_Won_result$Expected_Project_Duration_Months__c[i]
  }



#####################################################################################################


Opp_Won_result_work_max<-data.frame(Opp_Won_result$AccountId,Opp_Won_result$rebuild_LastModifiedDate__c,Opp_Won_result$Start_close)

# Take max date from extracted columns after additional work
Opp_Won_result_work_max<-apply(Opp_Won_result_work_max[,-1],1,max,na.rm=TRUE)
  
  
# Prepare Opp_Won_result_out data frame
Opp_Won_result_out <- data.frame(Opp_Won_result$AccountId,Opp_Won_result_work_max)
names(Opp_Won_result_out) <- c("AccountId","Max_date")
Opp_Won_result_out$Max_date<-as.Date(Opp_Won_result_out$Max_date)

  
Opp_Won_result_out <- Opp_Won_result_out %>% 
  select(AccountId, Max_date) %>%
  group_by(AccountId) %>%
  summarise(Max_date_distinct=max(Max_date))



###################################################################################

  
Opp_Part_notWon_result<-data.frame()
  j<-1
  j_max<-ceiling(length(ids)/i_rows)

for (i in seq(from=1,to=length(ids),by=i_rows)) {

  Opp_Part_notWon_soql<-paste0("
  SELECT Partner__c, Opportunity__r.rebuild_LastModifiedDate__c, Opportunity__r.CloseDate 
  from Opportunity_Partner__c 
  where
  Status__c != 'Withdrawn' and Opportunity__r.StageName not in ('Closed Won', 'Closed Won (Contract Signed/Funds Received)')
  and
  Partner__c in (
  "
  ,paste0("'",paste0(na.omit(ids[i:(i+(i_rows-1))]), sep = "", collapse = "','"),"'"),
  ")
  "
  )

Opp_Part_notWon_soql<-gsub(pattern = "\n", replacement = "", x = Opp_Part_notWon_soql)
Object_Content <- sf_query(Opp_Part_notWon_soql)
Opp_Part_notWon_result <- rbind(Opp_Part_notWon_result,Object_Content)
Opp_Part_notWon_result<-Opp_Part_notWon_result %>% distinct()
# print(paste0('Iteration number ',j," out of ",j_max," is done. Time: ",format(Sys.time(), "%X") ))
j<-j+1
}

  
# Take max date from extracted columns after additional work
Opp_Part_notWon_work_max<-apply(Opp_Part_notWon_result[,-1],1,max,na.rm=TRUE)
  
  
# Prepare Opp_Part_notWon_out data frame
Opp_Part_notWon_out <- data.frame(Opp_Part_notWon_result$Partner__c,Opp_Part_notWon_work_max)
names(Opp_Part_notWon_out) <- c("AccountId","Max_date")
Opp_Part_notWon_out$Max_date<-as.Date(Opp_Part_notWon_out$Max_date)
  

Opp_Part_notWon_out <- Opp_Part_notWon_out %>% 
  select(AccountId, Max_date) %>%
  group_by(AccountId) %>%
  summarise(Max_date_distinct=max(Max_date))



###################################################################################

  
Opp_Part_Won_result<-data.frame()
  j<-1
  j_max<-ceiling(length(ids)/i_rows)

for (i in seq(from=1,to=length(ids),by=i_rows)) {

  Opp_Part_Won_soql<-paste0("
  SELECT Partner__c, Opportunity__r.rebuild_LastModifiedDate__c, Opportunity__r.CloseDate, Opportunity__r.Project_Start_Date__c, Opportunity__r.Expected_Project_Duration_Months__c
  from Opportunity_Partner__c 
  where
  Status__c != 'Withdrawn' and Opportunity__r.StageName in ('Closed Won', 'Closed Won (Contract Signed/Funds Received)')
  and
  Partner__c in (
  "
  ,paste0("'",paste0(na.omit(ids[i:(i+(i_rows-1))]), sep = "", collapse = "','"),"'"),
  ")
  "
  )

Opp_Part_Won_soql<-gsub(pattern = "\n", replacement = "", x = Opp_Part_Won_soql)
Object_Content <- sf_query(Opp_Part_Won_soql)
Opp_Part_Won_result <- rbind(Opp_Part_Won_result,Object_Content)
Opp_Part_Won_result<-Opp_Part_Won_result %>% distinct()
# print(paste0('Iteration number ',j," out of ",j_max," is done. Time: ",format(Sys.time(), "%X") ))
j<-j+1
}


# Take max date from extracted columns
Start_close <- apply(Opp_Part_Won_result[,c(3,4)],1,max,na.rm=TRUE)
Start_close<-as.Date(Start_close)

Opp_Part_Won_result[is.na(Opp_Part_Won_result$Opportunity__r.Expected_Project_Duration_Months__c)==TRUE,"Opportunity__r.Expected_Project_Duration_Months__c"] <- 0


Opp_Part_Won_result<-Opp_Part_Won_result %>%
  select(Partner__c,Opportunity__r.rebuild_LastModifiedDate__c,Opportunity__r.Expected_Project_Duration_Months__c)


Opp_Part_Won_result<-data.frame(Opp_Part_Won_result,Start_close)



for (i in 1:length(Opp_Part_Won_result$Partner__c)){
  month(Opp_Part_Won_result$Start_close[i])<-month(Opp_Part_Won_result$Start_close[i])+Opp_Part_Won_result$Opportunity__r.Expected_Project_Duration_Months__c[i]
  }




#####################################################################################################


Opp_Part_Won_result_work_max<-data.frame(Opp_Part_Won_result$Partner__c,Opp_Part_Won_result$Opportunity__r.rebuild_LastModifiedDate__c,Opp_Part_Won_result$Start_close)

# Take max date from extracted columns after additional work
Opp_Part_Won_result_work_max<-apply(Opp_Part_Won_result_work_max[,-1],1,max,na.rm=TRUE)
  
  
# Prepare Opp_Won_result_out data frame
Opp_Part_Won_result_out <- data.frame(Opp_Part_Won_result$Partner__c,Opp_Part_Won_result_work_max)
names(Opp_Part_Won_result_out) <- c("AccountId","Max_date")
Opp_Part_Won_result_out$Max_date<-as.Date(Opp_Part_Won_result_out$Max_date)

  
Opp_Part_Won_result_out <- Opp_Part_Won_result_out %>% 
  select(AccountId, Max_date) %>%
  group_by(AccountId) %>%
  summarise(Max_date_distinct=max(Max_date))


###################################################################################

Opportunity_Competitor__c_result<-data.frame()
  j<-1
  j_max<-ceiling(length(ids)/i_rows)
  
for (i in seq(from=1,to=length(ids),by=i_rows)) {

  Opportunity_Competitor__c_soql<-paste0("
  SELECT Competitor_Name__c, rebuild_LastModifiedDate__c
  from Opportunity_Competitor__c
  WHERE Competitor_Name__c in (
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
  select(Competitor_Name__c, rebuild_LastModifiedDate__c) %>%
  group_by(Competitor_Name__c) %>%
  summarise(Max_date=max(rebuild_LastModifiedDate__c,na.rm = TRUE))  
  
  
###################################################################################


Opportunity_Partner__c_result<-data.frame()
  j<-1
  j_max<-ceiling(length(ids)/i_rows)
  
for (i in seq(from=1,to=length(ids),by=i_rows)) {

  Opportunity_Partner__c_soql<-paste0("
  SELECT Partner__c, rebuild_LastModifiedDate__c
  from Opportunity_Partner__c
  WHERE Partner__c in (
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
  select(Partner__c, rebuild_LastModifiedDate__c) %>%
  group_by(Partner__c) %>%
  summarise(Max_date=max(rebuild_LastModifiedDate__c,na.rm = TRUE))  
  
  

###################################################################################

Placement_Test_Attendance__c_result<-data.frame()
  j<-1
  j_max<-ceiling(length(ids)/i_rows)
  
for (i in seq(from=1,to=length(ids),by=i_rows)) {

  Placement_Test_Attendance__c_soql<-paste0("
  SELECT RelatedOpportunity__r.AccountId, Date_Of_Placement_Test__c, rebuild_LastModifiedDate__c
  from Placement_Test_Attendance__c
  WHERE RelatedOpportunity__r.AccountId in (
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
Placement_Test_Attendance__c_result_max<-apply(Placement_Test_Attendance__c_result[,-1],1,max,na.rm=TRUE)
  
  
# Prepare Opp_Won_result_out data frame
Placement_Test_Attendance__c_result_out <- data.frame(Placement_Test_Attendance__c_result$RelatedOpportunity__r.AccountId,Placement_Test_Attendance__c_result_max)
names(Placement_Test_Attendance__c_result_out) <- c("AccountId","Max_date")
Placement_Test_Attendance__c_result_out$Max_date<-as.Date(Placement_Test_Attendance__c_result_out$Max_date)
  
  
  
Placement_Test_Attendance__c_result_out <- Placement_Test_Attendance__c_result_out %>% 
  select(AccountId, Max_date) %>%
  group_by(AccountId) %>%
  summarise(Max_date_distinct=max(Max_date,na.rm = TRUE))  
  
  


  
###################################################################################

  
EmailMessageRelation_result<-data.frame()
  j<-1
  j_max<-ceiling(length(ids)/i_rows)

for (i in seq(from=1,to=length(ids),by=i_rows)) {

  EmailMessageRelation_soql<-paste0("
  SELECT RelationId, EmailMessage.CreatedDate
  FROM EmailMessageRelation
  WHERE
  relationid != NULL AND
  relationid in (SELECT personcontactid from account where id in (
  "
  ,paste0("'",paste0(na.omit(ids[i:(i+(i_rows-1))]), sep = "", collapse = "','"),"'"),
  "))
  "
  )

EmailMessageRelation_soql<-gsub(pattern = "\n", replacement = "", x = EmailMessageRelation_soql)
Object_Content <- sf_query(EmailMessageRelation_soql)
EmailMessageRelation_result <- rbind(EmailMessageRelation_result,Object_Content)
EmailMessageRelation_result<-EmailMessageRelation_result %>% distinct()
# print(paste0('Iteration number ',j," out of ",j_max," is done. Time: ",format(Sys.time(), "%X") ))
j<-j+1
}



Account_Contact_result<-data.frame()
  j<-1
  j_max<-ceiling(length(ids)/i_rows)

for (i in seq(from=1,to=length(ids),by=i_rows)) {

  Account_Contact_soql<-paste0("
  SELECT Id, PersonContactId
  FROM Account
  WHERE Id in (
  "
  ,paste0("'",paste0(na.omit(ids[i:(i+(i_rows-1))]), sep = "", collapse = "','"),"'"),
  ")
  "
  )

Account_Contact_soql<-gsub(pattern = "\n", replacement = "", x = Account_Contact_soql)
Object_Content <- sf_query(Account_Contact_soql)
Account_Contact_result <- rbind(Account_Contact_result,Object_Content)
Account_Contact_result<-Account_Contact_result %>% distinct()
# print(paste0('Iteration number ',j," out of ",j_max," is done. Time: ",format(Sys.time(), "%X") ))
j<-j+1
}



EmailMessageRelation_work <- EmailMessageRelation_result %>%
  select(RelationId, EmailMessage.CreatedDate) %>%
  group_by(RelationId) %>%
  summarise(Max_date=max(EmailMessage.CreatedDate))


EmailMessageRelation_out <- EmailMessageRelation_work %>%
  left_join(Account_Contact_result, by = c("RelationId"="PersonContactId"))

EmailMessageRelation_out<-data.frame(EmailMessageRelation_out$Id,EmailMessageRelation_out$Max_date)
names(EmailMessageRelation_out)<-c("Id","Max_date")

###################################################################################

Activity__c_result<-data.frame()
  j<-1
  j_max<-ceiling(length(ids)/i_rows)
  
for (i in seq(from=1,to=length(ids),by=i_rows)) {

  Activity__c_soql<-paste0("
  SELECT What_Account_Id__c, User_Last_Modified_Date__c, End_Date_Time__c
  from Activity__c
  WHERE What_Account_Id__c in (
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
Activity__c_result_out <- data.frame(Activity__c_result$What_Account_Id__c,Activity__c_result_max)
names(Activity__c_result_out) <- c("AccountId","Max_date")
Activity__c_result_out$Max_date<-as.Date(Activity__c_result_out$Max_date)
  
  
  
Activity__c_result_out <- Activity__c_result_out %>% 
  select(AccountId, Max_date) %>%
  group_by(AccountId) %>%
  summarise(Max_date_distinct=max(Max_date,na.rm = TRUE))  
  
  

###################################################################################
###################################################################################
###################################################################################
###################################################################################

# Let's make a merged dataframe

All_ids<-data.frame(ids)

Merged_out<-All_ids %>% 

  left_join(Account_result_out, by = c("ids"="Parent_Guardian__c")) %>%  
  left_join(Org_Account_result_out, by = c("ids"="AccountId")) %>%  
  left_join(LiveChatTranscript_result_out, by = c("ids"="AccountId")) %>%  
  left_join(Person_Account_result_out, by = c("ids"="Contact.AccountId")) %>%
  left_join(Event_result_out, by = c("ids"="AccountId")) %>%
  left_join(Key_User__c_result_out, by = c("ids"="Account__c")) %>%
  left_join(Case_result_out, by = c("ids"="AccountId")) %>%
  left_join(Lead_result_out, by = c("ids"="ConvertedAccountId")) %>%
  left_join(Opp_notWon_result_out, by = c("ids"="AccountId")) %>%
  left_join(Opp_Won_result_out, by = c("ids"="AccountId")) %>%
  left_join(Opp_Part_notWon_out, by = c("ids"="AccountId")) %>%
  left_join(Opp_Part_Won_result_out, by = c("ids"="AccountId")) %>%
  left_join(Opportunity_Competitor__c_result_out, by = c("ids"="Competitor_Name__c")) %>%
  left_join(Opportunity_Partner__c_result_out, by = c("ids"="Partner__c")) %>%
  left_join(Placement_Test_Attendance__c_result_out, by = c("ids"="AccountId")) %>%
  left_join(EmailMessageRelation_out, by = c("ids"="Id")) %>%
  left_join(Activity__c_result_out, by = c("ids"="AccountId"))


names(Merged_out)<-c("Ids","Acc Parent Guardian","ACR org","LiveChatTranscript","ACR person","Event","Key_User__c","Case","Lead","Opp_notWon","Opp_Won","Opp_Part_notWon","Opp_Part_Won","Opportunity_Competitor__c","Opportunity_Partner__c","Placement_Test_Attendance","EmailMessageRelation","Activity__c")


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

