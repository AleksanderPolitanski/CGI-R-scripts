# Used packages
libs<-c("dplyr","httpuv","openxlsx","salesforcer","beepr","readr")

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
###################################################################################
###################################################################################
###################################################################################

# Let's make a merged dataframe

All_ids<-data.frame(ids)

Merged_out<-All_ids %>% 
  left_join(Org_Account_result_out, by = c("ids"="AccountId")) %>%  
  left_join(Person_Account_result_out, by = c("ids"="Contact.AccountId")) %>%
  left_join(Event_result_out, by = c("ids"="AccountId")) %>%
  left_join(Key_User__c_result_out, by = c("ids"="Account__c")) %>%
  left_join(Case_result_out, by = c("ids"="AccountId"))


names(Merged_out)<-c("Ids","ACR org","ACR person","Event","Key_User__c","Case_result_out")


Merged_out$Max_date <- apply(Merged_out[,-1],1,max,na.rm=TRUE)


write.xlsx(x = Merged_out, file = paste0(path,"dates from children.xlsx"))



print(paste0('Finished at ',format(Sys.time(), "%X") ))