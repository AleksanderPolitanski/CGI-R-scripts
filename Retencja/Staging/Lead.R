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
path<-"C:/Users/aleksander.politansk/Desktop/RFC/RFC 200/STAGING ENV WORK/LEAD/"

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

ids<-read.table(paste0(path, "Lead ids ALL.txt"), stringsAsFactors = FALSE)
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
  SELECT WhoId, ActivityDateTime, EndDateTime, rebuild_LastModifiedDate__c
  FROM Event
  WHERE WhoId in (
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
Event_result_out <- data.frame(Event_result$WhoId,Event_resultx)
names(Event_result_out) <- c("WhoId","Max_date")
# Event_result_out$Max_date<-stri_datetime_parse(Event_result_out$Max_date,format = "uuuu-MM-dd HH:mm:ss",
#   lenient = FALSE, tz = NULL, locale = NULL)
  

Event_result_out <- Event_result_out %>% 
  select(WhoId, Max_date) %>%
  group_by(WhoId) %>%
  summarise(Max_date_distinct=max(Max_date, na.rm = TRUE))


###################################################################################

Task_result<-data.frame()
  j<-1
  j_max<-ceiling(length(ids)/i_rows)
  
for (i in seq(from=1,to=length(ids),by=i_rows)) {

  Task_soql<-paste0("
  SELECT WhoId, rebuild_LastModifiedDate__c, ActivityDate
  FROM Task
  WHERE WhoId in (
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

# Prepare Task_result_out data frame
Task_result_out <- data.frame(Task_result$WhoId,Task_resultx)
names(Task_result_out) <- c("WhoId","Max_date")
##########Task_result_out$Max_date<-stri_datetime_parse(Task_result_out$Max_date,format = "uuuu-MM-dd HH:mm:ss",lenient = FALSE, tz = NULL, locale = NULL)

Task_result_out$Max_date<-as.character(Task_result_out$Max_date)  

Task_result_out <- Task_result_out %>% 
  select(WhoId, Max_date) %>%
  group_by(WhoId) %>%
  summarise(Max_date_distinct=max(Max_date, na.rm = TRUE))


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
  relationid in (
  "
  ,paste0("'",paste0(na.omit(ids[i:(i+(i_rows-1))]), sep = "", collapse = "','"),"'"),
  ")
  "
  )

EmailMessageRelation_soql<-gsub(pattern = "\n", replacement = "", x = EmailMessageRelation_soql)
Object_Content <- sf_query(EmailMessageRelation_soql)
EmailMessageRelation_result <- rbind(EmailMessageRelation_result,Object_Content)
EmailMessageRelation_result<-EmailMessageRelation_result %>% distinct()
# print(paste0('Iteration number ',j," out of ",j_max," is done. Time: ",format(Sys.time(), "%X") ))
j<-j+1
}




EmailMessageRelation_out <- EmailMessageRelation_result %>%
  select(RelationId, EmailMessage.CreatedDate) %>%
  group_by(RelationId) %>%
  summarise(Max_date=max(EmailMessage.CreatedDate, na.rm = TRUE))



###################################################################################
###################################################################################
###################################################################################
###################################################################################

# Let's make a merged dataframe

All_ids<-data.frame(ids)

# if(length(Event_result_out[[1]])>0){
#   
#   Merged_out<-All_ids %>% 
#   left_join(Event_result_out, by = c("ids"="WhoId")) %>%  
#   left_join(Task_result_out, by = c("ids"="WhoId")) %>%  
#   left_join(EmailMessageRelation_out, by = c("ids"="RelationId"))
# 
# 
#   names(Merged_out)<-c("Ids","Event","Task","Email Message Relation")
# 
# } else {
# 
#   Merged_out<-All_ids %>% 
#   left_join(Task_result_out, by = c("ids"="WhoId")) %>%  
#   left_join(EmailMessageRelation_out, by = c("ids"="RelationId"))
# 
# 
#   names(Merged_out)<-c("Ids","Task","Email Message Relation")
# 
# }


Merged_out<-All_ids %>% 
  left_join(Task_result_out, by = c("ids"="WhoId"))

# Merged_out$Max_date <- apply(Merged_out[,-1],1,max,na.rm=TRUE)


# dates_edit<-Merged_out$Max_date
# 
# 
# for (i in 1:length(dates_edit)){
#   if(is.na(dates_edit[i])==TRUE){
#     dates_edit[i]<-dates_edit[i]
#   } else if(nchar(dates_edit[i])<11){
#     dates_edit[i]<-paste0(dates_edit[i]," 00:00:00")
#   }
# }
# 
# 
# Merged_out$Max_date<-dates_edit

# This is crucial! If not done, R changes time according to time zone difference
Merged_out<-force_tz(Merged_out,tzone="GMT")


write.xlsx(x = Merged_out, file = paste0(path,"dates from children.xlsx"))



print(paste0('Finished at ',format(Sys.time(), "%X") ))
