## Loop to query chosen fields from one object, for a given list of IDs
## Author: Alek Politanski
## Date: 08.Aug.2019
######################################################################################################################################################################
############################# SET ALL VARIABLES ######################################################################################################################

# Path to your folder
path<-"C:/Users/aleksander.politansk/Desktop/Test folder/"

# Which salesforce enviornment are you using?
# "Sandbox" or "Production"
env<-"Sandbox"

# API name of salesforce object you want to query
object<-"Key_User__c"

# Fields you want to query. If all: "ALL". Example: fields<-"Id, Name, Owner.Name"
fields<-"Id, Name, Owner.Name"

# How many rows you want to query in one iteration?
i_rows<-300

# What kind of output you need?
# "xlsx" - only xlsx file
# "csv" - only csv file
# "both" - xlsx and csv files
out<-"both"

# NOTE: PREPARE Id.txt file in your folder with Ids you want to query [format: without commas, without quotes, one id in row]

######################################################################################################################################################################
######################################################################################################################################################################

# Load libraries

libs<-c("dplyr","httpuv","openxlsx","salesforcer","beepr")
lapply(libs, FUN=library, character.only=TRUE )


###################################################################################

ids<-read.table(paste0(path, "Id.txt"), stringsAsFactors = FALSE)
ids<-ids$V1


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


if (fields == "ALL") {
  Object_Fields<-sf_describe_object_fields(object_name = object)
  Object_Names<-Object_Fields$name
  
  Query_result<-data.frame()
  j<-1
  j_max<-ceiling(length(ids)/i_rows)
  
for (i in seq(from=1,to=length(ids),by=i_rows)) {

  Object_Content_soql<-paste0("
  SELECT ",
  paste(Object_Names, sep = "", collapse = ','), "
  FROM ",
  object, "
  WHERE
  Id in (
  "
  ,paste0("'",paste0(na.omit(ids[i:(i+(i_rows-1))]), sep = "", collapse = "','"),"'"),
  "
  )
  "
  )
  
Object_Content_soql<-gsub(pattern = "\n", replacement = "", x = Object_Content_soql)
Object_Content <- sf_query(Object_Content_soql, api_type = "Bulk 1.0", object_name = object)
Query_result <- rbind(Query_result,Object_Content)
Query_result<-Query_result %>% distinct()
print(paste0('Iteration number ',j," out of ",j_max," is done. Time: ",format(Sys.time(), "%X") ))
j<-j+1
}

} else {
  
  Query_result<-data.frame()
  j<-1
  j_max<-ceiling(length(ids)/i_rows)
  
for (i in seq(from=1,to=length(ids),by=i_rows)) {

  Object_Content_soql<-paste0("
  SELECT ",
  paste(fields), "
  FROM ",
  object, "
  WHERE
  Id in (
  "
  ,paste0("'",paste0(na.omit(ids[i:(i+(i_rows-1))]), sep = "", collapse = "','"),"'"),
  "
  )
  "
  )
  
Object_Content_soql<-gsub(pattern = "\n", replacement = "", x = Object_Content_soql)
Object_Content <- sf_query(Object_Content_soql, api_type = "Bulk 1.0", object_name = object)
Query_result <- rbind(Query_result,Object_Content)
Query_result<-Query_result %>% distinct()
print(paste0('Iteration number ',j," out of ",j_max," is done. Time: ",format(Sys.time(), "%X") ))
j<-j+1

}
  
  
}


###################################################################################
# Write chosen output

creation_time<-gsub(x = format(Sys.time(), "%X"), pattern = "\\:", replacement = " ")

if(out == "both"){
write.xlsx(Query_result, file = paste0(path, "Query result at ",creation_time,".xlsx"))
write.csv(x = Query_result, file = paste0(path, "Query result at ",creation_time,".csv"), quote = TRUE, row.names = FALSE, fileEncoding = "UTF-8")
} else if (out == "csv") {
write.csv(x = Query_result, file = paste0(path, "Query result at ",creation_time,".csv"), quote = TRUE, row.names = FALSE, fileEncoding = "UTF-8")
} else if (out == "xlsx") {
write.xlsx(Query_result, file = paste0(path, "Query result at ",creation_time,".xlsx"))
}



###################################################################################
# Just a signal that program has finished running
beep(sound = 12)
