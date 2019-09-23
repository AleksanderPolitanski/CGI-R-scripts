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

# Used packages
libs<-c("dplyr","httpuv","openxlsx","salesforcer","beepr","readr")

# Install packages if needed
new.packages <- libs[!(libs %in% installed.packages()[,"Package"])]

if(length(new.packages)) {
  install.packages(new.packages)}

# Load packages
lapply(libs, FUN=library, character.only=TRUE )


###################################################################################

ids<-read.table(paste0(path, "Id.txt"), stringsAsFactors = FALSE)
ids<-ids$V1
setwd(path)

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

  Object_Fields<-sf_describe_object_fields(object_name = object)
  Object_Fields<-Object_Fields[Object_Fields$queryByDistance=="false",]
  Object_Names<-Object_Fields$name

if (fields == "ALL") {

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
Object_Content <- sf_query(Object_Content_soql)
Query_result <- rbind(Query_result,Object_Content)
Query_result<-Query_result %>% distinct()
print(paste0('Iteration number ',j," out of ",j_max," is done. Time: ",format(Sys.time(), "%X") ))
j<-j+1
}

} else {
  
  
  # Check if fields are valid, if not - don't query them
  
  fields_checked<-gsub(pattern = " ",replacement = "",x = fields)
  fields_checked<-data.frame(unlist(strsplit(fields_checked, split=",")))
  names(fields_checked)<-"fields"
  fields_checked<-data.frame(fields_checked) %>% left_join(Object_Fields, by=c("fields"="name"))
  fields_checked<-fields_checked[(fields_checked$queryByDistance=="false")|(is.na(fields_checked$queryByDistance)==TRUE),"fields"]
  
  # Proceed with query
  
  Query_result<-data.frame()
  j<-1
  j_max<-ceiling(length(ids)/i_rows)
  
for (i in seq(from=1,to=length(ids),by=i_rows)) {

  Object_Content_soql<-paste0("
  SELECT ",
  paste(fields_checked,collapse = ","), "
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
Object_Content <- sf_query(Object_Content_soql)
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
write_excel_csv(x = Query_result, path = paste0(path, "Query result at ",creation_time,".csv"), na = "")
} else if (out == "csv") {
write_excel_csv(x = Query_result, path = paste0(path, "Query result at ",creation_time,".csv"), na = "")
} else if (out == "xlsx") {
write.xlsx(Query_result, file = paste0(path, "Query result at ",creation_time,".xlsx"))
}



###################################################################################
# Just a signal that program has finished running
beep(sound = 12)

# Set back working directory to usual one
infor<-data.frame(Sys.info())
us<-as.character(infor["user",])
setwd(paste0("C:/Users/",us,"/Documents/"))
