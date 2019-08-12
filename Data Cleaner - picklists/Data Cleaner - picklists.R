## Data Cleaner - picklists check
## Author: Alek Politanski
## Date: 12.Aug.2019
######################################################################################################################################################################
############################# SET ALL VARIABLES ######################################################################################################################


# Set your path and file name
path<-"C:/Users/aleksander.politansk/Desktop/Data Cleaner - picklists/"

# File containing only columns with picklist values, with api names as headers
input_file_name<-"Picklists"

# Which salesforce enviornment are you using?
# "Sandbox" or "Production"
env<-"Sandbox"

# API name of salesforce object you want to query
object<-"Account"


######################################################################################################################################################################
######################################################################################################################################################################

# Load libraries

libs<-c("dplyr","httpuv","openxlsx","salesforcer","beepr")
lapply(libs, FUN=library, character.only=TRUE )


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

# Read file containing only columns with picklist values
picklist_input<-read.xlsx(xlsxFile = paste0(path,input_file_name,".xlsx"))

# Read object fields and field values from Salesforce
object_fields <- sf_describe_object_fields(object)

# save picklist selection options for given fields

picklist_values_list<-list()

for (i in 1:length(names(picklist_input)) ){
  
  picklist_values<-as.data.frame(
    object_fields %>% 
    filter(name == names(picklist_input)[i]) %>% 
    .$picklistValues
  )

  picklist_values <- picklist_values %>% select (value)

  picklist_values_list[i]<-data.frame(picklist_values)
  
  rm(picklist_values)
}


names(picklist_values_list)<-names(picklist_input)


###################################################################################

# Loop to check which values are aligned with Salesforce values

picklist_output<-picklist_input

for (i in 1:length(picklist_input)){

picklist_output[i+length(picklist_input)]<-picklist_input[[i]]%in%picklist_values_list[[i]]
  
}

# Set new column names

new_names<-names(picklist_output[1:(length(picklist_output)/2)])
new_names<-paste0(new_names," CHECK")

names(picklist_output)[((length(picklist_output)/2)+1):(length(picklist_output))] <- new_names



# Save workbook with conditional formatting

# Create R object, type workbook
wb <- createWorkbook()

# Create style for R object
negStyle <- createStyle(fontColour = "#9C0006", bgFill = "#FFC7CE")

# Create sheet named "Picklist_check" in R object, type workbook
addWorksheet(wb, "Picklist_check")

# Write data to R object, type workbook from dataframe picklist_output
writeData(wb, "Picklist_check", picklist_output, colNames=TRUE)

# Add conditional formating to chosen columns/rows to our R object, type workbook
conditionalFormatting(wb, "Picklist_check", cols = 1:dim(picklist_output)[2], rows = 1:(dim(picklist_output)[1]+1),type = "contains", rule="FALSE", style = negStyle)

# Add filter to our R object, type workbook
addFilter(wb, "Picklist_check", cols = 1:dim(picklist_output)[2], rows = 1)

# Save R object, type workbook to a given path (to a newly created folder)

creation_time<-gsub(x = format(Sys.time(), "%X"), pattern = "\\:", replacement = " ")
dir.create(path = paste0(path,"output at ",creation_time))

new_path<-paste0(path,"output at ",creation_time,"/")

saveWorkbook(wb, file = paste0(new_path,input_file_name," checked ",creation_time,".xlsx") )
write.xlsx(x = picklist_values_list, file = paste0(new_path,input_file_name," list.xlsx"), overwrite = FALSE)
