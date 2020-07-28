# Used packages
libs<-c("dplyr","httpuv","openxlsx","salesforcer","beepr","readr","RecordLinkage","glue")

# Install packages if needed
new.packages <- libs[!(libs %in% installed.packages()[,"Package"])]

if(length(new.packages)) {
  install.packages(new.packages)}

# Load packages
lapply(libs, FUN=library, character.only=TRUE )


###################################################################################

path<-"C:/Users/aleksander.politansk/Desktop/EU Rollout/Italy/"


data<-read.xlsx(glue({path},"Italy org names.xlsx"))


for(i in seq_along(data$Name) ) {
  
  compar_temp<-RecordLinkage::levenshteinDist(data$Name[i], data$Name)
  compar_temp<-data.frame(data$Name,compar_temp)
  names(compar_temp)<-c("Name","Score")
  compar_temp<-compar_temp %>% arrange(Score)
  compar_temp$Name<-as.character(compar_temp$Name)
  
  if(compar_temp$Score[2]<9){
    data$top1[i]<-compar_temp$Name[2]
    data$top2[i]<-compar_temp$Name[3]
    data$top3[i]<-compar_temp$Name[4]
    data$top4[i]<-compar_temp$Name[5]
    data$top5[i]<-compar_temp$Name[6]
    data$top6[i]<-compar_temp$Name[7]
  } else {
    data$top1[i]<-""
    data$top2[i]<-""
    data$top3[i]<-""
    data$top4[i]<-""
    data$top5[i]<-""
    data$top6[i]<-""
  }
  


}


write.xlsx(x = data, glue({path},"Italy org names similarities.xlsx"))
