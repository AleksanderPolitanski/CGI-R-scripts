## ID's translator
## Author: Alek Politanski
## Date: 19.Aug.2019
######################################################################################################################################################################
############################# SET ALL VARIABLES ######################################################################################################################

# Set your path and file name
path<-"C:/Users/aleksander.politansk/Desktop/R/ID translator/"

# Insert name of your file containing list of ID's (it has to be in txt format)
input_file_name<-"input"

######################################################################################################################################################################
######################################################################################################################################################################

# Load libraries
libs<-c("dplyr","openxlsx")
lapply(X = libs, FUN = library, character.only = TRUE)


###################################################################################

# Set working directory
setwd(dir = path)
# Read input file and make several formatting changes in it
short_ids<-read.csv(file = paste0(input_file_name,".txt"), header = FALSE)
short_ids<-as.data.frame(sapply(short_ids, FUN = as.character))
names(short_ids)<-"ids"

# Save each character of the input ID's and check if it's a capital letter
letters1 <- substr(short_ids$ids, 1, 1); letters1<-letters1 %in% LETTERS
letters2 <- substr(short_ids$ids, 2, 2); letters2<-letters2 %in% LETTERS
letters3 <- substr(short_ids$ids, 3, 3); letters3<-letters3 %in% LETTERS
letters4 <- substr(short_ids$ids, 4, 4); letters4<-letters4 %in% LETTERS
letters5 <- substr(short_ids$ids, 5, 5); letters5<-letters5 %in% LETTERS
letters6 <- substr(short_ids$ids, 6, 6); letters6<-letters6 %in% LETTERS
letters7 <- substr(short_ids$ids, 7, 7); letters7<-letters7 %in% LETTERS
letters8 <- substr(short_ids$ids, 8, 8); letters8<-letters8 %in% LETTERS
letters9 <- substr(short_ids$ids, 9, 9); letters9<-letters9 %in% LETTERS
letters10 <- substr(short_ids$ids, 10, 10); letters10<-letters10 %in% LETTERS
letters11 <- substr(short_ids$ids, 11, 11); letters11<-letters11 %in% LETTERS
letters12 <- substr(short_ids$ids, 12, 12); letters12<-letters12 %in% LETTERS
letters13 <- substr(short_ids$ids, 13, 13); letters13<-letters13 %in% LETTERS
letters14 <- substr(short_ids$ids, 14, 14); letters14<-letters14 %in% LETTERS
letters15 <- substr(short_ids$ids, 15, 15); letters15<-letters15 %in% LETTERS

# Assign 'points' for each occurance of capital letter in each character position
letters1<-as.data.frame(ifelse(letters1==TRUE, 1, 0)); names(letters1)<-"points"
letters2<-as.data.frame(ifelse(letters2==TRUE, 2, 0)); names(letters2)<-"points"
letters3<-as.data.frame(ifelse(letters3==TRUE, 4, 0)); names(letters3)<-"points"
letters4<-as.data.frame(ifelse(letters4==TRUE, 8, 0)); names(letters4)<-"points"
letters5<-as.data.frame(ifelse(letters5==TRUE, 16, 0)); names(letters5)<-"points"
letters6<-as.data.frame(ifelse(letters6==TRUE, 1, 0)); names(letters6)<-"points"
letters7<-as.data.frame(ifelse(letters7==TRUE, 2, 0)); names(letters7)<-"points"
letters8<-as.data.frame(ifelse(letters8==TRUE, 4, 0)); names(letters8)<-"points"
letters9<-as.data.frame(ifelse(letters9==TRUE, 8, 0)); names(letters9)<-"points"
letters10<-as.data.frame(ifelse(letters10==TRUE, 16, 0)); names(letters10)<-"points"
letters11<-as.data.frame(ifelse(letters11==TRUE, 1, 0)); names(letters11)<-"points"
letters12<-as.data.frame(ifelse(letters12==TRUE, 2, 0)); names(letters12)<-"points"
letters13<-as.data.frame(ifelse(letters13==TRUE, 4, 0)); names(letters13)<-"points"
letters14<-as.data.frame(ifelse(letters14==TRUE, 8, 0)); names(letters14)<-"points"
letters15<-as.data.frame(ifelse(letters15==TRUE, 16, 0)); names(letters15)<-"points"

# Sum up the 'points' for each of the new letters
points_for_letter_1<-letters1+letters2+letters3+letters4+letters5
points_for_letter_2<-letters6+letters7+letters8+letters9+letters10
points_for_letter_3<-letters11+letters12+letters13+letters14+letters15

# Create a dataframe that holds a key: which letter for wich amount of 'points'
key_n<-c(0:31)
key<-c("A","B","C","D","E","F","G","H","I","J","K","L","M","N","O","P","Q","R","S","T","U","V","W","X","Y","Z","0","1","2","3","4","5")

key_to_letters<-data.frame(key_n,key, stringsAsFactors = FALSE)
names(key_to_letters)<-c("points","key")

# Assign proper letters to be added on each record
letter_1<-points_for_letter_1 %>% left_join(key_to_letters)
letter_2<-points_for_letter_2 %>% left_join(key_to_letters)
letter_3<-points_for_letter_3 %>% left_join(key_to_letters)

# Add new letters to short ID's
long_ids<-paste(short_ids$ids,letter_1$key,letter_2$key,letter_3$key,sep = '')

# Make a dataframe with both short and long ID's and write it
short_and_long_ids<-data.frame(short_ids,long_ids,stringsAsFactors = FALSE)
write.xlsx(short_and_long_ids, file = "short to long.xlsx")

# Delete redundant variables form R session
rm(key,key_n,key_to_letters,letter_1,letter_2,letter_3,letters1,letters10,letters11,letters12,letters13,letters14,letters15,letters2,letters3,letters4,letters5,letters6,letters7,letters8,letters9,libs,long_ids,points_for_letter_1,points_for_letter_2,points_for_letter_3,short_ids)

# Set back working directory to a usual one
infor<-data.frame(Sys.info())
us<-as.character(infor["user",])
setwd(paste0("C:/Users/",us,"/Documents/"))
