library(tidyverse)
library(readxl)
library(writexl)
library(openxlsx)

#get some data :)

raw_data <- read_csv("data/Happiness.csv")
LAD_lu <- read_xlsx("data/LAD22_CTY22_RGN22_CTRY22_UK_LU.xlsx")

names(raw_data)

#do we need to know what year the data is from? - changing geographies? Out of scope: pre-GSS codes.


#find the column(s) with entity codes inside it
#understand the entity codes within the data - where do we get this info from? Linked data API? Code history database? RGC?
#need to substr first 3 digits of the 9 digit code

#option 1a (single entity code)
region_1a <- raw_data %>% filter(str_detect(AREACD, "E12")) #How to change AREACD to all cols? 1:14 didn't work.

#option 1b (multiple entity codes)
lads_1b <- raw_data %>% filter(str_detect(AREACD, "W06|E09|E08|E07|E06|S12|N09"))

lads_1c <- raw_data %>% filter(str_detect(AREACD, str_c(LADs, collapse = "|")))

#TODO can we find a way to automatically detect entity codes across all unknown columns?

#combine entity codes into geographic levels = eg. LADS (where multiple entity codes are represented as a single geog)
####Hierarchy Vectors####
RGNs = c("E12")
CTYAUTHs = c("E10", "E11", "E61", "E47")
LADUAs = c("W06", "E09", "E08", "E07", "E06", "S12", "N09")
#?Are we going to do combined census geogs across countries?
#TODO where in the hierarchy do we want UAs (W06, E06, S12, N09) to sit?
#TODO Do we want to include Country level (Wales, Scotland, NI) to sit in RGNs?


#next task: write a function that will take in a bigger geography, find out what type of geog it is, then look in the hierarchy to find out what sits beneath
#and does that smaller geog exist in the raw_data ? If answer is yes, then look through each big entity and extract all the smaller ones that belong to it
#(using the LU)
#might need to create multiple functions

#Working with one code manually to begin with.
test_rgn <- "E12000001"

#Extracting the first 3 chars of the 9 char Geography code
entity_check <- substr(test_rgn, 1,3)

#Creating a function that will compare the extracted 3 chars from the entity(or many entities) with a specified hierarchy vector and return a TRUE/FALSE result
hierarchy_check <- function(entity, hierarchy_vector){
  if (entity %in% hierarchy_vector){
    return(TRUE)
  } 
  else {
    return(FALSE)
  }
}

#run the function to check a single entity in the entity_check object against the CTYAUTHs vector
hierarchy_results <- hierarchy_check(entity_check, RGNs)

#to check multiple entities (define them as a vector first!) and specify the hierarchy vector to check against
#lapply(entity_check, hierarchy_check, RGNs)

#function to check the higher level geography and return what lower level geographies are within it (from the hierarchy list)
smaller_geog_extractor <- function(hierarchy_results, entity_check, hierarchy_list){
  if(hierarchy_results == TRUE){
    return(unname(unlist(hierarchy_list[entity_check]))) #TODO error checking here so that it goes to the else statement if no data returned from list
  } else {
    message("Input entity not found in this geography level. No results returned.")
  }
}

#run the above function
lower_test <- smaller_geog_extractor(hierarchy_results, entity_check, hierarchy_list)


#the hierarchy lookup list of doom - we are ignoring CTYS for now because they're complicated
hierarchy_list <- list("E12" = c("W06", "E09", "E08", "E07", "E06", "S12", "N09"), #regions E12 to LADS
                       "E09" = "E05", # LADs to Wards
                       "E08" = "E05",
                       "E07" = "E05",
                       "E06" = "E05")

#create a filtered dataframe containing all codes matching
raw_filter <- filter(raw_data, str_detect(AREACD, str_c(lower_test, collapse = "|")))

#function exploration
entity_data_presence <- function(raw_filter){
  if(nrow(raw_filter) == 0){
    return(FALSE)
  } else {
    return(TRUE)
  }
}

#TODO Fix the filtering and TRUE/FALSE assessment in the following function. The function preceeding this is a bodge for now.
#check for the presence of LTLA codes in the AREACD field of the raw_data using a vector - must collapse down the vector to make it work.
# entity_data_presence <- function(raw_data, code_col, lower_geogs){
#   if(nrow(filter(raw_data, str_detect(code_col, str_c(lower_geogs, collapse = "|")))) == 0){
#     return(FALSE)
#     } else {
#       return(TRUE)
#   }
# }

#using the function above to check the raw_filter df
entity_data_presence(raw_filter)

#creating a filtered df of the lookup table returning all entities that match the specified code in test_rgn and ordering them ascending order.
LAD_filter <- filter(LAD_lu, RGN22CD == test_rgn) %>% 
  arrange(., LAD22CD)


####CODING THE GEOGRAPHICAL HIERARCHY INTO COLUMN/CELL ARRANGEMENT####

##function to take an entity code, check if there's something smaller than it in the file, if so, grab it/them, alphabetise them, add them to master df
# does this (more details!):
#Take the biggest geography entity code (eg region)
#Look in entity hierarchy lookup to see what smaller entities make up the bigger geography (eg. County/UA)
#For regions we need to define the display order because it's not alphabetical
#take the first region GSS code and make a new master df with this one row in it
#use some lookup to find which smaller entities fall within it, order them alphabetically.
#add these rows onto the bottom of the master dataframe we are going to be building and adding to. (we might need to add some new columns before adding)

#use this recursively until all of the entities are exhausted

#QA checks:
#make sure all entities in the original data are in the new df


####PREPARING THE OUTPUT####
#produce front info/metadata tab

#make the font right
#find out which bits need to be bolded
#make column names English not computer version

#make machine readable version of data (area codes jammed into one column but in correct order)

#Export xlsx file

#celebrate/cry if it doesn't work

####Notes####
#functions
add_one<- function(x){
  y <- x + 1
  
  return(y)
}

answer <- add_one(2)