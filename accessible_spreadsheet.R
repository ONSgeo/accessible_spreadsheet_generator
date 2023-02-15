#packages
install.packages("writexl")

library(tidyverse)
library(readxl)
library(writexl)

#get some data :)

setwd("D:/Helen/Geospatial_Team_2022/Accessible_spreadsheets/data/")

raw_data <- read_csv("Happiness.csv")
LAD_lu <- read_xlsx("LAD22_CTY22_RGN22_CTRY22_UK_LU.xlsx")

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
RGNs = c("E12")
CTYUAs = c()
LADs = c("W06", "E09", "E08", "E07", "E06", "S12", "N09")
#?Are we going to do combined census geogs across countries?

#next task: write a function that will take in a bigger geography, find out what type of geog it is, then look in the hierarchy to find out what sits beneath
#and does that smaller geog exist in the raw_data ? If answer is yes, then look through each big entity and extract all the smaller ones that belong to it
#(using the LU)
#might need to create multiple functions

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

#functions
add_one<- function(x){
  y <- x + 1
  
  return(y)
}

answer <- add_one(2)


