library(tidyverse)
library(readxl)
library(openxlsx)

####Load in input data####
raw_data <- read_csv("data/Happiness.csv")

#using head of data for testing, because of special o character Ynys Mon
raw_head <- head(raw_data)

#create a dummy df for health lookup test
raw_data <- data.frame(AREACD = c("K02","K03", "K04", "E92", "E06", "E07", "E08", "E09", "E38", "E40", "E54"), 
                          data = c(12, 3, 56, 20, 8, 4, 7, 17, 19, 5, 1))



####Preparing data layout####

#create a new col in raw_data with the first 3 char of the entity code in it
ent_data <- raw_data %>% mutate(ENTCD = substr(AREACD, 1, 3))

#do unique on that col to get a list of all the entity codes that are used in our data
distinct_entities <- ent_data %>% distinct(ENTCD, .keep_all = FALSE)
#or use this in base R to create a vector
unique_entities <- unique(ent_data$ENTCD)

#TODO See if we can add a successfully loaded message to the ifelse statements in this function

# lookup_loader <- function(unique_entities){
#   if(exists("lookup_link") == FALSE){
#     lookup_link <- ifelse(unique_entities %in% c("K02","K03", "K04", "E92", "E06", "E07", "E08", "E09", "E10", "E11", "E12", "E13", "W92", "W06", "S92", "S12", "N92", "N09"), 
#                         "lookups/CTRY20_NAT20_RGN20_CTYUA20_LAD20_lookup.csv", message(""))
#     message("Loading Admin lookup")
#     } else if(exists("lookup_link") == FALSE){
#     lookup_link <- ifelse(unique_entities %in% c("K02","K03", "K04", "E92", "E06", "E07", "E08", "E09", "E38", "E40", "E54"), 
#                         "lookups/CTRY21_NAT21_NHSER21_STP21_CCG21_LAD21_lookup.csv", message(""))
#     message("Loading Health lookup")
#     } else {
#     message("No Lookups found")
#   }
#   
#   #lookup <- read_csv(lookup_link)
#   #return(lookup)
#   return(lookup_link)
# }
# 
# lookup <- lookup_loader(unique_entities)


#experimenting with case_when
lookup_link <- case_when(unique_entities %in% c("K02","K03", "K04", "E92", "E06", "E07", "E08", "E09", "E10", "E11", "E12", "E13", "W92", "W06", "S92", "S12", "N92", "N09") ~ "lookups/CTRY20_NAT20_RGN20_CTYUA20_LAD20_lookup.csv",
                 unique_entities %in% c("K02","K03", "K04", "E92", "E06", "E07", "E08", "E09", "E38", "E40", "E54") ~ "lookups/CTRY21_NAT21_NHSER21_STP21_CCG21_LAD21_lookup.csv")

unique_check <- length(unique(lookup_link)) == 1

if(unique_check == TRUE){
  lookup <- read_csv(lookup_link)
  message("Lookup successfully loaded")
} else if(unique_check == FALSE){
  message("Write more code to find the correct lookup")
} else {
  message("No lookup found :(")
}

#create more hierarchy lookups e.g. health, census 2011 and 2021 (separate), ITL hierarchy, admin (previous years), Fire
#add them into the if statement above. Use else if and add message stating which hierarchy lookup is being loaded.
#stretch - make it load in the correct year, looking for geogs that have changed - using full 9char code. look for entity and the changed geog code
#turn the above if into a function e.g. lookup_loader
#find out how to get R to prompt the user to define a column where the data is  - input_data_col_name
#try and fix the TODO on line 68
#any other QA checks to test for?
#try and fix the Ynys Mon char TODO. Follow it from the input to the output to check it's not the data itself to begin with.
#finish off the output formatting of the workbook - try and make it more generic by using functions

input_data_col_name <- "AREA21CD"
#join the raw_data to the appropriate lookup
ent_data_join <- left_join(lookup, ent_data, by = c("AREA21CD" = input_data_col_name))

#QA checks to make sure everything has joined correctly - number of rows, what entities have NA in the Values column

#anti join - tell us what hasn't joined
anti_join_check <- function(lookup, input_data, input_data_col_name){
  anti_join_data <- anti_join(lookup, input_data, by = c("AREA21CD" = input_data_col_name))  
  
  if(nrow(anti_join_data) > 0){
    message("Unmatched rows found. Check output results prior to proceeding.")
    message("Printing output results to console.")
    print(anti_join_data)
    return(anti_join_data)
  } else if (nrow(anti_join_data) == 0){
    message("No unmatched rows found :)") 
  } else {
    message("Unable to check. Please review inputs.")
  }
}

test <- anti_join_check(lookup, ent_data, input_data_col_name)

output <- filter(ent_data_join, !input_data_col_name %in% test$AREA21CD)
#TODO fix this row above!

#prep the data for export
# :)

####Preparing for data export####
#Using openxlsx to create and format a new xlsx workbook
wb <- createWorkbook() #create the workbook
modifyBaseFont(wb, fontSize = 14, fontColour = 'black', fontName = "Arial") #define the font, size & colour

#add sheets to the workbook
addWorksheet(wb, sheetName = "Cover Sheet", gridLines = TRUE)
addWorksheet(wb, sheetName = "Accessible Data", gridLines = TRUE)
addWorksheet(wb, sheetName = "Tidy Data", gridLines = TRUE)

#write data to sheet 1
writeData(wb, 1, "Listing geographical areas in tables - best practice examples", startCol = "A", startRow = 1)
writeData(wb, 1, "This spreadsheet contains two worksheets. Each includes a table that lists geographical areas.", startCol = "A", startRow = 2)
writeData(wb, 1, "One shows an accessible layout which meets the accessibility legislation that came into force in September 2020.", startCol = "A", startRow = 3)
writeData(wb, 1, "The other shows a 'tidy data' layout which is useful for machine readability. ", startCol = "A", startRow = 4)


#write data to sheet 2
writeData(wb, 2, "Layout of geographical areas - accessible version", startCol = "A", startRow = 1)
writeData(wb, 2, "This worksheet contains one table. The table contains some blank cells due to the layout of the geographical areas.", startCol = "A", startRow = 2)
writeDataTable(wb, sheet = 2, raw_head, startCol = "A", startRow = 3)
#TODO figure out how to deal with accented chars in Welsh names - UTF8 ? error - 
#"Error in stri_length(newStrs) : 
#invalid UTF-8 byte sequence detected; try calling stri_enc_toutf8()


#write the workbook
saveWorkbook(wb, "output/test_output_formatting.xlsx", overwrite = TRUE)
