####TODO list
#create more hierarchy lookups e.g. health, census 2011 and 2021 (separate), ITL hierarchy, admin (previous years), Fire
#stretch - make it load in the correct year, looking for geogs that have changed - using full 9char code. look for entity and the changed geog code
#turn the above if into a function e.g. lookup_loader
#find out how to get R to prompt the user to define a column where the data is  - input_data_col_name
#any other QA checks to test for?
#try and fix the Ynys Mon char TODO. Follow it from the input to the output to check it's not the data itself to begin with.
#finish off the output formatting of the workbook - try and make it more generic by using functions

library(tidyverse)
library(readxl)
library(openxlsx)

####Reference Data: Entity Vectors and Hierachy Links####
admin_entities <- c("K02","K03", "K04", "E92", "E06", "E07", "E08", "E09", "E10", "E11", "E12", "E13", "W92", "W06", "S92", "S12", "N92", "N09")
health_entities <- c("E38", "E40", "E54")
itl_entities <- c("TLN", "TLM", "TLD", "TLC", "TLE", "TLL", "TLG", "TLF", "TLH", "TLJ", "TLI", "TLK")
census_entities <- c("E00", "S00", "W00", "N00", "E01", "S01", "W01", "E02", "S02", "W02")

admin_link <- "lookups/CTRY20_NAT20_RGN20_CTYUA20_LAD20_lookup.csv"
health_link <- "lookups/CTRY21_NAT21_NHSER21_STP21_CCG21_LAD21_lookup.csv"

#make a function to ask the user their filename using readline, load that data in, then do the ent_data
load_user_data <- function(){
  filepath <- readline(prompt = "Insert the filepath for your input data: \n")
  
  if(grepl(".csv", filepath) == TRUE){
    if(file.exists(filepath) == TRUE){
      user_data <- read_csv(filepath, show_col_types = FALSE)
      AREACD_pos <- menu(colnames(user_data), graphics = FALSE, title = "Select the column containing the GSS Geography Codes")
      AREACD <- colnames(user_data)[AREACD_pos]
      user_data <- user_data %>% mutate(ENTCD = substr(AREACD, 1, 3))
      message("Input data successfully loaded")
      return(user_data)
    } else if(file.exists(filepath) == FALSE){
      message("File not found in specified filepath")
    }
        
  } else if (grepl(".xls", filepath) == TRUE){
    if(file.exists(filepath) == TRUE){
      user_data <- read_excel(filepath)
      AREACD_pos <- menu(colnames(user_data), graphics = FALSE, title = "Select the column containing the GSS Geography Codes")
      AREACD <- colnames(user_data)[AREACD_pos]
      user_data <- user_data %>% mutate(ENTCD = substr(AREACD, 1, 3))
      message("Input data successfully loaded")
      return(user_data)
    } else if(file.exists(filepath) == FALSE){
      message("File not found in specified filepath")
    }
  } else {
    message("No data loaded. Please check filepath location is correct and extension type is .csv or .xlsx")
  }
}

raw_data <- load_user_data()

####Load and prepare data####
#raw_data <- read_csv("data/Happiness.csv")

#create a dummy df for health lookup test
raw_data <- data.frame(AREACD = c("K02","K03", "K04", "E92", "E06", "E07", "E08", "E09", "E38", "E40", "E54"), 
                          data = c(12, 3, 56, 20, 8, 4, 7, 17, 19, 5, 1))

#do unique on that col to get a list of all the entity codes that are used in our data
distinct_entities <- ent_data %>% distinct(ENTCD, .keep_all = FALSE)
#or use this in base R to create a vector
unique_entities <- unique(ent_data$ENTCD)


####Loading correct lookup####
#TODO two hierarchies scenario?
#TODO add code to select the correct year of the admin hierarchy

lookup_link <- case_when(unique_entities %in% admin_entities ~ admin_link,
                 unique_entities %in% health_entities ~ health_link)

unique_check <- length(unique(lookup_link)) == 1

if(unique_check == TRUE){
  lookup <- read_csv(lookup_link)
  message("Lookup successfully loaded")
} else if(unique_check == FALSE){
  if(length(unique(lookup_link)) == 2 & admin_link %in% lookup_link){
    lookup_link <- setdiff(lookup_link, admin_link)
    lookup <- read_csv(lookup_link)
    message("Lookup successfully loaded")
  }
} else {
  message("No lookup found :(")
}


####Joining data to the lookup####
input_data_col_name <- "AREA21CD"
#join the raw_data to the appropriate lookup
ent_data_join <- left_join(lookup, ent_data, by = c("AREA21CD" = input_data_col_name))

####QA Checks####
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


####Excel export formatting####
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
writeDataTable(wb, sheet = 2, output, startCol = "A", startRow = 3)
#TODO figure out how to deal with accented chars in Welsh names - UTF8 ? error - 
#"Error in stri_length(newStrs) : 
#invalid UTF-8 byte sequence detected; try calling stri_enc_toutf8()


#write the workbook
saveWorkbook(wb, "output/test_output_formatting.xlsx", overwrite = TRUE)

# ʕ·ᴥ·ʔ
# :)
