####TODO list
#create more hierarchy lookups e.g. health, census 2011 and 2021 (separate), ITL hierarchy, admin (previous years), Fire
#stretch - make it load in the correct year, looking for geogs that have changed - using full 9char code. look for entity and the changed geog code
#turn the above if into a function e.g. lookup_loader
#any other QA checks to test for?
#finish off the output formatting of the workbook - try and make it more generic by using functions
#wb - write tidy data column - case_when and collapse down. Codes and Data only.

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
      user_data <- read.csv(filepath)
      #user_data <- read_csv(filepath, show_col_types = FALSE, locale = readr::locale(encoding = "UTF-8"))
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
distinct_entities <- raw_data %>% distinct(ENTCD, .keep_all = FALSE)
#or use this in base R to create a vector
unique_entities <- unique(raw_data$ENTCD)


####Loading correct lookup####
#TODO two hierarchies scenario?
#TODO add code to select the correct year of the admin hierarchy

lookup_link <- case_when(unique_entities %in% admin_entities ~ admin_link,
                 unique_entities %in% health_entities ~ health_link)

unique_check <- length(unique(lookup_link)) == 1

if(unique_check == TRUE){
  lookup <- read.csv(unique(lookup_link))
  message("Lookup successfully loaded")
} else if(unique_check == FALSE){
  if(length(unique(lookup_link)) == 2 & admin_link %in% lookup_link){
    lookup_link <- setdiff(lookup_link, admin_link)
    lookup <- read.csv(lookup_link)
    message("Lookup successfully loaded")
  }
} else {
  message("No lookup found :(")
}

#TODO check column names and inputs
####Joining data to the lookup####
input_data_col_name <- "AREA21CD"
#join the raw_data to the appropriate lookup
raw_data_join <- left_join(lookup, raw_data, by = c("AREA21CD" = input_data_col_name))

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

anti_join_test <- anti_join_check(lookup, raw_data, input_data_col_name)
#TODO create an export of unmatched data

####Preparing data for output####

data_col_pos <- menu(colnames(raw_data), graphics = FALSE, title = "Select the column containing the data variable for publication")
data_col <- colnames(raw_data)[data_col_pos]
lookup_col_no <- ncol(lookup)

#remove excess columns
raw_data_join <- select(raw_data_join, 1:all_of(lookup_col_no), all_of(data_col))

output <- filter(raw_data_join, (AREA21CD %in% anti_join_test$AREA21CD) == FALSE)


####Excel export formatting####

#styles
hsbold <- createStyle(textDecoration = "Bold")
hs1 <- createStyle(fontSize = 18, textDecoration = "Bold")

#Using openxlsx to create and format a new xlsx workbook
wb <- createWorkbook() #create the workbook
modifyBaseFont(wb, fontSize = 12, fontColour = 'black', fontName = "Arial") #define the font, size & colour

#add sheets to the workbook
addWorksheet(wb, sheetName = "Cover Sheet", gridLines = TRUE)
addWorksheet(wb, sheetName = "Accessible Data", gridLines = TRUE)
addWorksheet(wb, sheetName = "Tidy Data", gridLines = TRUE)

#write data to sheet 1
writeData(wb, 1, "Listing geographical areas in tables - best practice examples", startCol = "A", startRow = 1)
writeData(wb, 1, "This spreadsheet contains two worksheets. Each includes a table that lists geographical areas.", startCol = "A", startRow = 2)
writeData(wb, 1, "One shows an accessible layout which meets the accessibility legislation that came into force in September 2020.", startCol = "A", startRow = 3)
writeData(wb, 1, "The other shows a 'tidy data' layout which is useful for machine readability. ", startCol = "A", startRow = 4)
addStyle(wb, 1, hs1, rows = 1, cols = 1)

#write data to sheet 2
writeData(wb, 2, "Layout of geographical areas - accessible version", startCol = "A", startRow = 1)
writeData(wb, 2, "This worksheet contains one table. The table contains some blank cells due to the layout of the geographical areas.", startCol = "A", startRow = 2)
writeData(wb, sheet = 2, output, startCol = "A", startRow = 3, headerStyle = hsbold)
addStyle(wb, 2, hs1, rows = 1, cols = 1)


#write the workbook
saveWorkbook(wb, "output/test_output_formatting.xlsx", overwrite = TRUE)

# ʕ·ᴥ·ʔ
# :)
