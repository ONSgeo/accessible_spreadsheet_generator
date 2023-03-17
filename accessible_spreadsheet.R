####TODO list
#create more hierarchy lookups e.g. health, census 2011 and 2021 (separate), ITL hierarchy, admin (previous years), Fire
#Make the lookup_loader function load in the correct year, looking for geogs that have changed - using full 9char code. look for entity and the changed geog code
#finish off the output formatting of the workbook - try and make it more generic by using functions
#wb - write tidy data column - case_when and collapse down. Codes and Data only.
#move functions into another r script and tidy up this one for users.
#two hierarchies scenario? export and run the unjoined data through the process again to create another workbook? >:)
#create a function to create and export unmatched data (that doesn't join to the lookup)
#add export on the unmatched lookup check function


library(tidyverse)
library(readxl)
library(openxlsx)

####Load and prepare data####
#make a function to ask the user their filename using readline, load that data in, then do the ent_data
load_user_data <- function(){
  filepath <- readline(prompt = "Insert the filepath for your input data: \n")
  
  if(grepl(".csv", filepath) == TRUE){
    if(file.exists(filepath) == TRUE){
      user_data <- read.csv(filepath)
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

#raw_data <- read_csv("data/Happiness.csv")

#create a dummy df for health lookup test
#raw_data <- data.frame(AREACD = c("K02","K03", "K04", "E92", "E06", "E07", "E08", "E09", "E38", "E40", "E54"), 
                          #data = c(12, 3, 56, 20, 8, 4, 7, 17, 19, 5, 1))

create_unique_entities <- function(raw_data){
  unique_entities <- unique(raw_data$ENTCD) #create a vector of all unique entity codes in the input data
  return(unique_entities)
}

unique_entities <- create_unique_entities(raw_data)

####Loading correct lookup####

lookup_loader <- function(unique_entities){
  ##Reference Data: Entity Vectors and Hierachy Links##
  admin_entities <- c("K02","K03", "K04", "E92", "E06", "E07", "E08", "E09", "E10", "E11", "E12", "E13", "W92", "W06", "S92", "S12", "N92", "N09")
  health_entities <- c("E38", "E40", "E54")
  itl_entities <- c("TLN", "TLM", "TLD", "TLC", "TLE", "TLL", "TLG", "TLF", "TLH", "TLJ", "TLI", "TLK")
  census_entities <- c("E00", "S00", "W00", "N00", "E01", "S01", "W01", "E02", "S02", "W02")
  
  admin_link <- "lookups/CTRY20_NAT20_RGN20_CTYUA20_LAD20_lookup.csv"
  health_link <- "lookups/CTRY21_NAT21_NHSER21_STP21_CCG21_LAD21_lookup.csv"
  census_link <- "lookups/MSOA21_LSOA21_OA21_lookup.csv"
  itl_link <- "lookups/ITL121_ITL221_ITL321_lookup.csv"
  
  ##function begins here##
  #checks for the presence of entities contained in unique_entities, in the admin, health etc entity vectors
  #produces a vector with the correct lookup link
  lookup_link <- case_when(unique_entities %in% admin_entities ~ admin_link,
                           unique_entities %in% health_entities ~ health_link,
                           unique_entities %in% itl_entities ~ itl_link,
                           unique_entities %in% census_entities ~ census_link)
  
  #checks to see whether the lookup link is unique, or if there are multiple lookups selected
  unique_check <- length(unique(lookup_link)) == 1
  
  #if the lookup link is unique, then that lookup will be read in using the lookup_link
  #if not unique, then checks to see if admin is present with another lookup, if so then the alternative lookup will be selected
  if(unique_check == TRUE){
    lookup <- read.csv(unique(lookup_link))
    message("Lookup successfully loaded")
    return(lookup)
  } else if(unique_check == FALSE){
    if(length(unique(lookup_link)) == 2 & admin_link %in% lookup_link){
      lookup_link <- setdiff(lookup_link, admin_link)
      lookup <- read.csv(lookup_link)
      message("Lookup successfully loaded")
      return(lookup)
    } else if(length(unique(lookup_link)) >= 3 | (length(unique(lookup_link)) >= 2 & !(admin_link %in% lookup_link))){
      lookup_link_pos <- menu(unique(lookup_link), graphics = FALSE, title = "Multiple potential lookups found. Select the desired lookup from these options:")
      lookup_link <- unique(lookup_link)[lookup_link_pos] 
      lookup <- read.csv(lookup_link)
      message("Lookup successfully loaded")
      return(lookup)
    }
  } else {
    message("No matching lookup found :(")
  }
}

lookup <- lookup_loader(unique_entities)

####Joining data to the lookup####

data_joiner <- function(raw_data, lookup){
  input_data_col_name_pos <- menu(colnames(raw_data), graphics = FALSE, title = "Select the column containing the GSS Geography Codes in the input data:")
  input_data_col_name <- colnames(raw_data)[input_data_col_name_pos]
  lookup_col_name_pos <- menu(colnames(lookup), graphics = FALSE, title = "Select the column containing the GSS Geography Codes in the lookup:")
  lookup_col_name <- colnames(lookup)[lookup_col_name_pos]
  raw_data_join <- left_join(lookup, raw_data, by = setNames(input_data_col_name, lookup_col_name)) #using setnames inverts the input variable order
  message("Data successfully joined")
  return(raw_data_join)
}

raw_data_join <- data_joiner(raw_data, lookup) 

####QA Checks####
#QA checks to make sure everything has joined correctly - number of rows, what entities have NA in the Values column

#anti join - tell us what hasn't joined
unmatched_lookup_check <- function(lookup, raw_data){
  input_data_col_name_pos <- menu(colnames(raw_data), graphics = FALSE, title = "Select the column containing the GSS Geography Codes in the input data:")
  input_data_col_name <- colnames(raw_data)[input_data_col_name_pos]
  lookup_col_name_pos <- menu(colnames(lookup), graphics = FALSE, title = "Select the column containing the GSS Geography Codes in the lookup:")
  lookup_col_name <- colnames(lookup)[lookup_col_name_pos]
  
  anti_join_data <- anti_join(lookup, raw_data, by = setNames(input_data_col_name, lookup_col_name))  
  
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

anti_join_test <- unmatched_lookup_check(lookup, raw_data)


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
