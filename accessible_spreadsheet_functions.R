#Accessible Spreadsheet Generator 2023
#This script accompanies the accessible_spreadsheet.R script
#Developed by Helen Calder and Heather Porter, March 2023 
#ONS Geospatial - Contact: geospatial@ons.gov.uk

#library packages
library(tidyverse)
library(readxl)
library(openxlsx)
library(janitor)

####Load and prepare data####
#make a function to ask the user their filename using readline, load that data in, then extract the first 3 char of the Geog entity code into new col - ENTCD
load_user_data <- function(){
  filepath <- readline(prompt = "Insert the filepath for your input data: \n")
  
  if(grepl(".csv", filepath) == TRUE){ 
    if(file.exists(filepath) == TRUE){
      user_data <- read.csv(filepath)
      GSSCD_pos <- menu(colnames(user_data), graphics = FALSE, title = "Select the column containing the GSS Geography Codes")
      GSSCD <- colnames(user_data)[GSSCD_pos] #converting the position of the GSSCD_pos into the column name instead - This variable must not be named the same as the code column, else it will break
      user_data <- user_data %>% mutate(ENTCD = str_sub(.[,GSSCD], 1, 3)) #keeps the first 3 characters of the GSSCD and creates new col - ENTCD
      message("Input data successfully loaded")
      return(user_data)
    } else if(file.exists(filepath) == FALSE){
      message("File not found in specified filepath")
    }
    
  } else if (grepl(".xls", filepath) == TRUE){
    if(file.exists(filepath) == TRUE){
      user_data <- read_excel(filepath)
      GSSCD_pos <- menu(colnames(user_data), graphics = FALSE, title = "Select the column containing the GSS Geography Codes")
      GSSCD <- colnames(user_data)[GSSCD_pos]
      user_data <- user_data %>% mutate(ENTCD = str_sub(.[,GSSCD], 1, 3))
      message("Input data successfully loaded")
      return(user_data)
    } else if(file.exists(filepath) == FALSE){
      message("File not found in specified filepath")
    }
  } else {
    message("No data loaded. Please check filepath location is correct and extension type is .csv or .xlsx")
  }
}

####Create unique entities####
#Create a vector of unique entities in the user data
create_unique_entities <- function(raw_data){
  unique_entities <- unique(raw_data$ENTCD) #create a vector of all unique entity codes in the input data
  return(unique_entities)
}

####Loading correct lookup####
#Cross references the unique entities against the known entity vectors and assigns the correct lookup link. Then loads that lookup if appropriate.
lookup_loader <- function(raw_data, unique_entities){
  ##Reference Data: Entity Vectors and Hierachy Links##
  admin_entities <- c("K02","K03", "K04", "E92", "E06", "E07", "E08", "E09", "E10", "E11", "E12", "E13", "W92", "W06", "S92", "S12", "N92", "N09")
  health_entities <- c("E38", "E40", "E54")
  itl_entities <- c("TLN", "TLM", "TLD", "TLC", "TLE", "TLL", "TLG", "TLF", "TLH", "TLJ", "TLI", "TLK")
  census_entities <- c("E00", "S00", "W00", "N00", "E01", "S01", "W01", "E02", "S02", "W02")
  
  admin_18_link <- "lookups/CTRY18_NAT18_RGN18_CTYUA18_LAD18_lookup.csv"
  admin_19_link <- "lookups/CTRY19_NAT19_RGN19_CTYUA20_LAD19_lookup.csv"
  admin_20_link <- "lookups/CTRY20_NAT20_RGN20_CTYUA20_LAD20_lookup.csv"
  admin_21_link <- "lookups/CTRY21_NAT21_RGN21_CTYUA21_LAD21_lookup.csv"
  admin_22_link <- "lookups/CTRY22_NAT22_RGN22_CTYUA22_LAD22_lookup.csv"
  admin_link <- c(admin_18_link, admin_19_link, admin_20_link, admin_21_link, admin_22_link) #make sure all admin lookups are included in this vector
  health_link <- "lookups/CTRY21_NAT21_NHSER21_STP21_CCG21_LAD21_lookup.csv"
  census_link <- "lookups/MSOA21_LSOA21_OA21_lookup.csv"
  itl_link <- "lookups/CTRY21_NAT21_ITL121_ITL221_ITL321_lookup.csv"
  
  ##function actions begin here##
  admin_code_pos <- menu(colnames(raw_data), graphics = FALSE, title = "Select the column containing the GSS Geography Codes in the input data:")
  admin_code_name <- colnames(raw_data)[admin_code_pos]
  #checks for the presence of entities contained in unique_entities, in the admin, health etc entity vectors
  #produces a vector with the associated lookup links
  #when adding new years of lookups, they **MUST** be above the older ones of the same geography (case_when works from top to bottom - newest lookup to oldest)
  #hierarchies that have multiple years associated with them are selected by individual GSS codes that underwent changes in that particular year
  lookup_link <- case_when(unique_entities %in% admin_entities ~ admin_22_link,
                           (unique_entities %in% admin_entities & ("E06000061" %in% raw_data[,admin_code_name]|"E06000062" %in% raw_data[,admin_code_name])) ~ admin_21_link,
                           (unique_entities %in% admin_entities & ("E06000060" %in% raw_data[,admin_code_name])) ~ admin_20_link,
                           (unique_entities %in% admin_entities & ("E10000002" %in% raw_data[,admin_code_name])) ~ admin_19_link,
                           (unique_entities %in% admin_entities & ("E06000029" %in% raw_data[,admin_code_name]|"E06000028" %in% raw_data[,admin_code_name])) ~ admin_18_link,
                           unique_entities %in% health_entities ~ health_link,
                           unique_entities %in% itl_entities ~ itl_link,
                           unique_entities %in% census_entities ~ census_link) %>% discard(is.na)
  
  #checks to see whether the lookup link is unique, or if there are multiple lookups selected
  unique_check <- length(unique(lookup_link)) == 1
  
  #if the lookup link is unique, then that lookup will be read in using the lookup_link
  #if not unique, then checks to see if admin is present with another lookup, if so then the alternative lookup will be selected
  #this is because admin geographies are commonly included with other geographies
  if(unique_check == TRUE){
    lookup <- read.csv(unique(lookup_link))
    message("Lookup successfully loaded")
    return(lookup)
  } else if(unique_check == FALSE){
    if(length(unique(lookup_link)) == 2 & admin_link %in% lookup_link){
      lookup_link <- setdiff(lookup_link, admin_link) #setdiff removes the admin_link
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

####Define column names for later use####
#used in the data joiner and QA checks
define_col_names <- function(lookup, raw_data){
  input_data_col_name_pos <- menu(colnames(raw_data), graphics = FALSE, title = "Select the column containing the GSS Geography Codes in the input data:")
  input_data_col_name <- colnames(raw_data)[input_data_col_name_pos]
  lookup_col_name_pos <- menu(colnames(lookup), graphics = FALSE, title = "Select the column containing the GSS Geography Codes in the lookup:")
  lookup_col_name <- colnames(lookup)[lookup_col_name_pos]
  return(list(input_data_col_name, lookup_col_name))
} 

####Joining data to the lookup####
data_joiner <- function(raw_data, lookup, input_data_col_name, lookup_col_name){
  raw_data_join <- left_join(lookup, raw_data, by = setNames(input_data_col_name, lookup_col_name)) #using setnames inverts the input variable order
  message("Data successfully joined")
  return(raw_data_join)
}

####QA Checks####
#QA checks to make sure everything has joined correctly - number of rows, what entities have NA in the Values column

#unmatched_lookup_check - tell us what rows in the lookup don't have data joined to them
unmatched_lookup_check <- function(lookup, raw_data, input_data_col_name, lookup_col_name){
  
  anti_join_data <- anti_join(lookup, raw_data, by = setNames(input_data_col_name, lookup_col_name))  
  
  if(nrow(anti_join_data) > 0){
    message("Unmatched rows in the lookup found. Check output results prior to proceeding.")
    return(anti_join_data)
  } else if (nrow(anti_join_data) == 0){
    message("No unmatched rows found :)") 
  } else {
    message("Unable to check. Please review inputs.")
  }
}

#unmatched_data_check - tell us what rows in the data aren't found in the lookup
unmatched_data_check <- function(lookup, raw_data, input_data_col_name, lookup_col_name){
  
  anti_join_data <- anti_join(raw_data, lookup, by = setNames(lookup_col_name, input_data_col_name))  
  
  if(nrow(anti_join_data) > 0){
    message("Unmatched data found. Check output results prior to proceeding.")
    return(anti_join_data)
  } else if (nrow(anti_join_data) == 0){
    message("No unmatched data found :)") 
  } else {
    message("Unable to check. Please review inputs.")
  }
}

#function to export unmatched QA check results
export_test_results <- function(unmatched_lookup_test, unmatched_data_test){
  timestamp <- format(Sys.time(), "%Y%m%d_%H%M")
  write.csv(unmatched_lookup_test, paste0("output/unmatched_lookup_results_", timestamp, ".csv"), row.names = FALSE)
  write.csv(unmatched_data_test, paste0("output/unmatched_data_results_", timestamp, ".csv"), row.names = FALSE)
}

####Preparing data for output####
#User selects the data variable column and excess columns are removed. Only code, names and data are kept.
output_preparation <- function(raw_data_join, lookup, unmatched_lookup_test, input_data_col_name, lookup_col_name){
  #preparing the human readable data
  data_col_pos <- menu(colnames(raw_data_join), graphics = FALSE, title = "Select the column containing the data variable for publication")
  data_col <- colnames(raw_data_join)[data_col_pos]
  lookup_col_no <- ncol(lookup)
  raw_data_join <- select(raw_data_join, 1:all_of(lookup_col_no), all_of(data_col)) #remove excess columns
  unmatched_lookup_codes <- unmatched_lookup_test[,lookup_col_name] #prepares unmatched lookup codes for use in next line - to be filtered out
  human_output <- filter(raw_data_join, (raw_data_join[,lookup_col_name] %in% unmatched_lookup_codes) == FALSE) %>% clean_names(case = "title")
  #preparing the tidy data sheet
  machine_output <- human_output %>% unite("Names", 2:all_of(lookup_col_no), sep = "", remove = TRUE, na.rm = TRUE) %>% clean_names()
  #above line compresses names into one column, removes underscores and NA values
  message("Data ready for output.")
  message("Reminder: Run QA checks before exporting to file to ensure all data is accounted for")
  return(list(human_output, machine_output)) #user unlists the outputs in the main script
}

####Excel export formatting####
#All the required formatting for accessibility. Three worksheets are created.
export_workbook <- function(human_output, machine_output){
  #styles
  hsbold <- createStyle(textDecoration = "Bold")
  hs1 <- createStyle(fontSize = 18, textDecoration = "Bold")
  
  #Using openxlsx to create and format a new xlsx workbook
  wb_title <- readline(prompt = "Input the title of your workbook for Excel title metadata: ") #(In Excel: File > Info > Title)
  wb <- createWorkbook(title = wb_title) #create the workbook
  modifyBaseFont(wb, fontSize = 12, fontColour = 'black', fontName = "Arial") #define the font, size & colour
  
  #add sheets to the workbook and names them
  addWorksheet(wb, sheetName = "Cover Sheet", gridLines = TRUE)
  addWorksheet(wb, sheetName = "Accessible Data", gridLines = TRUE)
  addWorksheet(wb, sheetName = "Tidy Data", gridLines = TRUE) #user requirement - nothing to do with tidy data in the tidyverse!
  
  #write data to sheet 1
  writeData(wb, 1, "Listing geographical areas in tables - best practice examples", startCol = "A", startRow = 1)
  writeData(wb, 1, "This spreadsheet contains two worksheets. Each includes a table that lists geographical areas.", startCol = "A", startRow = 2)
  writeData(wb, 1, "One shows an accessible layout which meets the accessibility legislation that came into force in September 2020.", startCol = "A", startRow = 3)
  writeData(wb, 1, "The other shows a 'tidy data' layout which is useful for machine readability. ", startCol = "A", startRow = 4)
  addStyle(wb, 1, hs1, rows = 1, cols = 1) #applies to A1 as a header
  
  #write data to sheet 2
  writeData(wb, 2, "Layout of geographical areas - accessible version", startCol = "A", startRow = 1)
  writeData(wb, 2, "This worksheet contains one table. The table contains some blank cells due to the layout of the geographical areas.", startCol = "A", startRow = 2)
  writeData(wb, sheet = 2, human_output, startCol = "A", startRow = 3, headerStyle = hsbold)
  addStyle(wb, 2, hs1, rows = 1, cols = 1)
  
  #write machine data to sheet 3 - known as Tidy data but do not confuse with tidyverse!
  writeData(wb, sheet = 3, machine_output, startCol = "A", startRow = 1)
  
  #write the workbook
  output_filepath <- readline(prompt = "Input the output filepath and filename, e.g. output/output_file: ") #user states where the file will be output
  saveWorkbook(wb, paste0(output_filepath, ".xlsx"), overwrite = TRUE) #writes the output and adds .xlsx to the stated filename
  message("Output exported to ", output_filepath, ".xlsx")
}

# ʕ·ᴥ·ʔ
# :)