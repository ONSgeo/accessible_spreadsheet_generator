####ONS Geography Accessible Spreadsheet Generator####
#This generator takes input data from the user and exports in an accessible Excel format suitable for screen readers 
#Exported data is presented in the correct geographical hierarchy display order
#The ASG accepts input data in .csv or Excel formats
#This script also requires the accompanying accessible_spreadsheet_functions.R script in order to run custom functions
#Developed by Helen Calder and Heather Porter, March 2023 
#ONS Geospatial - Contact: geospatial@ons.gov.uk

##Load library packages
library(tidyverse)
library(readxl)
library(openxlsx)
library(janitor)



##Load the custom functions
source("accessible_spreadsheet_functions.R")



####Load and prepare data####
#load the user input data - file format must be a .csv or .xlsx / .xls
#requires a user interaction in the console to input the location filepath and select the geography code column
raw_data <- load_user_data()

#Delete before publishing
#data/Happiness.csv
#data/wellbeing_testdata_2021.csv

#creates a vector of all unique entity codes in the input data
#this is used later in the lookup selection process
unique_entities <- create_unique_entities(raw_data)

#this loads the appropriate lookup for the input data
#uses the unique_entities and compares them against known entities in each separate geographic hierarchy
#requires a user interaction in the console to select the correct column for the geography code
#if multiple geographies are found in the input data, a further user interaction may be required.
lookup <- lookup_loader(raw_data, unique_entities)

#this is a helper function for later use - gets the column names of specific columns
#requires a user interaction in the console to select the correct column for the geography code in the user input data AND the lookup
define_col_names_result <- define_col_names(lookup, raw_data)

#unlists the column names
input_data_col_name <- define_col_names_result[1] %>% unlist()
lookup_col_name <- define_col_names_result[2] %>% unlist()

#joins the input user data to the appropriate lookup
raw_data_join <- data_joiner(raw_data, lookup, input_data_col_name, lookup_col_name) 



####QA Checks####
#this unmatched lookup test checks what rows in the lookup didn't have data joined to them
unmatched_lookup_test <- unmatched_lookup_check(lookup, raw_data, input_data_col_name, lookup_col_name)
#this unmatched data test checks what rows in the data didn't join to the lookup i.e. the geography code associated with the data was not found in the lookup
unmatched_data_test <- unmatched_data_check(lookup, raw_data, input_data_col_name, lookup_col_name)

#export the results of the QA checks (optional function)
export_test_results(unmatched_lookup_test, unmatched_data_test)



####Data output####
#formatting data in a format ready for output - prepares human and machine readable data outputs
#requires a user interaction in the console to select the correct column for the data variable for publication 
output <- output_preparation(raw_data_join, lookup, unmatched_lookup_test, input_data_col_name, lookup_col_name)

#unlisting the output results
human_output <- output[[1]]
machine_output <- output[[2]]

#formats and exports the Excel workbook
#requires a user interaction in the console to assign the Excel workbook Title metadata (In Excel: File > Info > Title)
#requires a user interaction in the console to input the output filepath and file name
export_workbook(human_output, machine_output)

