library(tidyverse)
library(readxl)
library(openxlsx)

####Load in input data####
raw_data <- read_csv("data/Happiness.csv")

#using head of data for testing, because of special o character Ynys Mon
raw_head <- head(raw_data)

####Preparing data layout####

#create a new col in raw_data with the first 3 char of the entity code in it
ent_data <- raw_data %>% mutate(ENTCD = substr(AREACD, 1, 3))

#do unique on that col to get a list of all the entity codes that are used in our data
distinct_entities <- ent_data %>% distinct(ENTCD, .keep_all = FALSE)
#or use this in base R to create a vector
unique_entities <- unique(ent_data$ENTCD)

#TODO check the unique_entities %in% is assessing all entities against the defined vector
#then use an if statement to load in the appropriate lookup based on what hierarchy the entity code falls in
if(unique_entities %in% c("K02","K03", "K04", "E92", "E06", "E07", "E08", "E09", "E10", "E11", "E12", "E13", "W92", "W06", "S92", "S12", "N92", "N09") == TRUE){
  lookup <- read_csv("lookups/CTRY22_NAT22_RGN22_CTYUA22_LAD22_lookup.csv") #admin hierarchy lookup
} else {
  message("Lookup not found")
}

input_data_col_name <- "AREACD"
#join the raw_data to the appropriate lookup
ent_data_join <- left_join(lookup, ent_data, by = c("AREA22CD" = input_data_col_name))
#QA checks to make sure everything has joined correctly - number of rows, what entities have NA values and which cols
na_data <- filter(ent_data_join, is.na(Value) == TRUE)
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
