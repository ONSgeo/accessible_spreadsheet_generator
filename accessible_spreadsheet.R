library(tidyverse)
library(readxl)
library(openxlsx)

####Load in input data####
raw_data <- read_csv("data/Happiness.csv")

#using head of data for testing, because of special o character Ynys Mon
raw_head <- head(raw_data)

####Preparing data layout####


#create a new col in raw_data with the first 3 char of the entity code in it
#do unique on that col to get a list of all the entity codes that are used in our data
#then use an if statement to load in the appropriate lookup based on what hierarchy the entity code falls in
#join the raw_data to the appropriate lookup
#QA checks to make sure everything has joined correctly
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
