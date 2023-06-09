# Accessible Spreadsheet Generator
A tool to generate accessible spreadsheets with the correct hierarchical ordering of geographies.

## Introduction
This tool takes a csv or xlsx input with a column for GSS geography code (mixed geographies, possibly in the wrong display order) and additional column(s) with associated statistics in. It selects the appropriate type and year of GSS geography code, organises the data into the correct display order for the particularly geogrpahical hierarchy and outputs an accessible excel file of the results. 

The output file is an Excel spreadsheet formatted in the GSS accessible format with the following sheets:
* introduction sheet explaining what the dataset is
* human-friendly data, arranged suitably so a screen reader can properly read the data out to the user
* machine readable data in a *tidy* format.

## Available geographic hierarchies
The following geograhpic hierarchies are available in this tool at present:
* Country, nation, region, county/UA, LAD for 2022, 2021, 2020, 2019, 2018
* Country, nation, NHS England region, STP, CCG, LAD for 2021
* Country, nation, ITL1, ITL2, ITL3 for 2021

## How to use the tool
1) Locate an R user 
2) Ensure the R user has a functional install of R and R studio
3) Pull the repo to download a local copy of the project - How-to guide can be found here: https://docs.github.com/en/repositories/creating-and-managing-repositories/cloning-a-repository
4) Open the project in R studio and open accessible_spreadsheet.R
5) Load the libraries and custom functions using the code at the top of the script.
6) To process your data run the functions in the *Load and prepare data* section one by one, noting user required inputs. Prompts will appear in the console pane.
7) Run the functions in the *QA Checks* section to ensure all your data has been accounted for. Users should note that **the outputs of these checks should be manually inspected and fully understood before undertaking the rest of the process**. 
8) Finally, continue with the *Data Output* section functions to finalise and produce the Excel output. 

## Input data with multiple types of geography
If the user's input data has multiple geographical heirarchies within one file, the user will find that only one hierachy will be dealt with at a time. The inactive hierarchy data will be output in the QA checks file. If you wish to format this data properly for output, export the QA checks file using the optional function in the script and run the file throught he process again to produce the desired output.
