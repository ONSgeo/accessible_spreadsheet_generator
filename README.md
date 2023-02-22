# Accessible Spreadsheet Generator
A tool to generate accessible spreadsheets with the correct hierarchical ordering of geographies.

## Introduction
This tool takes a csv or xlsx input with a column for GSS geography code (mixed geographies, possibly in the wrong display order) and additional column(s) with associated statistics in. 

The output file is an Excel spreadsheet formatted in the GSS accessible format with the following sheets:
* introduction sheet explaining what the dataset is
* human-friendly data, arranged suitably so a screen reader can properly read the data out to the user
* machine readable data in a *tidy* format.

## Available geographic hierarchies
The following geograhpic hierarchies are available in this tool at present:
* Country, nation, region, county/UA, LAD for 2022
