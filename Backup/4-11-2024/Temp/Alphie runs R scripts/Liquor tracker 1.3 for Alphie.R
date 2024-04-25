# ---
# title: "TRD's Liquor tracker | Version: 1.3"
# author: 
#   - Asif Rasool^[Senior Economist, 
#     Tax Analysis, Research & Statistics, 
#     Office of the Secretary, 
#     New Mexico Taxation & Revenue Department, 
#     last updated Apr 1st, 2024,
#     asif.rasool@tax.nm.gov, 
#      ]
# 
# date: October 16, 2023
# 
# output: 
#   html_document:
#     toc: true
#     toc_float: true
#     theme: cerulean
# ---
# 
# knitr::include_graphics("//trdecomsrv/H/Liquor Excise Tax/Data/Liquor Tax Tracker/input/trdlogo.jpg")


# How to run this script: Ctrl + Alt + R

# To generate a report, please click on the Knit button above

gc()


suppressPackageStartupMessages(library(knitr)) 
suppressPackageStartupMessages(library(dplyr))
suppressPackageStartupMessages(library(openxlsx))
suppressPackageStartupMessages(library(writexl))
suppressPackageStartupMessages(library(readxl))

suppressPackageStartupMessages(library(DBI))
suppressPackageStartupMessages(library(odbc))
suppressPackageStartupMessages(library(openxlsx))

suppressPackageStartupMessages(library(lubridate))



# Running SQL query

# Set up the connection parameters
server <- 'GenSQLreport,2187'
database <- 'master'
driver <- 'ODBC Driver 17 for SQL Server'  # Update the driver if necessary

# Set up the connection string with Windows authentication
conn_str <- paste0('Driver=', driver, ';Server=', server, ';Database=', database, ';Trusted_Connection=yes')

# Create a connection
conn <- dbConnect(odbc::odbc(), .connection_string = conn_str)

# Path to the .sql file
sql_file <- "//trdecomsrv/H/Liquor Excise Tax/Data/Liquor Tax Tracker/input/SQL Query_LIquor Returns with select variables.sql"

# Read the .sql file
sql_script <- readChar(sql_file, file.info(sql_file)$size)


# Execute the SQL script
dbExecute(conn, sql_script)

# Fetch the result set from tblNM_ReturnTpt joining tblReturn on flngDocKey
result_query <- "SELECT DISTINCT t1.*, t2.fdtmFilingperiod
                 FROM tblNM_Returnliq AS t1
                 INNER JOIN tblReturn AS t2 ON t1.flngDocKey = t2.flngDocKey"



result <- dbGetQuery(conn, result_query)

# Save the result set as an .xlsx file
xlsx_file <- "//trdecomsrv/H/Liquor Excise Tax/Data/Liquor Tax Tracker/input/Liquor Excise Tax Raw Return Data v.1.3.xlsx"
write.xlsx(result, xlsx_file, rowNames = FALSE)

# Close the connection
dbDisconnect(conn)


# Loadng raw data (SQL output)

df <- suppressWarnings({read_xlsx("//trdecomsrv/H/Liquor Excise Tax/Data/Liquor Tax Tracker/input/Liquor Excise Tax Raw Return Data v.1.3.xlsx", sheet = "Sheet 1")})

previous_month_first_day <- floor_date(Sys.Date(), "month") - days(1)
df <- df %>% filter(fdtmFilingperiod > as.Date('2004-06-01'))
df <- df %>% filter(fdtmFilingperiod < previous_month_first_day)
# df <- subset(df, !(fdtmFilingperiod == as.Date("2021-12-21")))




# Prepartion of the excel workbook

## Defining the formats

OUT <- createWorkbook()
LabelStyle <- createStyle(halign = "center",
                          border = c("bottom", "right"), 
                          borderStyle = "thin", 
                          textDecoration = "bold", 
                          fgFill = "#0491A1", 
                          fontColour = "white")
NumStyle <- createStyle(halign = "right", numFmt = "0.00")
TextStyle <- createStyle(halign = "center", 
                         border = "bottom", 
                         borderStyle = "thin")

DateStyle <- createStyle(halign = "center", numFmt = "mm-dd-yyyy")


## Creating the worksheets

        addWorksheet(OUT, "ReadMe")
        addWorksheet(OUT, "BeerPerGallon")
        addWorksheet(OUT, "MicrobeerPerGallon2a")
        addWorksheet(OUT, "MicrobeerPerGallon2b")
        addWorksheet(OUT, "MicrobeerPerGallon2c")
        addWorksheet(OUT, "CiderPerGallon")
        addWorksheet(OUT, "CiderSmallPerGallon4a")
        addWorksheet(OUT, "CiderSmallPerGallon4b")
        addWorksheet(OUT, "CiderSmallPerGallon4c")
        addWorksheet(OUT, "SpiritPerLiters")
        addWorksheet(OUT, "SpiritCraftPerLiters6a")
        addWorksheet(OUT, "SpiritCraftPerLiters6b")
        addWorksheet(OUT, "SpiritCraftPerLiters6c")
        addWorksheet(OUT, "SpiritCraftPerLiters6d")
        addWorksheet(OUT, "WinePerLiters")
        addWorksheet(OUT, "FWinePerLiters")
        addWorksheet(OUT, "SWinePerLiters9a")
        addWorksheet(OUT, "SWinePerLiters9b")
        addWorksheet(OUT, "SWinePerLiters9c")


# Processing Read Me sheet


insertImage(OUT, "ReadMe", "//trdecomsrv/H/Liquor Excise Tax/Data/Liquor Tax Tracker/input/trdlogo.jpg" , startRow = 2, startCol = 2, width = 2, height = 1)
freezePane(wb = OUT, sheet = "ReadMe" , firstActiveRow = 35, firstCol = FALSE)

insertImage(OUT, "ReadMe", "//trdecomsrv/H/Liquor Excise Tax/Data/Liquor Tax Tracker/input/readmetext.png" , startRow = 7, startCol = 2, width = 4.141, height = 6)

date <- Sys.Date()
date <- format(as.POSIXct(date,format='%m/%d/%Y %H:%M:%S'),format='%m/%d/%Y')

setColWidths(OUT, sheet = "ReadMe", cols = 7, widths =10)
writeData(wb=OUT, sheet = "ReadMe", x = date, startCol = 7, startRow = 34)
addStyle(OUT, sheet = "ReadMe", style = LabelStyle, rows = 34, cols = 7, 
         gridExpand = TRUE, stack = FALSE)



# Processing Beer per gallon

## Creating a subset from the raw SQL file

BeerPerGallon <- as.data.frame(df %>% select("fdtmFilingperiod", "fcurBeerSold", 
                                 "fcurBeerExemptions","fcurBeerTaxableGals", 
                                 "fcurBeerTax"))


## Calculating the summations

BeerPerGallon <- as.data.frame(aggregate(cbind(fcurBeerSold, fcurBeerExemptions, fcurBeerTaxableGals, fcurBeerTax)~ fdtmFilingperiod, data=BeerPerGallon, FUN=sum));


## Renaming the columns

BeerPerGallon <- BeerPerGallon %>% rename("Filing Period" = fdtmFilingperiod,
                                              "Beer Total Gallons sold"= fcurBeerSold,
                                              "Beer Deductions / Exemptions Gallons" = fcurBeerExemptions,
                                              "Beer Taxable Gallons sold" = fcurBeerTaxableGals,
                                              "Beer Tax ($)" = fcurBeerTax) 


## Formating and exporting to an Excel sheet


setColWidths(OUT, sheet = "BeerPerGallon", cols = 1:20, widths ="auto")
freezePane(wb = OUT, sheet = "BeerPerGallon" , firstRow = TRUE, firstCol = FALSE)
writeData(wb = OUT, sheet = "BeerPerGallon",x = BeerPerGallon, 
          startCol = 1, startRow = 1, colNames = TRUE)
addStyle(OUT, sheet = "BeerPerGallon", style = LabelStyle, rows = 1, cols = 1:5, 
         gridExpand = TRUE, stack = FALSE)
addStyle(OUT, sheet = "BeerPerGallon", style = DateStyle, cols = 1, rows = 2:1000,
         gridExpand = FALSE, stack = FALSE)
addStyle(OUT, sheet = "BeerPerGallon", style = NumStyle, cols = 2, rows = 2:1000, 
         gridExpand = FALSE, stack = FALSE)
addStyle(OUT, sheet = "BeerPerGallon", style = NumStyle, cols = 3, rows = 2:1000, 
         gridExpand = FALSE, stack = FALSE)
addStyle(OUT, sheet = "BeerPerGallon", style = NumStyle, cols = 4, rows = 2:1000, 
         gridExpand = FALSE, stack = FALSE)
addStyle(OUT, sheet = "BeerPerGallon", style = NumStyle, cols = 5, rows = 2:1000, 
         gridExpand = FALSE, stack = FALSE)


# Processing Microbeer per gallon

## Processing: 2a. First 30,000 barrels sold

### Creating a subset from the raw SQL file

MicrobeerPerGallon2a <- as.data.frame(df %>% select("fdtmFilingperiod", "fcurMBeerSoldA", 
                                                    "fcurMBeerExemptionsA","fcurMBeerTaxableGalsA", 
                                                    "fcurMBeerTaxA"))


### Calculating the summations

MicrobeerPerGallon2a <- as.data.frame(aggregate(cbind(fcurMBeerSoldA, fcurMBeerExemptionsA, 
                                               fcurMBeerTaxableGalsA, 
                                               fcurMBeerTaxA)~ fdtmFilingperiod, 
                                         data=MicrobeerPerGallon2a, FUN=sum))


### Renaming the columns

MicrobeerPerGallon2a <- MicrobeerPerGallon2a %>% rename("Filing Period" = fdtmFilingperiod,
                                          "Microbeer (First 30,000 barrels) Total Gallons sold"= fcurMBeerSoldA,
                                          "Deductions / Exemptions Gallons" = fcurMBeerExemptionsA,
                                          "Taxable Gallons sold" = fcurMBeerTaxableGalsA,
                                          "Microbeer (First 30,000 barrels) Tax ($)" = fcurMBeerTaxA) 


### Formating and exporting to an Excel sheet

setColWidths(OUT, sheet = "MicrobeerPerGallon2a", cols = 1:20, widths ="auto")
freezePane(wb = OUT, sheet = "MicrobeerPerGallon2a" , firstRow = TRUE, firstCol = FALSE)
writeData(wb = OUT, sheet = "MicrobeerPerGallon2a",x = MicrobeerPerGallon2a, 
          startCol = 1, startRow = 1, colNames = TRUE)
addStyle(OUT, sheet = "MicrobeerPerGallon2a", style = LabelStyle, rows = 1, cols = 1:5, 
         gridExpand = TRUE, stack = FALSE)
addStyle(OUT, sheet = "MicrobeerPerGallon2a", style = DateStyle, cols = 1, rows = 2:1000,
         gridExpand = FALSE, stack = FALSE)
addStyle(OUT, sheet = "MicrobeerPerGallon2a", style = NumStyle, cols = 2, rows = 2:1000, 
         gridExpand = FALSE, stack = FALSE)
addStyle(OUT, sheet = "MicrobeerPerGallon2a", style = NumStyle, cols = 3, rows = 2:1000, 
         gridExpand = FALSE, stack = FALSE)
addStyle(OUT, sheet = "MicrobeerPerGallon2a", style = NumStyle, cols = 4, rows = 2:1000, 
         gridExpand = FALSE, stack = FALSE)
addStyle(OUT, sheet = "MicrobeerPerGallon2a", style = NumStyle, cols = 5, rows = 2:1000, 
         gridExpand = FALSE, stack = FALSE)


## Processing: 2b. Sales: 30,001 to 59,999 barrels

### Creating a subset from the raw SQL file

MicrobeerPerGallon2b <- as.data.frame(df %>% select("fdtmFilingperiod", "fcurMBeerSoldB", 
                                                    "fcurMBeerExemptionsB","fcurMBeerTaxableGalsB", 
                                                    "fcurMBeerTaxB"))


### Calculating the summations

MicrobeerPerGallon2b <- as.data.frame(aggregate(cbind(fcurMBeerSoldB, fcurMBeerExemptionsB, 
                                                      fcurMBeerTaxableGalsB, 
                                                      fcurMBeerTaxB)~ fdtmFilingperiod, 
                                                data=MicrobeerPerGallon2b, FUN=sum))


### Renaming the columns

MicrobeerPerGallon2b <- MicrobeerPerGallon2b %>% rename("Filing Period" = fdtmFilingperiod,
                                                "Microbeer (Sales: 30,001 to 59,999 barrels) Total Gallons sold"= fcurMBeerSoldB,
                                                "Deductions / Exemptions Gallons" = fcurMBeerExemptionsB,
                                                "Taxable Gallons sold" = fcurMBeerTaxableGalsB,
                                                "Microbeer (Sales: 30,001 to 59,999 barrels) Tax ($)" = fcurMBeerTaxB) 


### Formating and exporting to an Excel sheet


setColWidths(OUT, sheet = "MicrobeerPerGallon2b", cols = 1:20, widths ="auto")
freezePane(wb = OUT, sheet = "MicrobeerPerGallon2b" , firstRow = TRUE, firstCol = FALSE)
writeData(wb = OUT, sheet = "MicrobeerPerGallon2b",x = MicrobeerPerGallon2b, 
          startCol = 1, startRow = 1, colNames = TRUE)
addStyle(OUT, sheet = "MicrobeerPerGallon2b", style = LabelStyle, rows = 1, cols = 1:5, 
         gridExpand = TRUE, stack = FALSE)
addStyle(OUT, sheet = "MicrobeerPerGallon2b", style = DateStyle, cols = 1, rows = 2:1000,
         gridExpand = FALSE, stack = FALSE)
addStyle(OUT, sheet = "MicrobeerPerGallon2b", style = NumStyle, cols = 2, rows = 2:1000, 
         gridExpand = FALSE, stack = FALSE)
addStyle(OUT, sheet = "MicrobeerPerGallon2b", style = NumStyle, cols = 3, rows = 2:1000, 
         gridExpand = FALSE, stack = FALSE)
addStyle(OUT, sheet = "MicrobeerPerGallon2b", style = NumStyle, cols = 4, rows = 2:1000, 
         gridExpand = FALSE, stack = FALSE)
addStyle(OUT, sheet = "MicrobeerPerGallon2b", style = NumStyle, cols = 5, rows = 2:1000, 
         gridExpand = FALSE, stack = FALSE)


## Processing: 2c. Sales: 60,000 or more barrels

### Creating a subset from the raw SQL file

MicrobeerPerGallon2c <- as.data.frame(df %>% select("fdtmFilingperiod", "fcurMBeerSoldC", 
                                                    "fcurMBeerExemptionsC","fcurMBeerTaxableGalsC", 
                                                    "fcurMBeerTaxC"))


### Calculating the summations

MicrobeerPerGallon2c <- as.data.frame(aggregate(cbind(fcurMBeerSoldC, fcurMBeerExemptionsC, 
                                                      fcurMBeerTaxableGalsC, 
                                                      fcurMBeerTaxC)~ fdtmFilingperiod, 
                                                data=MicrobeerPerGallon2c, FUN=sum))


### Renaming the columns

MicrobeerPerGallon2c <- MicrobeerPerGallon2c %>% rename("Filing Period" = fdtmFilingperiod,
                              "Microbeer (60,000 or more barrels) Total Gallons sold"= fcurMBeerSoldC,
                              "Deductions / Exemptions Gallons" = fcurMBeerExemptionsC,
                              "Taxable Gallons sold" = fcurMBeerTaxableGalsC,
                              "Microbeer (60,000 or more barrels) Tax ($)" = fcurMBeerTaxC)


### Formating and exporting to an Excel sheet


setColWidths(OUT, sheet = "MicrobeerPerGallon2c", cols = 1:20, widths ="auto")

freezePane(wb = OUT, sheet = "MicrobeerPerGallon2c" , firstRow = TRUE, firstCol = FALSE)
writeData(wb = OUT, sheet = "MicrobeerPerGallon2c",x = MicrobeerPerGallon2c, 
          startCol = 1, startRow = 1, colNames = TRUE)
addStyle(OUT, sheet = "MicrobeerPerGallon2c", style = LabelStyle, rows = 1, cols = 1:5, 
         gridExpand = TRUE, stack = FALSE)
addStyle(OUT, sheet = "MicrobeerPerGallon2c", style = DateStyle, cols = 1, rows = 2:1000,
         gridExpand = FALSE, stack = FALSE)
addStyle(OUT, sheet = "MicrobeerPerGallon2c", style = NumStyle, cols = 2, rows = 2:1000, 
         gridExpand = FALSE, stack = FALSE)
addStyle(OUT, sheet = "MicrobeerPerGallon2c", style = NumStyle, cols = 3, rows = 2:1000, 
         gridExpand = FALSE, stack = FALSE)
addStyle(OUT, sheet = "MicrobeerPerGallon2c", style = NumStyle, cols = 4, rows = 2:1000, 
         gridExpand = FALSE, stack = FALSE)
addStyle(OUT, sheet = "MicrobeerPerGallon2c", style = NumStyle, cols = 5, rows = 2:1000, 
         gridExpand = FALSE, stack = FALSE)


# Processing: Cider per gallon

## Creating a subset from the raw SQL file

CiderPerGallon <- as.data.frame(df %>% select("fdtmFilingperiod", "fcurCiderSold", 
                                 "fcurCiderExemptions","fcurCiderTaxableGals", 
                                 "fcurCiderTax"))


## Calculating the summations

CiderPerGallon <- as.data.frame(aggregate(cbind(fcurCiderSold, fcurCiderExemptions, fcurCiderTaxableGals, fcurCiderTax)~ fdtmFilingperiod, data=CiderPerGallon, FUN=sum));


## Renaming the columns

CiderPerGallon <- CiderPerGallon %>% rename("Filing Period" = fdtmFilingperiod,
                                              "Cider Total Gallons sold"= fcurCiderSold,
                                              "Cider Deductions / Exemptions Gallons" = fcurCiderExemptions,
                                              "Cider Taxable Gallons sold" = fcurCiderTaxableGals,
                                              "Cider Tax ($)" = fcurCiderTax) 


## Formating and exporting to an Excel sheet


setColWidths(OUT, sheet = "CiderPerGallon", cols = 1:20, widths ="auto")
freezePane(wb = OUT, sheet = "CiderPerGallon" , firstRow = TRUE, firstCol = FALSE)
writeData(wb = OUT, sheet = "CiderPerGallon",x = CiderPerGallon, 
          startCol = 1, startRow = 1, colNames = TRUE)
addStyle(OUT, sheet = "CiderPerGallon", style = LabelStyle, rows = 1, cols = 1:5, 
         gridExpand = TRUE, stack = FALSE)
addStyle(OUT, sheet = "CiderPerGallon", style = DateStyle, cols = 1, rows = 2:1000,
         gridExpand = FALSE, stack = FALSE)
addStyle(OUT, sheet = "CiderPerGallon", style = NumStyle, cols = 2, rows = 2:1000, 
         gridExpand = FALSE, stack = FALSE)
addStyle(OUT, sheet = "CiderPerGallon", style = NumStyle, cols = 3, rows = 2:1000, 
         gridExpand = FALSE, stack = FALSE)
addStyle(OUT, sheet = "CiderPerGallon", style = NumStyle, cols = 4, rows = 2:1000, 
         gridExpand = FALSE, stack = FALSE)
addStyle(OUT, sheet = "CiderPerGallon", style = NumStyle, cols = 5, rows = 2:1000, 
         gridExpand = FALSE, stack = FALSE)


# Processing Cider per gallon: Cider manufactured/produced by a small winegrower and sold in New Mexico.

## Processing: 4a. First 30,000 barrels sold

### Creating a subset from the raw SQL file

CiderSmallPerGallon4a <- as.data.frame(df %>% select("fdtmFilingperiod", "fcurCiderSmallSoldA", 
                                                    "fcurCiderSmallExemptionsA","fcurCiderSmallTaxableGalsA", 
                                                    "fcurCiderSmallTaxA"))


### Calculating the summations

CiderSmallPerGallon4a <- as.data.frame(aggregate(cbind(fcurCiderSmallSoldA, fcurCiderSmallExemptionsA, 
                                               fcurCiderSmallTaxableGalsA, 
                                               fcurCiderSmallTaxA)~ fdtmFilingperiod, 
                                         data=CiderSmallPerGallon4a, FUN=sum))


### Renaming the columns

CiderSmallPerGallon4a <- CiderSmallPerGallon4a %>% rename("Filing Period" = fdtmFilingperiod,
                                          "Cider (First 30,000 barrels) Total Gallons sold"= fcurCiderSmallSoldA,
                                          "Deductions / Exemptions Gallons" = fcurCiderSmallExemptionsA,
                                          "Taxable Gallons sold" = fcurCiderSmallTaxableGalsA,
                                          "Cider (First 30,000 barrels) Tax ($)" = fcurCiderSmallTaxA) 


### Formating and exporting to an Excel sheet


setColWidths(OUT, sheet = "CiderSmallPerGallon4a", cols = 1:20, widths ="auto")
freezePane(wb = OUT, sheet = "CiderSmallPerGallon4a" , firstRow = TRUE, firstCol = FALSE)
writeData(wb = OUT, sheet = "CiderSmallPerGallon4a",x = CiderSmallPerGallon4a, 
          startCol = 1, startRow = 1, colNames = TRUE)
addStyle(OUT, sheet = "CiderSmallPerGallon4a", style = LabelStyle, rows = 1, cols = 1:5, 
         gridExpand = TRUE, stack = FALSE)
addStyle(OUT, sheet = "CiderSmallPerGallon4a", style = DateStyle, cols = 1, rows = 2:1000,
         gridExpand = FALSE, stack = FALSE)
addStyle(OUT, sheet = "CiderSmallPerGallon4a", style = NumStyle, cols = 2, rows = 2:1000, 
         gridExpand = FALSE, stack = FALSE)
addStyle(OUT, sheet = "CiderSmallPerGallon4a", style = NumStyle, cols = 3, rows = 2:1000, 
         gridExpand = FALSE, stack = FALSE)
addStyle(OUT, sheet = "CiderSmallPerGallon4a", style = NumStyle, cols = 4, rows = 2:1000, 
         gridExpand = FALSE, stack = FALSE)
addStyle(OUT, sheet = "CiderSmallPerGallon4a", style = NumStyle, cols = 5, rows = 2:1000, 
         gridExpand = FALSE, stack = FALSE)


## Processing: 4b. Sales: 30,001 to 59,999 barrels

### Creating a subset from the raw SQL file

CiderSmallPerGallon4b <- as.data.frame(df %>% select("fdtmFilingperiod", "fcurCiderSmallSoldB", 
                                                    "fcurCiderSmallExemptionsB","fcurCiderSmallTaxableGalsB", 
                                                    "fcurCiderSmallTaxB"))


### Calculating the summations

CiderSmallPerGallon4b <- as.data.frame(aggregate(cbind(fcurCiderSmallSoldB, fcurCiderSmallExemptionsB, 
                                               fcurCiderSmallTaxableGalsB, 
                                               fcurCiderSmallTaxB)~ fdtmFilingperiod, 
                                         data=CiderSmallPerGallon4b, FUN=sum))


### Renaming the columns

CiderSmallPerGallon4b <- CiderSmallPerGallon4b %>% rename("Filing Period" = fdtmFilingperiod,
                                          "Cider (Sales: 30,001 to 59,999 barrels) Total Gallons sold"= fcurCiderSmallSoldB,
                                          "Deductions / Exemptions Gallons" = fcurCiderSmallExemptionsB,
                                          "Taxable Gallons sold" = fcurCiderSmallTaxableGalsB,
                                          "Cider (Sales: 30,001 to 59,999 barrels) Tax ($)" = fcurCiderSmallTaxB) 


### Formating and exporting to an Excel sheet


setColWidths(OUT, sheet = "CiderSmallPerGallon4b", cols = 1:20, widths ="auto")
freezePane(wb = OUT, sheet = "CiderSmallPerGallon4b" , firstRow = TRUE, firstCol = FALSE)


writeData(wb = OUT, sheet = "CiderSmallPerGallon4b",x = CiderSmallPerGallon4b, 
          startCol = 1, startRow = 1, colNames = TRUE)
addStyle(OUT, sheet = "CiderSmallPerGallon4b", style = LabelStyle, rows = 1, cols = 1:5, 
         gridExpand = TRUE, stack = FALSE)
addStyle(OUT, sheet = "CiderSmallPerGallon4b", style = DateStyle, cols = 1, rows = 2:1000,
         gridExpand = FALSE, stack = FALSE)
addStyle(OUT, sheet = "CiderSmallPerGallon4b", style = NumStyle, cols = 2, rows = 2:1000, 
         gridExpand = FALSE, stack = FALSE)
addStyle(OUT, sheet = "CiderSmallPerGallon4b", style = NumStyle, cols = 3, rows = 2:1000, 
         gridExpand = FALSE, stack = FALSE)
addStyle(OUT, sheet = "CiderSmallPerGallon4b", style = NumStyle, cols = 4, rows = 2:1000, 
         gridExpand = FALSE, stack = FALSE)
addStyle(OUT, sheet = "CiderSmallPerGallon4b", style = NumStyle, cols = 5, rows = 2:1000, 
         gridExpand = FALSE, stack = FALSE)


## Processing: 4c. Sales: 60,000 or more barrels

### Creating a subset from the raw SQL file

CiderSmallPerGallon4c <- as.data.frame(df %>% select("fdtmFilingperiod", "fcurCiderSmallSoldC", 
                                                    "fcurCiderSmallExemptionsC","fcurCiderSmallTaxableGalsC", 
                                                    "fcurCiderSmallTaxC"))


### Calculating the summations

CiderSmallPerGallon4c <- as.data.frame(aggregate(cbind(fcurCiderSmallSoldC, fcurCiderSmallExemptionsC, 
                                               fcurCiderSmallTaxableGalsC, 
                                               fcurCiderSmallTaxC)~ fdtmFilingperiod, 
                                         data=CiderSmallPerGallon4c, FUN=sum))


### Renaming the columns

CiderSmallPerGallon4c <- CiderSmallPerGallon4c %>% rename("Filing Period" = fdtmFilingperiod,
                                          "Cider (Sales: 60,000 or more barrels) Total Gallons sold"= fcurCiderSmallSoldC,
                                          "Deductions / Exemptions Gallons" = fcurCiderSmallExemptionsC,
                                          "Taxable Gallons sold" = fcurCiderSmallTaxableGalsC,
                                          "Cider (Sales: 60,000 or more barrels) Tax ($)" = fcurCiderSmallTaxC) 


### Formating and exporting to an Excel sheet


setColWidths(OUT, sheet = "CiderSmallPerGallon4c", cols = 1:20, widths ="auto")
freezePane(wb = OUT, sheet = "CiderSmallPerGallon4c" , firstRow = TRUE, firstCol = FALSE)

writeData(wb = OUT, sheet = "CiderSmallPerGallon4c",x = CiderSmallPerGallon4c, 
          startCol = 1, startRow = 1, colNames = TRUE)
addStyle(OUT, sheet = "CiderSmallPerGallon4c", style = LabelStyle, rows = 1, cols = 1:5, 
         gridExpand = TRUE, stack = FALSE)
addStyle(OUT, sheet = "CiderSmallPerGallon4c", style = DateStyle, cols = 1, rows = 2:1000,
         gridExpand = FALSE, stack = FALSE)
addStyle(OUT, sheet = "CiderSmallPerGallon4c", style = NumStyle, cols = 2, rows = 2:1000, 
         gridExpand = FALSE, stack = FALSE)
addStyle(OUT, sheet = "CiderSmallPerGallon4c", style = NumStyle, cols = 3, rows = 2:1000, 
         gridExpand = FALSE, stack = FALSE)
addStyle(OUT, sheet = "CiderSmallPerGallon4c", style = NumStyle, cols = 4, rows = 2:1000, 
         gridExpand = FALSE, stack = FALSE)
addStyle(OUT, sheet = "CiderSmallPerGallon4c", style = NumStyle, cols = 5, rows = 2:1000, 
         gridExpand = FALSE, stack = FALSE)


# Processing Spirituous liquor per liter

## Creating a subset from the raw SQL file

SpiritPerLiters <- as.data.frame(df %>% select("fdtmFilingperiod", "fcurSpiritSold", 
                                 "fcurSpiritExemptions","fcurSpiritTaxableLiters", 
                                 "fcurSpiritTax"))


## Calculating the summations

SpiritPerLiters <- as.data.frame(aggregate(cbind(fcurSpiritSold, fcurSpiritExemptions, fcurSpiritTaxableLiters, fcurSpiritTax)~ fdtmFilingperiod, data=SpiritPerLiters, FUN=sum));


## Renaming the columns

SpiritPerLiters <- SpiritPerLiters %>% rename("Filing Period" = fdtmFilingperiod,
                                              "Spirit Total Liters sold"= fcurSpiritSold,
                                              "Spirit Deductions / Exemptions Liters" = fcurSpiritExemptions,
                                              "Spirit Taxable Liters sold" = fcurSpiritTaxableLiters,
                                              "Spirit Tax ($)" = fcurSpiritTax) 


## Formating and exporting to an Excel sheet


setColWidths(OUT, sheet = "SpiritPerLiters", cols = 1:20, widths ="auto")
freezePane(wb = OUT, sheet = "SpiritPerLiters" , firstRow = TRUE, firstCol = FALSE)

writeData(wb = OUT, sheet = "SpiritPerLiters",x = SpiritPerLiters, 
          startCol = 1, startRow = 1, colNames = TRUE)
addStyle(OUT, sheet = "SpiritPerLiters", style = LabelStyle, rows = 1, cols = 1:5, 
         gridExpand = TRUE, stack = FALSE)
addStyle(OUT, sheet = "SpiritPerLiters", style = DateStyle, cols = 1, rows = 2:1000,
         gridExpand = FALSE, stack = FALSE)
addStyle(OUT, sheet = "SpiritPerLiters", style = NumStyle, cols = 2, rows = 2:1000, 
         gridExpand = FALSE, stack = FALSE)
addStyle(OUT, sheet = "SpiritPerLiters", style = NumStyle, cols = 3, rows = 2:1000, 
         gridExpand = FALSE, stack = FALSE)
addStyle(OUT, sheet = "SpiritPerLiters", style = NumStyle, cols = 4, rows = 2:1000, 
         gridExpand = FALSE, stack = FALSE)
addStyle(OUT, sheet = "SpiritPerLiters", style = NumStyle, cols = 5, rows = 2:1000, 
         gridExpand = FALSE, stack = FALSE)


# Spirituous liquor per liter: Spirituous liquor manufactured/produced by a craft distiller.

## Processing: 6a. First 250,000 liters sold (10% ABV or lower)

### Creating a subset from the raw SQL file

SpiritCraftPerLiters6a <- as.data.frame(df %>% select("fdtmFilingperiod", "fcurSpiritCraftSoldA", 
                                                    "fcurSpiritCraftExemptionsA","fcurSpiritCraftTaxableLitersA", 
                                                    "fcurSpiritCraftTaxA"))


### Calculating the summations

SpiritCraftPerLiters6a <- as.data.frame(aggregate(cbind(fcurSpiritCraftSoldA, fcurSpiritCraftExemptionsA, 
                                               fcurSpiritCraftTaxableLitersA, 
                                               fcurSpiritCraftTaxA)~ fdtmFilingperiod, 
                                         data=SpiritCraftPerLiters6a, FUN=sum))


### Renaming the columns

SpiritCraftPerLiters6a <- SpiritCraftPerLiters6a %>% rename("Filing Period" = fdtmFilingperiod,
                                          "SpiritCraft (First 250,000 liters sold) Total Liters sold (10% ABV or lower)"= fcurSpiritCraftSoldA,
                                          "Deductions / Exemptions Liters" = fcurSpiritCraftExemptionsA,
                                          "Taxable Liters sold" = fcurSpiritCraftTaxableLitersA,
                                          "SpiritCraft (First 250,000 liters sold) Tax ($)" = fcurSpiritCraftTaxA) 


### Formating and exporting to an Excel sheet


setColWidths(OUT, sheet = "SpiritCraftPerLiters6a", cols = 1:20, widths ="auto")
freezePane(wb = OUT, sheet = "SpiritCraftPerLiters6a" , firstRow = TRUE, firstCol = FALSE)
writeData(wb = OUT, sheet = "SpiritCraftPerLiters6a",x = SpiritCraftPerLiters6a, 
          startCol = 1, startRow = 1, colNames = TRUE)
addStyle(OUT, sheet = "SpiritCraftPerLiters6a", style = LabelStyle, rows = 1, cols = 1:5, 
         gridExpand = TRUE, stack = FALSE)
addStyle(OUT, sheet = "SpiritCraftPerLiters6a", style = DateStyle, cols = 1, rows = 2:1000,
         gridExpand = FALSE, stack = FALSE)
addStyle(OUT, sheet = "SpiritCraftPerLiters6a", style = NumStyle, cols = 2, rows = 2:1000, 
         gridExpand = FALSE, stack = FALSE)
addStyle(OUT, sheet = "SpiritCraftPerLiters6a", style = NumStyle, cols = 3, rows = 2:1000, 
         gridExpand = FALSE, stack = FALSE)
addStyle(OUT, sheet = "SpiritCraftPerLiters6a", style = NumStyle, cols = 4, rows = 2:1000, 
         gridExpand = FALSE, stack = FALSE)
addStyle(OUT, sheet = "SpiritCraftPerLiters6a", style = NumStyle, cols = 5, rows = 2:1000, 
         gridExpand = FALSE, stack = FALSE)


## Processing: 6b. Next 250,000 liters (10% ABV or lower)

### Creating a subset from the raw SQL file

SpiritCraftPerLiters6b <- as.data.frame(df %>% select("fdtmFilingperiod", "fcurSpiritCraftSoldB", 
                                                    "fcurSpiritCraftExemptionsB","fcurSpiritCraftTaxableLitersB", 
                                                    "fcurSpiritCraftTaxB"))


### Calculating the summations

SpiritCraftPerLiters6b <- as.data.frame(aggregate(cbind(fcurSpiritCraftSoldB, fcurSpiritCraftExemptionsB, 
                                               fcurSpiritCraftTaxableLitersB, 
                                               fcurSpiritCraftTaxB)~ fdtmFilingperiod, 
                                         data=SpiritCraftPerLiters6b, FUN=sum))


### Renaming the columns

SpiritCraftPerLiters6b <- SpiritCraftPerLiters6b %>% rename("Filing Period" = fdtmFilingperiod,
                                          "SpiritCraft (Next 250,000 liters sold) Total Liters sold (10% ABV or lower)"= fcurSpiritCraftSoldB,
                                          "Deductions / Exemptions Liters" = fcurSpiritCraftExemptionsB,
                                          "Taxable Liters sold" = fcurSpiritCraftTaxableLitersB,
                                          "SpiritCraft (Next 250,000 liters sold) Tax ($)" = fcurSpiritCraftTaxB) 


### Formating and exporting to an Excel sheet


setColWidths(OUT, sheet = "SpiritCraftPerLiters6b", cols = 1:20, widths ="auto")
freezePane(wb = OUT, sheet = "SpiritCraftPerLiters6b" , firstRow = TRUE, firstCol = FALSE)

writeData(wb = OUT, sheet = "SpiritCraftPerLiters6b",x = SpiritCraftPerLiters6b, 
          startCol = 1, startRow = 1, colNames = TRUE)
addStyle(OUT, sheet = "SpiritCraftPerLiters6b", style = LabelStyle, rows = 1, cols = 1:5, 
         gridExpand = TRUE, stack = FALSE)
addStyle(OUT, sheet = "SpiritCraftPerLiters6b", style = DateStyle, cols = 1, rows = 2:1000,
         gridExpand = FALSE, stack = FALSE)
addStyle(OUT, sheet = "SpiritCraftPerLiters6b", style = NumStyle, cols = 2, rows = 2:1000, 
         gridExpand = FALSE, stack = FALSE)
addStyle(OUT, sheet = "SpiritCraftPerLiters6b", style = NumStyle, cols = 3, rows = 2:1000, 
         gridExpand = FALSE, stack = FALSE)
addStyle(OUT, sheet = "SpiritCraftPerLiters6b", style = NumStyle, cols = 4, rows = 2:1000, 
         gridExpand = FALSE, stack = FALSE)
addStyle(OUT, sheet = "SpiritCraftPerLiters6b", style = NumStyle, cols = 5, rows = 2:1000, 
         gridExpand = FALSE, stack = FALSE)


## Processing: 6c. First 175,000 liters sold (over 10% ABV)

### Creating a subset from the raw SQL file

SpiritCraftPerLiters6c <- as.data.frame(df %>% select("fdtmFilingperiod", "fcurSpiritCraftSoldC", 
                                                    "fcurSpiritCraftExemptionsC","fcurSpiritCraftTaxableLitersC", 
                                                    "fcurSpiritCraftTaxC"))


### Calculating the summations

SpiritCraftPerLiters6c <- as.data.frame(aggregate(cbind(fcurSpiritCraftSoldC, fcurSpiritCraftExemptionsC, 
                                               fcurSpiritCraftTaxableLitersC, 
                                               fcurSpiritCraftTaxC)~ fdtmFilingperiod, 
                                         data=SpiritCraftPerLiters6c, FUN=sum))


### Renaming the columns

SpiritCraftPerLiters6c <- SpiritCraftPerLiters6c %>% rename("Filing Period" = fdtmFilingperiod,
                                          "SpiritCraft (First 175,000 liters sold (over 10% ABV)) Total Liters sold"= fcurSpiritCraftSoldC,
                                          "Deductions / Exemptions Liters" = fcurSpiritCraftExemptionsC,
                                          "Taxable Liters sold" = fcurSpiritCraftTaxableLitersC,
                                          "SpiritCraft (First 175,000 liters sold (over 10% ABV)) Tax ($)" = fcurSpiritCraftTaxC) 


### Formating and exporting to an Excel sheet



setColWidths(OUT, sheet = "SpiritCraftPerLiters6c", cols = 1:20, widths ="auto")
freezePane(wb = OUT, sheet = "SpiritCraftPerLiters6c" , firstRow = TRUE, firstCol = FALSE)
writeData(wb = OUT, sheet = "SpiritCraftPerLiters6c",x = SpiritCraftPerLiters6c, 
          startCol = 1, startRow = 1, colNames = TRUE)
addStyle(OUT, sheet = "SpiritCraftPerLiters6c", style = LabelStyle, rows = 1, cols = 1:5, 
         gridExpand = TRUE, stack = FALSE)
addStyle(OUT, sheet = "SpiritCraftPerLiters6c", style = DateStyle, cols = 1, rows = 2:1000,
         gridExpand = FALSE, stack = FALSE)
addStyle(OUT, sheet = "SpiritCraftPerLiters6c", style = NumStyle, cols = 2, rows = 2:1000, 
         gridExpand = FALSE, stack = FALSE)
addStyle(OUT, sheet = "SpiritCraftPerLiters6c", style = NumStyle, cols = 3, rows = 2:1000, 
         gridExpand = FALSE, stack = FALSE)
addStyle(OUT, sheet = "SpiritCraftPerLiters6c", style = NumStyle, cols = 4, rows = 2:1000, 
         gridExpand = FALSE, stack = FALSE)
addStyle(OUT, sheet = "SpiritCraftPerLiters6c", style = NumStyle, cols = 5, rows = 2:1000, 
         gridExpand = FALSE, stack = FALSE)


## 6d. Next 200,000 liters (over 10% ABV)

### Creating a subset from the raw SQL file

SpiritCraftPerLiters6d <- as.data.frame(df %>% select("fdtmFilingperiod", "fcurSpiritCraftSoldD", 
                                                    "fcurSpiritCraftExemptionsD","fcurSpiritCraftTaxableLitersD", 
                                                    "fcurSpiritCraftTaxD"))


### Calculating the summations

SpiritCraftPerLiters6d <- as.data.frame(aggregate(cbind(fcurSpiritCraftSoldD, fcurSpiritCraftExemptionsD, 
                                               fcurSpiritCraftTaxableLitersD, 
                                               fcurSpiritCraftTaxD)~ fdtmFilingperiod, 
                                         data=SpiritCraftPerLiters6d, FUN=sum))


### Renaming the columns

SpiritCraftPerLiters6d <- SpiritCraftPerLiters6d %>% rename("Filing Period" = fdtmFilingperiod,
                                          "SpiritCraft (Next 200,000 liters (over 10% ABV)) Total Liters sold"= fcurSpiritCraftSoldD,
                                          "Deductions / Exemptions Liters" = fcurSpiritCraftExemptionsD,
                                          "Taxable Liters sold" = fcurSpiritCraftTaxableLitersD,
                                          "SpiritCraft (Next 200,000 liters (over 10% ABV)) Tax ($)" = fcurSpiritCraftTaxD) 


### Formating and exporting to an Excel sheet


setColWidths(OUT, sheet = "SpiritCraftPerLiters6d", cols = 1:20, widths ="auto")
freezePane(wb = OUT, sheet = "SpiritCraftPerLiters6d" , firstRow = TRUE, firstCol = FALSE)

writeData(wb = OUT, sheet = "SpiritCraftPerLiters6d",x = SpiritCraftPerLiters6d, 
          startCol = 1, startRow = 1, colNames = TRUE)
addStyle(OUT, sheet = "SpiritCraftPerLiters6d", style = LabelStyle, rows = 1, cols = 1:5, 
         gridExpand = TRUE, stack = FALSE)
addStyle(OUT, sheet = "SpiritCraftPerLiters6d", style = DateStyle, cols = 1, rows = 2:1000,
         gridExpand = FALSE, stack = FALSE)
addStyle(OUT, sheet = "SpiritCraftPerLiters6d", style = NumStyle, cols = 2, rows = 2:1000, 
         gridExpand = FALSE, stack = FALSE)
addStyle(OUT, sheet = "SpiritCraftPerLiters6d", style = NumStyle, cols = 3, rows = 2:1000, 
         gridExpand = FALSE, stack = FALSE)
addStyle(OUT, sheet = "SpiritCraftPerLiters6d", style = NumStyle, cols = 4, rows = 2:1000, 
         gridExpand = FALSE, stack = FALSE)
addStyle(OUT, sheet = "SpiritCraftPerLiters6d", style = NumStyle, cols = 5, rows = 2:1000, 
         gridExpand = FALSE, stack = FALSE)


# Processing Wine per liter

## Creating a subset from the raw SQL file

WinePerLiters <- as.data.frame(df %>% select("fdtmFilingperiod", "fcurWineSold", 
                                 "fcurWineExemptions","fcurWineTaxableLiters", 
                                 "fcurWineTax"))


## Calculating the summations

WinePerLiters <- as.data.frame(aggregate(cbind(fcurWineSold, fcurWineExemptions, fcurWineTaxableLiters, fcurWineTax)~ fdtmFilingperiod, data=WinePerLiters, FUN=sum));


## Renaming the columns

WinePerLiters <- WinePerLiters %>% rename("Filing Period" = fdtmFilingperiod,
                                              "Wine Total Liters sold"= fcurWineSold,
                                              "Wine Deductions / Exemptions Liters" = fcurWineExemptions,
                                              "Wine Taxable Liters sold" = fcurWineTaxableLiters,
                                              "Wine Tax ($)" = fcurWineTax) 


## Formating and exporting to an Excel sheet


setColWidths(OUT, sheet = "WinePerLiters", cols = 1:20, widths ="auto")
freezePane(wb = OUT, sheet = "WinePerLiters" , firstRow = TRUE, firstCol = FALSE)
writeData(wb = OUT, sheet = "WinePerLiters",x = WinePerLiters, 
          startCol = 1, startRow = 1, colNames = TRUE)
addStyle(OUT, sheet = "WinePerLiters", style = LabelStyle, rows = 1, cols = 1:5, 
         gridExpand = TRUE, stack = FALSE)
addStyle(OUT, sheet = "WinePerLiters", style = DateStyle, cols = 1, rows = 2:1000,
         gridExpand = FALSE, stack = FALSE)
addStyle(OUT, sheet = "WinePerLiters", style = NumStyle, cols = 2, rows = 2:1000, 
         gridExpand = FALSE, stack = FALSE)
addStyle(OUT, sheet = "WinePerLiters", style = NumStyle, cols = 3, rows = 2:1000, 
         gridExpand = FALSE, stack = FALSE)
addStyle(OUT, sheet = "WinePerLiters", style = NumStyle, cols = 4, rows = 2:1000, 
         gridExpand = FALSE, stack = FALSE)
addStyle(OUT, sheet = "WinePerLiters", style = NumStyle, cols = 5, rows = 2:1000, 
         gridExpand = FALSE, stack = FALSE)


# Processing Fortified wine per liter

## Creating a subset from the raw SQL file

FWinePerLiters <- as.data.frame(df %>% select("fdtmFilingperiod", "fcurFWineSold", 
                                 "fcurFWineExemptions","fcurFWineTaxableLiters", 
                                 "fcurFWineTax"))


## Calculating the summations

FWinePerLiters <- as.data.frame(aggregate(cbind(fcurFWineSold, fcurFWineExemptions, fcurFWineTaxableLiters, fcurFWineTax)~ fdtmFilingperiod, data=FWinePerLiters, FUN=sum));


## Renaming the columns

FWinePerLiters <- FWinePerLiters %>% rename("Filing Period" = fdtmFilingperiod,
                                              "FWine Total Liters sold"= fcurFWineSold,
                                              "FWine Deductions / Exemptions Liters" = fcurFWineExemptions,
                                              "FWine Taxable Liters sold" = fcurFWineTaxableLiters,
                                              "FWine Tax ($)" = fcurFWineTax) 


## Formating and exporting to an Excel sheet


setColWidths(OUT, sheet = "FWinePerLiters", cols = 1:20, widths ="auto")
freezePane(wb = OUT, sheet = "FWinePerLiters" , firstRow = TRUE, firstCol = FALSE)
writeData(wb = OUT, sheet = "FWinePerLiters",x = FWinePerLiters, 
          startCol = 1, startRow = 1, colNames = TRUE)
addStyle(OUT, sheet = "FWinePerLiters", style = LabelStyle, rows = 1, cols = 1:5, 
         gridExpand = TRUE, stack = FALSE)
addStyle(OUT, sheet = "FWinePerLiters", style = DateStyle, cols = 1, rows = 2:1000,
         gridExpand = FALSE, stack = FALSE)
addStyle(OUT, sheet = "FWinePerLiters", style = NumStyle, cols = 2, rows = 2:1000, 
         gridExpand = FALSE, stack = FALSE)
addStyle(OUT, sheet = "FWinePerLiters", style = NumStyle, cols = 3, rows = 2:1000, 
         gridExpand = FALSE, stack = FALSE)
addStyle(OUT, sheet = "FWinePerLiters", style = NumStyle, cols = 4, rows = 2:1000, 
         gridExpand = FALSE, stack = FALSE)
addStyle(OUT, sheet = "FWinePerLiters", style = NumStyle, cols = 5, rows = 2:1000, 
         gridExpand = FALSE, stack = FALSE)


# Processing Small winery or winegrower

## Processing 9a. First 80,000 liters sold

### Creating a subset from the raw SQL file

SWinePerLiters9a <- as.data.frame(df %>% select("fdtmFilingperiod", "fcurSWineSoldA", 
                                 "fcurSWineExemptionsA","fcurSWineTaxableLitersA", 
                                 "fcurSWineTaxA"))


### Calculating the summations

SWinePerLiters9a <- as.data.frame(aggregate(cbind(fcurSWineSoldA, fcurSWineExemptionsA, fcurSWineTaxableLitersA, fcurSWineTaxA)~ fdtmFilingperiod, data=SWinePerLiters9a, FUN=sum));


### Renaming the columns

SWinePerLiters9a <- SWinePerLiters9a %>% rename("Filing Period" = fdtmFilingperiod,
                                              "SWine (First 80,000 liters sold) Total Liters sold"= fcurSWineSoldA,
                                              "SWine Deductions / Exemptions Liters" = fcurSWineExemptionsA,
                                              "SWine Taxable Liters sold" = fcurSWineTaxableLitersA,
                                              "SWine (First 80,000 liters sold) Tax ($)" = fcurSWineTaxA) 


### Formating and exporting to an Excel sheet


setColWidths(OUT, sheet = "SWinePerLiters9a", cols = 1:20, widths ="auto")
freezePane(wb = OUT, sheet = "SWinePerLiters9a" , firstRow = TRUE, firstCol = FALSE)
writeData(wb = OUT, sheet = "SWinePerLiters9a",x = SWinePerLiters9a, 
          startCol = 1, startRow = 1, colNames = TRUE)
addStyle(OUT, sheet = "SWinePerLiters9a", style = LabelStyle, rows = 1, cols = 1:5, 
         gridExpand = TRUE, stack = FALSE)
addStyle(OUT, sheet = "SWinePerLiters9a", style = DateStyle, cols = 1, rows = 2:1000,
         gridExpand = FALSE, stack = FALSE)
addStyle(OUT, sheet = "SWinePerLiters9a", style = NumStyle, cols = 2, rows = 2:1000, 
         gridExpand = FALSE, stack = FALSE)
addStyle(OUT, sheet = "SWinePerLiters9a", style = NumStyle, cols = 3, rows = 2:1000, 
         gridExpand = FALSE, stack = FALSE)
addStyle(OUT, sheet = "SWinePerLiters9a", style = NumStyle, cols = 4, rows = 2:1000, 
         gridExpand = FALSE, stack = FALSE)
addStyle(OUT, sheet = "SWinePerLiters9a", style = NumStyle, cols = 5, rows = 2:1000, 
         gridExpand = FALSE, stack = FALSE)


## Processing 9b. Sales: 80,001 to 950,000 liters

### Creating a subset from the raw SQL file

SWinePerLiters9b <- as.data.frame(df %>% select("fdtmFilingperiod", "fcurSWineSoldB", 
                                 "fcurSWineExemptionsB","fcurSWineTaxableLitersB", 
                                 "fcurSWineTaxB"))


### Calculating the summations

SWinePerLiters9b <- as.data.frame(aggregate(cbind(fcurSWineSoldB, fcurSWineExemptionsB, fcurSWineTaxableLitersB, fcurSWineTaxB)~ fdtmFilingperiod, data=SWinePerLiters9b, FUN=sum));


### Renaming the columns

SWinePerLiters9b <- SWinePerLiters9b %>% rename("Filing Period" = fdtmFilingperiod,
                                              "SWine (Sales: 80,001 to 950,000 liters) Total Liters sold"= fcurSWineSoldB,
                                              "SWine Deductions / Exemptions Liters" = fcurSWineExemptionsB,
                                              "SWine Taxable Liters sold" = fcurSWineTaxableLitersB,
                                              "SWine (Sales: 80,001 to 950,000 liters) Tax ($)" = fcurSWineTaxB) 


### Formating and exporting to an Excel sheet


setColWidths(OUT, sheet = "SWinePerLiters9b", cols = 1:20, widths ="auto")
freezePane(wb = OUT, sheet = "SWinePerLiters9b" , firstRow = TRUE, firstCol = FALSE)
writeData(wb = OUT, sheet = "SWinePerLiters9b",x = SWinePerLiters9b, 
          startCol = 1, startRow = 1, colNames = TRUE)
addStyle(OUT, sheet = "SWinePerLiters9b", style = LabelStyle, rows = 1, cols = 1:5, 
         gridExpand = TRUE, stack = FALSE)
addStyle(OUT, sheet = "SWinePerLiters9b", style = DateStyle, cols = 1, rows = 2:1000,
         gridExpand = FALSE, stack = FALSE)
addStyle(OUT, sheet = "SWinePerLiters9b", style = NumStyle, cols = 2, rows = 2:1000, 
         gridExpand = FALSE, stack = FALSE)
addStyle(OUT, sheet = "SWinePerLiters9b", style = NumStyle, cols = 3, rows = 2:1000, 
         gridExpand = FALSE, stack = FALSE)
addStyle(OUT, sheet = "SWinePerLiters9b", style = NumStyle, cols = 4, rows = 2:1000, 
         gridExpand = FALSE, stack = FALSE)
addStyle(OUT, sheet = "SWinePerLiters9b", style = NumStyle, cols = 5, rows = 2:1000, 
         gridExpand = FALSE, stack = FALSE)


## Processing 9c. Sales: 950,001 to 1,500,000 liters

### Creating a subset from the raw SQL file

SWinePerLiters9c <- as.data.frame(df %>% select("fdtmFilingperiod", "fcurSWineSoldC", 
                                 "fcurSWineExemptionsC","fcurSWineTaxableLitersC", 
                                 "fcurSWineTaxC"))


### Calculating the summations

SWinePerLiters9c <- as.data.frame(aggregate(cbind(fcurSWineSoldC, fcurSWineExemptionsC, fcurSWineTaxableLitersC, fcurSWineTaxC)~ fdtmFilingperiod, data=SWinePerLiters9c, FUN=sum));


### Renaming the columns

SWinePerLiters9c <- SWinePerLiters9c %>% rename("Filing Period" = fdtmFilingperiod,
                                              "SWine (Sales: 950,001 to 1,500,000 liters) Total Liters sold"= fcurSWineSoldC,
                                              "SWine Deductions / Exemptions Liters" = fcurSWineExemptionsC,
                                              "SWine Taxable Liters sold" = fcurSWineTaxableLitersC,
                                              "SWine (Sales: 950,001 to 1,500,000 liters) Tax ($)" = fcurSWineTaxC) 


### Formating and exporting to an Excel sheet


setColWidths(OUT, sheet = "SWinePerLiters9c", cols = 1:20, widths ="auto")
freezePane(wb = OUT, sheet = "SWinePerLiters9c" , firstRow = TRUE, firstCol = FALSE)
writeData(wb = OUT, sheet = "SWinePerLiters9c",x = SWinePerLiters9c, 
          startCol = 1, startRow = 1, colNames = TRUE)
addStyle(OUT, sheet = "SWinePerLiters9c", style = LabelStyle, rows = 1, cols = 1:5, 
         gridExpand = TRUE, stack = FALSE)
addStyle(OUT, sheet = "SWinePerLiters9c", style = DateStyle, cols = 1, rows = 2:1000,
         gridExpand = FALSE, stack = FALSE)
addStyle(OUT, sheet = "SWinePerLiters9c", style = NumStyle, cols = 2, rows = 2:1000, 
         gridExpand = FALSE, stack = FALSE)
addStyle(OUT, sheet = "SWinePerLiters9c", style = NumStyle, cols = 3, rows = 2:1000, 
         gridExpand = FALSE, stack = FALSE)
addStyle(OUT, sheet = "SWinePerLiters9c", style = NumStyle, cols = 4, rows = 2:1000, 
         gridExpand = FALSE, stack = FALSE)
addStyle(OUT, sheet = "SWinePerLiters9c", style = NumStyle, cols = 5, rows = 2:1000, 
         gridExpand = FALSE, stack = FALSE)


# Saving the excel workbook

saveWorkbook(OUT, "//trdecomsrv/H/Liquor Excise Tax/Data/Liquor Tax Tracker/output/TRD Liquor Tracker v.1.3.xlsx", overwrite = TRUE)

 