## Code to create a RP500 Industry database
# Set working directory
setwd("//trdecomsrv/H/CRS Reports/R Backup Databases/RP500 Industry Datasets/")


libraries <- c("readxl", "writexl", "dplyr", "foreign", "lubridate", "haven", "purrr", "sas7bdat", "tidyr")
lapply(libraries, library, character.only = TRUE)


#read and combine files
files <- list.files(
  path = "//trdecomsrv/H/CRS Reports/R Backup Databases/RP500Ind Excel post 2018/New folder/",
  pattern = "\\.xls$|\\.xlsx$",
  full.names = TRUE
)

#create R combined files
RP500Ind <- sapply(files, read_excel, simplify = FALSE) %>% bind_rows(.id = "id")
#delete columns that are blank
RP500Ind$id <- NULL
RP500Ind$...4 <- NULL
RP500Ind$...5 <- NULL
RP500Ind$...6 <- NULL
RP500Ind$...7 <- NULL
RP500Ind$...9 <- NULL
RP500Ind$...11 <- NULL
RP500Ind$...12 <- NULL
RP500Ind$...18 <- NULL
RP500Ind$...19 <- NULL
RP500Ind$...20 <- NULL
RP500Ind$...22 <- NULL

#rename the columns
names(RP500Ind)[names(RP500Ind) == "New Mexico"] <- "County"
names(RP500Ind)[names(RP500Ind) == "...2"] <- "Industry"
names(RP500Ind)[names(RP500Ind) == "...3"] <- "Type"
names(RP500Ind)[names(RP500Ind) == "...8"] <- "Returns"
names(RP500Ind)[names(RP500Ind) == "...10"] <- "TotalGR"
names(RP500Ind)[names(RP500Ind) == "...13"] <- "TaxableGR"
names(RP500Ind)[names(RP500Ind) == "...14"] <- "MatchedTGR"
names(RP500Ind)[names(RP500Ind) == "...15"] <- "TaxDue"
names(RP500Ind)[names(RP500Ind) == "...16"] <- "Taxpaid"
names(RP500Ind)[names(RP500Ind) == "...17"] <- "RecipientDue"
names(RP500Ind)[names(RP500Ind) == "...21"] <- "RecipientPaid"
names(RP500Ind)[names(RP500Ind) == "...23"] <- "Adjustment"

#fill in missing values
RP500Ind <- fill(RP500Ind, County, Industry, .direction = c("down"))

RP500 <- subset(RP500Ind, `County`!="Taxation and Revenue" & `County`!="Local Government" 
                & `County`!="Distribution" & `County`!="by Industry"
                & `County`!="Reporting Month:" & `County`!="Distribution Month:"
                & `County`!="Report Run:" & `County`!="From:"
                & `County`!="To:", ) 

today <- Sys.Date()
start_date <- today %m-% months(189)
end_date <- start_date + years(30)  # Add 30 years to the current date

monthly_dates <- seq(start_date, end_date, by = "month")
formatted_dates_5 <- format(monthly_dates, "%b-%Y")
formatted_dates_5

formatted_dates_6 <- format(monthly_dates, "%Y-%m")
formatted_dates_6

#add activity month based on accrual month
Activity <- RP500Ind %>%
  mutate(
    Activity = case_when(
      RP500Ind$TotalGR %in% formatted_dates_5 ~ as.character(RP500Ind$TotalGR %in% formatted_dates_6),
      # FALSE ~ NA_character_ 
      
      
    )
  )  

Activity <- Activity[!grepl("Reporting Month:", Activity$County), ]
# Activity$Activity <- lag(lag(Activity$TotalGR)[Activity$Activity == FALSE])

Activity$Activity <- Activity$TotalGR[Activity$Activity == FALSE]


#fill in missing activity dates
Activity <- fill(Activity, Activity, .direction = c("down"))
#remove rows where type equals NA
Activity <- Activity[!is.na(Activity$Type), ]
#create Location column duplicating County
Activity$Location=Activity$County
#replace county name with location code
Activity$Location[Activity$Location=="Santa Fe County"] <- "01001"
Activity$Location[Activity$Location=="Santa Fe, City of"] <- "01123"
Activity$Location[Activity$Location=="Edgewood, Town of"] <- "01320"
Activity$Location[Activity$Location=="Pueblo Of Nambe"] <- "01952"
Activity$Location[Activity$Location=="Pojoaque Pueblo"] <- "01962"
Activity$Location[Activity$Location=="Bernalillo County"] <- "02002"
Activity$Location[Activity$Location=="Albuquerque, City of"] <- "02100"
Activity$Location[Activity$Location=="Los Ranchosde Alb"] <- "02200"
Activity$Location[Activity$Location=="Tijeras"] <- "02318"
Activity$Location[Activity$Location=="Mesa Del Sol District"] <- "02606"
Activity$Location[Activity$Location=="Eddy County"] <- "03003"
Activity$Location[Activity$Location=="Carlsbad"] <- "03106"
Activity$Location[Activity$Location=="Artesia"] <- "03205"
Activity$Location[Activity$Location=="Hope"] <- "03304"
Activity$Location[Activity$Location=="Loving"] <- "03403"
Activity$Location[Activity$Location=="Chaves County"] <- "04004"
Activity$Location[Activity$Location=="Roswell"] <- "04101"
Activity$Location[Activity$Location=="Dexter"] <- "04201"
Activity$Location[Activity$Location=="Hagerman"] <- "04300"
Activity$Location[Activity$Location=="Lake Arthur"] <- "04400"
Activity$Location[Activity$Location=="Curry County"] <- "05005"
Activity$Location[Activity$Location=="Clovis"] <- "05103"
Activity$Location[Activity$Location=="Village Of Grady"] <- "015203"
Activity$Location[Activity$Location=="Texico"] <- "05302"
Activity$Location[Activity$Location=="Melrose"] <- "05402"
Activity$Location[Activity$Location=="Lea County"] <- "06006"
Activity$Location[Activity$Location=="Hobbs"] <- "06111"
Activity$Location[Activity$Location=="Eunice"] <- "06210"
Activity$Location[Activity$Location=="Jal"] <- "06306"
Activity$Location[Activity$Location=="Lovington"] <- "06405"
Activity$Location[Activity$Location=="Tatum"] <- "06500"
Activity$Location[Activity$Location=="Dona Ana County"] <- "07007"
Activity$Location[Activity$Location=="Las Cruces, City of"] <- "07105"
Activity$Location[Activity$Location=="Hatch"] <- "07204"
Activity$Location[Activity$Location=="Mesilla, Town of"] <- "07303"
Activity$Location[Activity$Location=="Sunland Park"] <- "07416"
Activity$Location[Activity$Location=="Grant County"] <- "08008"
Activity$Location[Activity$Location=="Silver City"] <- "08107"
Activity$Location[Activity$Location=="Bayard"] <- "08206"
Activity$Location[Activity$Location=="Santa Clara"] <- "08305"
Activity$Location[Activity$Location=="Hurley"] <- "08404"
Activity$Location[Activity$Location=="Colfax County"] <- "09009"
Activity$Location[Activity$Location=="Raton"] <- "09102"
Activity$Location[Activity$Location=="Maxwell"] <- "09202"
Activity$Location[Activity$Location=="Springer"] <- "09301"
Activity$Location[Activity$Location=="Cimarron"] <- "09401"
Activity$Location[Activity$Location=="Eagle Nest"] <- "09509"
Activity$Location[Activity$Location=="Angel Fire"] <- "09600"
Activity$Location[Activity$Location=="Quay County"] <- "10010"
Activity$Location[Activity$Location=="Tucumcari"] <- "10117"
Activity$Location[Activity$Location=="San Jon"] <- "10214"
Activity$Location[Activity$Location=="Logan"] <- "10309"
Activity$Location[Activity$Location=="House"] <- "10407"
Activity$Location[Activity$Location=="Roosevelt County"] <- "11011"
Activity$Location[Activity$Location=="Portales"] <- "11119"
Activity$Location[Activity$Location=="Elida"] <- "11216"
Activity$Location[Activity$Location=="Dora"] <- "11310"
Activity$Location[Activity$Location=="Causey"] <- "11408"
Activity$Location[Activity$Location=="Floyd"] <- "11502"
Activity$Location[Activity$Location=="San Miguel Co"] <- "12012"
Activity$Location[Activity$Location=="Las Vegas"] <- "12122"
Activity$Location[Activity$Location=="Pecos"] <- "12313"
Activity$Location[Activity$Location=="Mckinley County"] <- "13013"
Activity$Location[Activity$Location=="Gallup"] <- "13114"
Activity$Location[Activity$Location=="Valencia County"] <- "14014"
Activity$Location[Activity$Location=="Belen"] <- "14129"
Activity$Location[Activity$Location=="Los Lunas, Village of"] <- "14316"
Activity$Location[Activity$Location=="Peralta"] <- "14412"
Activity$Location[Activity$Location=="Bosque Farms"] <- "14505"
Activity$Location[Activity$Location=="Otero County"] <- "15015"
Activity$Location[Activity$Location=="Alamogordo"] <- "15116"
Activity$Location[Activity$Location=="Cloudcroft"] <- "15213"
Activity$Location[Activity$Location=="Tularosa"] <- "15308"
Activity$Location[Activity$Location=="San Juan County"] <- "16016"
Activity$Location[Activity$Location=="Farmington"] <- "16121"
Activity$Location[Activity$Location=="Aztec"] <- "16218"
Activity$Location[Activity$Location=="Bloomfield"] <- "16312"
Activity$Location[Activity$Location=="Rio Arriba County"] <- "17017"
Activity$Location[Activity$Location=="Chama"] <- "17118"
Activity$Location[Activity$Location=="Espanola"] <- "17215"
Activity$Location[Activity$Location=="Santa Clara Pueblo"] <- "17904"
Activity$Location[Activity$Location=="Union County"] <- "18018"
Activity$Location[Activity$Location=="Clayton"] <- "18128"
Activity$Location[Activity$Location=="Des Moines"] <- "18224"
Activity$Location[Activity$Location=="Folsom"] <- "18411"
Activity$Location[Activity$Location=="Luna County"] <- "19019"
Activity$Location[Activity$Location=="Deming"] <- "19113"
Activity$Location[Activity$Location=="Columbus"] <- "19212"
Activity$Location[Activity$Location=="Taos County"] <- "20020"
Activity$Location[Activity$Location=="Taos"] <- "20126"
Activity$Location[Activity$Location=="Questa"] <- "20222"
Activity$Location[Activity$Location=="Red River"] <- "20317"
Activity$Location[Activity$Location=="Taos Ski Valley, Village of"] <- "20414"
Activity$Location[Activity$Location=="Sierra County"] <- "21021"
Activity$Location[Activity$Location=="T or C"] <- "21124"
Activity$Location[Activity$Location=="Williamsburg"] <- "21220"
Activity$Location[Activity$Location=="Elephant Butte, City of"] <- "21319"
Activity$Location[Activity$Location=="Torrance County"] <- "22022"
Activity$Location[Activity$Location=="Mountainair"] <- "22127"
Activity$Location[Activity$Location=="Moriarty"] <- "22223"
Activity$Location[Activity$Location=="Willard"] <- "22314"
Activity$Location[Activity$Location=="Encino"] <- "22410"
Activity$Location[Activity$Location=="Estancia"] <- "22503"
Activity$Location[Activity$Location=="Hidalgo County"] <- "23023"
Activity$Location[Activity$Location=="Lordsburg"] <- "23110"
Activity$Location[Activity$Location=="Virden, Village of"] <- "23209"
Activity$Location[Activity$Location=="Guadalupe County"] <- "24024"
Activity$Location[Activity$Location=="Santa Rosa"] <- "24108"
Activity$Location[Activity$Location=="Vaughn"] <- "24207"
Activity$Location[Activity$Location=="Socorro County"] <- "25025"
Activity$Location[Activity$Location=="Socorro, City of"] <- "25125"
Activity$Location[Activity$Location=="Magdalena"] <- "25221"
Activity$Location[Activity$Location=="Lincoln County"] <- "26026"
Activity$Location[Activity$Location=="Ruidoso, Village of"] <- "26112"
Activity$Location[Activity$Location=="Capitan"] <- "26211"
Activity$Location[Activity$Location=="Carrizozo"] <- "26307"
Activity$Location[Activity$Location=="Corona"] <- "26406"
Activity$Location[Activity$Location=="Ruidoso Downs"] <- "26501"
Activity$Location[Activity$Location=="De Baca County"] <- "27027"
Activity$Location[Activity$Location=="Ft Sumner"] <- "27104"
Activity$Location[Activity$Location=="Catron County"] <- "28028"
Activity$Location[Activity$Location=="Reserve"] <- "28130"
Activity$Location[Activity$Location=="Sandoval County"] <- "29029"
Activity$Location[Activity$Location=="Bernalillo"] <- "29120"
Activity$Location[Activity$Location=="Jemez Springs"] <- "29217"
Activity$Location[Activity$Location=="Cuba"] <- "29311"
Activity$Location[Activity$Location=="San Ysidro"] <- "29409"
Activity$Location[Activity$Location=="Corrales"] <- "29504"
Activity$Location[Activity$Location=="Rio Rancho"] <- "29524"
Activity$Location[Activity$Location=="Sandia, Pueblo of"] <- "29912"
Activity$Location[Activity$Location=="Jicarilla Apache Nation"] <- "29932"
Activity$Location[Activity$Location=="Santa Ana Pueblo"] <- "29952"
Activity$Location[Activity$Location=="Cochiti Pueblo"] <- "29972"
Activity$Location[Activity$Location=="Santo Domingo Pueblo"] <- "29974"
Activity$Location[Activity$Location=="Mora County"] <- "30030"
Activity$Location[Activity$Location=="Wagon Mound"] <- "30115"
Activity$Location[Activity$Location=="Harding County"] <- "31031"
Activity$Location[Activity$Location=="Roy, Village Of"] <- "31109"
Activity$Location[Activity$Location=="Mosquero"] <- "31208"
Activity$Location[Activity$Location=="Los Alamos"] <- "32032"
Activity$Location[Activity$Location=="Cibola County"] <- "33033"
Activity$Location[Activity$Location=="Milan"] <- "33131"
Activity$Location[Activity$Location=="Grants"] <- "33227"
Activity$Location[Activity$Location=="Laguna, Pueblo of"] <- "33902"
Activity$Location[Activity$Location=="State Park & Rec Area Capital"] <- "CRSEMN"
Activity$Location[Activity$Location=="Office of Cultural Affairs"] <- "CRSOCA"
Activity$Location[Activity$Location=="NMFA Public Project Revolving Fund"] <- "CRSPPR"
Activity$Location[Activity$Location=="NMFA Public Project Revolving"] <- "CRSPPR"
Activity$Location[Activity$Location=="NM Youth Conservation Corp"] <- "CRSYCC"
Activity$Location[Activity$Location=="Leased Vehicle - Infrastructure"] <- "S444"
Activity$Location[Activity$Location=="Leased Vehicle -Infrastructure"] <- "S444"
Activity$Location[Activity$Location=="Leased Vehicle - County Road"] <- "S444C"
Activity$Location[Activity$Location=="General Fund - Gross Receipts - CRS"] <- "SGRT"
Activity$Location[Activity$Location=="General Fund-Gross Receipts-CRS"] <- "SGRT"
Activity$Location[Activity$Location=="County Supported Medicaid Fund"] <- "SMEDIC"
Activity$Location[Activity$Location=="Grand Total"] <- "ZZZZZ"
Activity$Location[Activity$Location=="SF Indian School"] <- "01907"
Activity$Location[Activity$Location=="San Ildefonso Pueblo"] <- "01975"
Activity$Location[Activity$Location=="Anthony, City of"] <- "07507"
Activity$Location[Activity$Location=="Kirtland"] <- "16323"
Activity$Location[Activity$Location=="Ohkay Owingeh Pueblo"] <- "17942"
Activity$Location[Activity$Location=="Ohkay Owingeh Puebo"] <- "17942"
Activity$Location[Activity$Location=="Taos Pueblo"] <- "20913"
Activity$Location[Activity$Location=="Picuris Pueblo"] <- "20918"
Activity$Location[Activity$Location=="Grenville"] <- "18315"
Activity$Location[Activity$Location=="Tesuque Pueblo"] <- "01953"
Activity$Location[Activity$Location=="19 Pueblos District"] <- "02905"
Activity$Location[Activity$Location=="Rio Communities"] <- "14037"
Activity$Location[Activity$Location=="Acoma Pueblo"] <- "33909"
Activity$Location[Activity$Location=="AIS Property/Nineteen Pueblo"] <- "02905"
Activity$Location[Activity$Location=="Zuni Pueblo"] <- "13901"
#reorder columns, move activity and location to the beginning
RP500Indthirdpart <- Activity[c("Activity", "County", "Location", "Industry", "Type", "Returns", "TotalGR", "TaxableGR", "MatchedTGR", "TaxDue", "Taxpaid", "RecipientDue", "RecipientPaid", "Adjustment")]
#format values
RP500Indthirdpart$Returns <- as.numeric(RP500Indthirdpart$Returns)
RP500Indthirdpart$TotalGR <- as.numeric(RP500Indthirdpart$TotalGR)
RP500Indthirdpart$TaxableGR <- as.numeric(RP500Indthirdpart$TaxableGR)
RP500Indthirdpart$MatchedTGR <- as.numeric(RP500Indthirdpart$MatchedTGR)
RP500Indthirdpart$TaxDue <- as.numeric(RP500Indthirdpart$TaxDue)
RP500Indthirdpart$Taxpaid <- as.numeric(RP500Indthirdpart$Taxpaid)
RP500Indthirdpart$RecipientDue <- as.numeric(RP500Indthirdpart$RecipientDue)
RP500Indthirdpart$RecipientPaid <- as.numeric(RP500Indthirdpart$RecipientPaid)


#save as R file
save(RP500Indthirdpart, file = "//trdecomsrv/H/CRS Reports/R Backup Databases/RP500 Industry Datasets/RP500Ind third Part_Alphie.RData")


### Combine Rp 500 Industry R data files and save to SAS and STATA

library(writexl)
library(foreign)
library(haven)
library(sas7bdat)

#RP500 Industry database
load("//trdecomsrv/H/CRS Reports/R Backup Databases/RP500 Industry Datasets/RP500Ind through 2018-04.RData")
load("//trdecomsrv/H/CRS Reports/R Backup Databases/RP500 Industry Datasets/RP500Ind Second Part.RData")
load("//trdecomsrv/H/CRS Reports/R Backup Databases/RP500 Industry Datasets/RP500Ind third Part_Alphie.RData")

#combine the three data frames
RP500IndCombined <- rbind(RP500IndComb, RP500Indsecpart, RP500Indthirdpart)

#save as R file
save(RP500IndCombined, file = "//trdecomsrv/H/CRS Reports/R Backup Databases/RP500 Industry Datasets/RP500Ind Unredacted Database_Alphie.RData")
#save as SAS file
write.foreign(df = RP500IndCombined,
              datafile = 'RP500IndCombined.RData',
              codefile = 'RP500Ind Unredacted Database_ALphie.sas',
              package = 'SAS')
#save as STATA file
require(foreign)
write.dta(RP500IndCombined, "//trdecomsrv/H/CRS Reports/STATA Databases/RP500Ind Unredacted Database_Alphie.dta")