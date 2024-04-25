############################################################################
# This script will generate the RP500 Distributions GRT Database
############################################################################

libraries <- c("readxl", "writexl", "dplyr", "foreign", "lubridate", "haven", "purrr", "sas7bdat", "tidyr")
lapply(libraries, library, character.only = TRUE)


#read and combine files you pull out from GenTax
files <- list.files(path="//trdecomsrv/H/CRS Reports/RP-500 database/Unredacted/DistributionByAccrualMonth", 
                    pattern = ".xlsx", full.names = T)

# #create R combined files
RP500DistGRT <- map_dfr(files, read_excel, .id = "id")

#delete unnecessary columns
RP500DistGRT$id <- NULL
RP500DistGRT$...9 <- NULL

#rename the columns
names(RP500DistGRT)[names(RP500DistGRT) == "New Mexico Taxation and"] <- "Option"

#delete "Option" column
RP500DistGRT$Option <- NULL

#rename the columns
names(RP500DistGRT)[names(RP500DistGRT) == "...2"] <- "Tax"
names(RP500DistGRT)[names(RP500DistGRT) == "...3"] <- "Penalty"
names(RP500DistGRT)[names(RP500DistGRT) == "...4"] <- "Interest"
names(RP500DistGRT)[names(RP500DistGRT) == "...5"] <- "Admin"
names(RP500DistGRT)[names(RP500DistGRT) == "...6"] <- "Contracts"
names(RP500DistGRT)[names(RP500DistGRT) == "...7"] <- "Payback"
names(RP500DistGRT)[names(RP500DistGRT) == "...8"] <- "Disbursement"

# Duplicate the first column
RP500DistGRT$Option <- RP500DistGRT$`New Mexico     Taxation and`
# Now delete the first column
RP500DistGRT$`New Mexico     Taxation and` <- NULL

today <- Sys.Date()
start_date <- today %m-% months(189)
end_date <- start_date + years(30)  # Add 30 years to the current date

monthly_dates <- seq(start_date, end_date, by = "month")
formatted_dates_5 <- format(monthly_dates, "%B-%Y")
formatted_dates_5

formatted_dates_6 <- format(monthly_dates, "%Y-%m")
formatted_dates_6


# Reorder columns
RP500DistGRT <- RP500DistGRT[c("Option", "Tax", "Penalty", "Interest", "Admin", "Contracts", "Payback", "Disbursement")]

#add activity month based on accrual month
Activity <- RP500DistGRT %>%
  mutate(
  Activity = case_when(
    RP500DistGRT$Tax %in% formatted_dates_5 ~ as.character(RP500DistGRT$Tax %in% formatted_dates_6),
    )
)  

Activity$Activity <- Activity$Tax[Activity$Activity == FALSE]

# Delete rows with unnecessary information
RP500DistGRT <- subset(Activity, `Option`!="Taxation and Revenue" & `Option`!="Local Government Distribution" 
                & `Option`!="Revenue" & `Option`!="Summary"
                & `Option`!="Reporting Month:" & `Option`!="Distribution Month:"
                & `Option`!="Report Run:" & `Option`!="From:"
                & `Option`!="To:", )

#fill in missing activity dates
RP500DistGRT <- fill(RP500DistGRT, Activity, .direction = c("down"))

#add county names
County <- RP500DistGRT %>% mutate(
  County = case_when(
    RP500DistGRT$Option == "Santa Fe County" ~ "Santa Fe County",
    RP500DistGRT$Option == "Santa Fe, City of" ~ "Santa Fe, City of",
    RP500DistGRT$Option == 	"Edgewood, Town of"	 ~ 	"Edgewood, Town of",
    RP500DistGRT$Option == 	"Pueblo Of Nambe"	 ~ 	"Pueblo Of Nambe",
    RP500DistGRT$Option == 	"Pojoaque Pueblo"	 ~ 	"Pojoaque Pueblo",
    RP500DistGRT$Option == 	"Bernalillo County"	 ~ 	"Bernalillo County",
    RP500DistGRT$Option == 	"Albuquerque, City of"	 ~ 	"Albuquerque, City of",
    RP500DistGRT$Option == 	"Los Ranchosde Alb"	 ~ 	"Los Ranchosde Alb",
    RP500DistGRT$Option == 	"Tijeras"	 ~ 	"Tijeras",
    RP500DistGRT$Option == 	"Mesa Del Sol District"	 ~ 	"Mesa Del Sol District",
    RP500DistGRT$Option == 	"Eddy County"	~	"Eddy County",
    RP500DistGRT$Option == 	"Carlsbad"	 ~ 	"Carlsbad",
    RP500DistGRT$Option == 	"Artesia"	 ~ 	"Artesia",
    RP500DistGRT$Option == 	"Hope"	 ~ 	"Hope",
    RP500DistGRT$Option == 	"Loving"	 ~ 	"Loving",
    RP500DistGRT$Option == 	"Chaves County"	 ~ 	"Chaves County",
    RP500DistGRT$Option == 	"Roswell"	 ~ 	"Roswell",
    RP500DistGRT$Option == 	"Dexter"	 ~ 	"Dexter",
    RP500DistGRT$Option == 	"Hagerman"	 ~ 	"Hagerman",
    RP500DistGRT$Option == 	"Lake Arthur"	 ~ 	"Lake Arthur",
    RP500DistGRT$Option == 	"Curry County"	 ~ 	"Curry County",
    RP500DistGRT$Option == 	"Clovis"	 ~ 	"Clovis",
    RP500DistGRT$Option == 	"Village Of Grady"	 ~ 	"Village Of Grady",
    RP500DistGRT$Option == 	"Texico"	~	"Texico",
    RP500DistGRT$Option == 	"Melrose"	 ~ 	"Melrose",
    RP500DistGRT$Option == 	"Lea County"	 ~ 	"Lea County",
    RP500DistGRT$Option == 	"Hobbs"	 ~ 	"Hobbs",
    RP500DistGRT$Option == 	"Eunice"	 ~ 	"Eunice",
    RP500DistGRT$Option == 	"Jal"	 ~ 	"Jal",
    RP500DistGRT$Option == 	"Lovington"	 ~ 	"Lovington",
    RP500DistGRT$Option == 	"Tatum"	 ~ 	"Tatum",
    RP500DistGRT$Option == 	"Dona Ana County"	 ~ 	"Dona Ana County",
    RP500DistGRT$Option == 	"Las Cruces, City of"	 ~ 	"Las Cruces, City of",
    RP500DistGRT$Option == 	"Hatch"	 ~ 	"Hatch",
    RP500DistGRT$Option == 	"Mesilla, Town of"	 ~ 	"Mesilla, Town of",
    RP500DistGRT$Option == 	"Sunland Park"	 ~ 	"Sunland Park",
    RP500DistGRT$Option == 	"Grant County"	~	"Grant County",
    RP500DistGRT$Option == 	"Silver City"	 ~ 	"Silver City",
    RP500DistGRT$Option == 	"Bayard"	 ~ 	"Bayard",
    RP500DistGRT$Option == 	"Santa Clara"	 ~ 	"Santa Clara",
    RP500DistGRT$Option == 	"Hurley"	 ~ 	"Hurley",
    RP500DistGRT$Option == 	"Colfax County"	 ~ 	"Colfax County",
    RP500DistGRT$Option == 	"Raton"	 ~ 	"Raton",
    RP500DistGRT$Option == 	"Maxwell"	 ~ 	"Maxwell",
    RP500DistGRT$Option == 	"Springer"	 ~ 	"Springer",
    RP500DistGRT$Option == 	"Cimarron"	 ~ 	"Cimarron",
    RP500DistGRT$Option == 	"Eagle Nest"	 ~ 	"Eagle Nest",
    RP500DistGRT$Option == 	"Angel Fire"	 ~ 	"Angel Fire",
    RP500DistGRT$Option == 	"Quay County"	 ~ 	"Quay County",
    RP500DistGRT$Option == 	"Tucumcari"	~	"Tucumcari",
    RP500DistGRT$Option == 	"San Jon"	 ~ 	"San Jon",
    RP500DistGRT$Option == 	"Logan"	 ~ 	"Logan",
    RP500DistGRT$Option == 	"House"	 ~ 	"House",
    RP500DistGRT$Option == 	"Roosevelt County"	 ~ 	"Roosevelt County",
    RP500DistGRT$Option == 	"Portales"	 ~ 	"Portales",
    RP500DistGRT$Option == 	"Elida"	 ~ 	"Elida",
    RP500DistGRT$Option == 	"Dora"	 ~ 	"Dora",
    RP500DistGRT$Option == 	"Causey"	 ~ 	"Causey",
    RP500DistGRT$Option == 	"Floyd"	 ~ 	"Floyd",
    RP500DistGRT$Option == 	"San Miguel Co"	 ~ 	"San Miguel Co",
    RP500DistGRT$Option == 	"Las Vegas"	 ~ 	"Las Vegas",
    RP500DistGRT$Option == 	"Pecos"	 ~ 	"Pecos",
    RP500DistGRT$Option == 	"Mckinley County"	~	"Mckinley County",
    RP500DistGRT$Option == 	"McKinley County"	~	"McKinley County",
    RP500DistGRT$Option == 	"Gallup"	 ~ 	"Gallup",
    RP500DistGRT$Option == 	"Valencia County"	 ~ 	"Valencia County",
    RP500DistGRT$Option == 	"Belen"	 ~ 	"Belen",
    RP500DistGRT$Option == 	"Los Lunas, Village of"	 ~ 	"Los Lunas, Village of",
    RP500DistGRT$Option == 	"Peralta"	 ~ 	"Peralta",
    RP500DistGRT$Option == 	"Bosque Farms"	 ~ 	"Bosque Farms",
    RP500DistGRT$Option == 	"Otero County"	 ~ 	"Otero County",
    RP500DistGRT$Option == 	"Alamogordo"	 ~ 	"Alamogordo",
    RP500DistGRT$Option == 	"Cloudcroft"	 ~ 	"Cloudcroft",
    RP500DistGRT$Option == 	"Tularosa"	 ~ 	"Tularosa",
    RP500DistGRT$Option == 	"San Juan County"	 ~ 	"San Juan County",
    RP500DistGRT$Option == 	"Farmington"	 ~ 	"Farmington",
    RP500DistGRT$Option == 	"Aztec"	~	"Aztec",
    RP500DistGRT$Option == 	"Bloomfield"	 ~ 	"Bloomfield",
    RP500DistGRT$Option == 	"Rio Arriba County"	 ~ 	"Rio Arriba County",
    RP500DistGRT$Option == 	"Chama"	 ~ 	"Chama",
    RP500DistGRT$Option == 	"Espanola"	 ~ 	"Espanola",
    RP500DistGRT$Option == 	"Santa Clara Pueblo"	 ~ 	"Santa Clara Pueblo",
    RP500DistGRT$Option == 	"Union County"	 ~ 	"Union County",
    RP500DistGRT$Option == 	"Clayton"	 ~ 	"Clayton",
    RP500DistGRT$Option == 	"Des Moines"	 ~ 	"Des Moines",
    RP500DistGRT$Option == 	"Folsom"	 ~ 	"Folsom",
    RP500DistGRT$Option == 	"Luna County"	 ~ 	"Luna County",
    RP500DistGRT$Option == 	"Deming"	 ~ 	"Deming",
    RP500DistGRT$Option == 	"Columbus"	 ~ 	"Columbus",
    RP500DistGRT$Option == 	"Taos County"	~	"Taos County",
    RP500DistGRT$Option == 	"Taos"	 ~ 	"Taos",
    RP500DistGRT$Option == 	"Questa"	 ~ 	"Questa",
    RP500DistGRT$Option == 	"Red River"	 ~ 	"Red River",
    RP500DistGRT$Option == 	"Taos Ski Valley, Village of"	 ~ 	"Taos Ski Valley, Village of",
    RP500DistGRT$Option == 	"Sierra County"	 ~ 	"Sierra County",
    RP500DistGRT$Option == 	"T or C"	 ~ 	"T or C",
    RP500DistGRT$Option == 	"Williamsburg"	 ~ 	"Williamsburg",
    RP500DistGRT$Option == 	"Elephant Butte, City of"	 ~ 	"Elephant Butte, City of",
    RP500DistGRT$Option == 	"Torrance County"	 ~ 	"Torrance County",
    RP500DistGRT$Option == 	"Mountainair"	 ~ 	"Mountainair",
    RP500DistGRT$Option == 	"Moriarty"	 ~ 	"Moriarty",
    RP500DistGRT$Option == 	"Willard"	 ~ 	"Willard",
    RP500DistGRT$Option == 	"Encino"	~	"Encino",
    RP500DistGRT$Option == 	"Estancia"	 ~ 	"Estancia",
    RP500DistGRT$Option == 	"Hidalgo County"	 ~ 	"Hidalgo County",
    RP500DistGRT$Option == 	"Lordsburg"	 ~ 	"Lordsburg",
    RP500DistGRT$Option == 	"Virden, Village of"	 ~ 	"Virden, Village of",
    RP500DistGRT$Option == 	"Guadalupe County"	 ~ 	"Guadalupe County",
    RP500DistGRT$Option == 	"Santa Rosa"	 ~ 	"Santa Rosa",
    RP500DistGRT$Option == 	"Vaughn"	 ~ 	"Vaughn",
    RP500DistGRT$Option == 	"Socorro County"	 ~ 	"Socorro County",
    RP500DistGRT$Option == 	"Socorro, City of"	 ~ 	"Socorro, City of",
    RP500DistGRT$Option == 	"Magdalena"	 ~ 	"Magdalena",
    RP500DistGRT$Option == 	"Lincoln County"	 ~ 	"Lincoln County",
    RP500DistGRT$Option == 	"Ruidoso, Village of"	 ~ 	"Ruidoso, Village of",
    RP500DistGRT$Option == 	"Capitan"	~	"Capitan",
    RP500DistGRT$Option == 	"Carrizozo"	 ~ 	"Carrizozo",
    RP500DistGRT$Option == 	"Corona"	 ~ 	"Corona",
    RP500DistGRT$Option == 	"Ruidoso Downs"	 ~ 	"Ruidoso Downs",
    RP500DistGRT$Option == 	"De Baca County"	 ~ 	"De Baca County",
    RP500DistGRT$Option == 	"Ft Sumner"	 ~ 	"Ft Sumner",
    RP500DistGRT$Option == 	"Catron County"	 ~ 	"Catron County",
    RP500DistGRT$Option == 	"Reserve"	 ~ 	"Reserve",
    RP500DistGRT$Option == 	"Sandoval County"	 ~ 	"Sandoval County",
    RP500DistGRT$Option == 	"Bernalillo"	 ~ 	"Bernalillo",
    RP500DistGRT$Option == 	"Jemez Springs"	 ~ 	"Jemez Springs",
    RP500DistGRT$Option == 	"Cuba"	 ~ 	"Cuba",
    RP500DistGRT$Option == 	"San Ysidro"	 ~ 	"San Ysidro",
    RP500DistGRT$Option == 	"Corrales"	~	"Corrales",
    RP500DistGRT$Option == 	"Rio Rancho"	 ~ 	"Rio Rancho",
    RP500DistGRT$Option == 	"Sandia, Pueblo of"	 ~ 	"Sandia, Pueblo of",
    RP500DistGRT$Option == 	"Jicarilla Apache Nation"	 ~ 	"Jicarilla Apache Nation",
    RP500DistGRT$Option == 	"Santa Ana Pueblo"	 ~ 	"Santa Ana Pueblo",
    RP500DistGRT$Option == 	"Cochiti Pueblo"	 ~ 	"Cochiti Pueblo",
    RP500DistGRT$Option == 	"Santo Domingo Pueblo"	 ~ 	"Santo Domingo Pueblo",
    RP500DistGRT$Option == 	"Mora County"	 ~ 	"Mora County",
    RP500DistGRT$Option == 	"Wagon Mound"	 ~ 	"Wagon Mound",
    RP500DistGRT$Option == 	"Harding County"	 ~ 	"Harding County",
    RP500DistGRT$Option == 	"Roy, Village Of"	 ~ 	"Roy, Village Of",
    RP500DistGRT$Option == 	"Mosquero"	 ~ 	"Mosquero",
    RP500DistGRT$Option == 	"Los Alamos"	 ~ 	"Los Alamos",
    RP500DistGRT$Option == 	"Cibola County"	~	"Cibola County",
    RP500DistGRT$Option == 	"Milan"	 ~ 	"Milan",
    RP500DistGRT$Option == 	"Grants"	 ~ 	"Grants",
    RP500DistGRT$Option == 	"Laguna, Pueblo of"	 ~ 	"Laguna, Pueblo of",
    RP500DistGRT$Option == 	"State Park & Rec Area Capital"	 ~ 	"State Park & Rec Area Capital",
    RP500DistGRT$Option == 	"Office of Cultural Affairs"	 ~ 	"Office of Cultural Affairs",
    RP500DistGRT$Option == 	"NMFA Public Project Revolving Fund"	 ~ 	"NMFA Public Project Revolving Fund",
    RP500DistGRT$Option == 	"NMFA Public Project Revolving"	 ~ 	"NMFA Public Project Revolving",
    RP500DistGRT$Option == 	"NM Youth Conservation Corp"	 ~ 	"NM Youth Conservation Corp",
    RP500DistGRT$Option == 	"Leased Vehicle - Infrastructure"	 ~ 	"Leased Vehicle - Infrastructure",
    RP500DistGRT$Option == 	"Leased Vehicle -Infrastructure"	 ~ 	"Leased Vehicle -Infrastructure",
    RP500DistGRT$Option == 	"Leased Vehicle - County Road"	 ~ 	"Leased Vehicle - County Road",
    RP500DistGRT$Option == 	"General Fund - Gross Receipts - CRS"	 ~ 	"General Fund - Gross Receipts - GRT",
    RP500DistGRT$Option == 	"General Fund-Gross Receipt-CRS"	~	"General Fund - Gross Receipts - GRT",
    RP500DistGRT$Option == 	"General Fund-Gross Receipt-GRT"	~	"General Fund - Gross Receipts - GRT",
    RP500DistGRT$Option == 	"County Supported Medicaid Fund"	 ~ 	"County Supported Medicaid Fund",
    RP500DistGRT$Option == 	"Grand Total:"	 ~ 	"Grand Total:",
    RP500DistGRT$Option == 	"SF Indian School"	 ~ 	"SF Indian School",
    RP500DistGRT$Option == 	"San Ildefonso Pueblo"	 ~ 	"San Ildefonso Pueblo",
    RP500DistGRT$Option == 	"Anthony, City of"	 ~ 	"Anthony, City of",
    RP500DistGRT$Option == 	"Anthony"	 ~ 	"Anthony, City of",
    RP500DistGRT$Option == 	"Kirtland"	 ~ 	"Kirtland",
    #RP500DistGRT$Option == 	"Ohkay Owingeh Pueblo"	 ~ 	"Ohkay Owingeh Pueblo",
    RP500DistGRT$Option == 	"Ohkay Owingeh Puebo"	 ~ 	"Ohkay Owingeh Pueblo",
    RP500DistGRT$Option == 	"Taos Pueblo"	 ~ 	"Taos Pueblo",
    RP500DistGRT$Option == 	"Picuris Pueblo"	 ~ 	"Picuris Pueblo",
    RP500DistGRT$Option == 	"Grenville"	 ~ 	"Grenville",
    RP500DistGRT$Option == 	"Tesuque Pueblo"	 ~ 	"Tesuque Pueblo",
    RP500DistGRT$Option == 	"19 Pueblos District"	~	"19 Pueblos District",
    RP500DistGRT$Option == 	"Rio Communities"	 ~ 	"Rio Communities",
    RP500DistGRT$Option == 	"Acoma Pueblo"	 ~ 	"Acoma Pueblo",
    RP500DistGRT$Option == 	"AIS Property/Nineteen Pueblo"	 ~ 	"AIS Property/Nineteen Pueblo",
    RP500DistGRT$Option == 	"Zuni Pueblo"	 ~ 	"Zuni Pueblo",
    RP500DistGRT$Option == 	"Jemez Pueblo"	 ~ 	"Jemez Pueblo",
    RP500DistGRT$Option == 	"Zia Pueblo"	 ~ 	"Zia Pueblo",
    RP500DistGRT$Option == 	"Village at Rio Rancho TIDD"	 ~ 	"Village at Rio Rancho TIDD",
    RP500DistGRT$Option == 	"Los Diamantes TIDD"	 ~ 	"Los Diamantes TIDD",
    RP500DistGRT$Option == 	"Taos Ski Valley TIDD"	 ~ 	"Taos Ski Valley TIDD",
    RP500DistGRT$Option == 	"Lower Petroglyphs TIDD"	 ~ 	"Lower Petroglyphs TIDD",
    RP500DistGRT$Option == 	"Las Cruces TIDD"	 ~ 	"Las Cruces TIDD",
    RP500DistGRT$Option == 	"Stonegate TIDD"	 ~ 	"Stonegate TIDD",
    RP500DistGRT$Option == 	"Winrock Town TIDD Dist 2"	 ~ 	"Winrock Town TIDD Dist 2",
    RP500DistGRT$Option == 	"Winrock Town TIDD"	 ~ 	"Winrock Town TIDD",
    RP500DistGRT$Option == 	"South Campus TIDD"	 ~ 	"South Campus TIDD",
    RP500DistGRT$Option == 	"Santolina TIDD District 18"	 ~ 	"Santolina TIDD District 18",
    RP500DistGRT$Option == 	"Santolina TIDD District 12"	 ~ 	"Santolina TIDD District 12",
    RP500DistGRT$Option == 	"Santolina TIDD District 10"	 ~ 	"Santolina TIDD District 10",
    RP500DistGRT$Option == 	"Santolina TIDD District 7"	 ~ 	"Santolina TIDD District 7",
    RP500DistGRT$Option == 	"Santolina TIDD District 6"	 ~ 	"Santolina TIDD District 6",
    RP500DistGRT$Option == 	"Santolina TIDD District 5"	 ~ 	"Santolina TIDD District 5",
    RP500DistGRT$Option == 	"Santolina TIDD District 4"	 ~ 	"Santolina TIDD District 4",
    RP500DistGRT$Option == 	"Santolina TIDD District 1"	 ~ 	"Santolina TIDD District 1",
    RP500DistGRT$Option == 	"Santolina TIDD District 14"	 ~ 	"Santolina TIDD District 14",
    RP500DistGRT$Option == 	"Santolina TIDD District 2"	 ~ 	"Santolina TIDD District 2",
    RP500DistGRT$Option == 	"Santolina TIDD District 8"	 ~ 	"Santolina TIDD District 8",
    RP500DistGRT$Option == 	"Santolina TIDD District 17"	 ~ 	"Santolina TIDD District 17",
    RP500DistGRT$Option == 	"Santolina TIDD District 20"	 ~ 	"Santolina TIDD District 20",
    RP500DistGRT$Option == 	"Santolina TIDD District 11"	 ~ 	"Santolina TIDD District 11",
    RP500DistGRT$Option == 	"Santolina TIDD District 3"	 ~ 	"Santolina TIDD District 3",
    RP500DistGRT$Option == 	"Santolina TIDD District 19"	 ~ 	"Santolina TIDD District 19",
    RP500DistGRT$Option == 	"Santolina TIDD District 15"	 ~ 	"Santolina TIDD District 15",
    RP500DistGRT$Option == 	"Santolina TIDD District 16"	 ~ 	"Santolina TIDD District 16",
    RP500DistGRT$Option == 	"Santolina TIDD District 9"	 ~ 	"Santolina TIDD District 9",
    RP500DistGRT$Option == 	"Santolina TIDD District 13"	 ~ 	"Santolina TIDD District 13",
    RP500DistGRT$Option == 	"Quorum Uptown TIDD"	 ~ 	"Quorum Uptown TIDD",
    RP500DistGRT$Option == 	"Valley Water and Sanitation Di"	 ~ 	"Valley Water and Sanitation Di",
  )
)

#fill in missing location names dates
RP500DistGRT <- fill(County, County, .direction = c("down"))

# Duplicate County column
RP500DistGRT$Location <- RP500DistGRT$County
#replace location name with location code
RP500DistGRT$Location[RP500DistGRT$Location=="Santa Fe County"] <- "01001"
RP500DistGRT$Location[RP500DistGRT$Location=="Santa Fe, City of"] <- "01123"
RP500DistGRT$Location[RP500DistGRT$Location=="Edgewood, Town of"] <- "01320"
RP500DistGRT$Location[RP500DistGRT$Location=="Pueblo Of Nambe"] <- "01952"
RP500DistGRT$Location[RP500DistGRT$Location=="Pojoaque Pueblo"] <- "01962"
RP500DistGRT$Location[RP500DistGRT$Location=="Bernalillo County"] <- "02002"
RP500DistGRT$Location[RP500DistGRT$Location=="Albuquerque, City of"] <- "02100"
RP500DistGRT$Location[RP500DistGRT$Location=="Los Ranchosde Alb"] <- "02200"
RP500DistGRT$Location[RP500DistGRT$Location=="Tijeras"] <- "02318"
RP500DistGRT$Location[RP500DistGRT$Location=="Mesa Del Sol District"] <- "02606"
RP500DistGRT$Location[RP500DistGRT$Location=="Eddy County"] <- "03003"
RP500DistGRT$Location[RP500DistGRT$Location=="Carlsbad"] <- "03106"
RP500DistGRT$Location[RP500DistGRT$Location=="Artesia"] <- "03205"
RP500DistGRT$Location[RP500DistGRT$Location=="Hope"] <- "03304"
RP500DistGRT$Location[RP500DistGRT$Location=="Loving"] <- "03403"
RP500DistGRT$Location[RP500DistGRT$Location=="Chaves County"] <- "04004"
RP500DistGRT$Location[RP500DistGRT$Location=="Roswell"] <- "04101"
RP500DistGRT$Location[RP500DistGRT$Location=="Dexter"] <- "04201"
RP500DistGRT$Location[RP500DistGRT$Location=="Hagerman"] <- "04300"
RP500DistGRT$Location[RP500DistGRT$Location=="Lake Arthur"] <- "04400"
RP500DistGRT$Location[RP500DistGRT$Location=="Curry County"] <- "05005"
RP500DistGRT$Location[RP500DistGRT$Location=="Clovis"] <- "05103"
RP500DistGRT$Location[RP500DistGRT$Location=="Village Of Grady"] <- "015203"
RP500DistGRT$Location[RP500DistGRT$Location=="Texico"] <- "05302"
RP500DistGRT$Location[RP500DistGRT$Location=="Melrose"] <- "05402"
RP500DistGRT$Location[RP500DistGRT$Location=="Lea County"] <- "06006"
RP500DistGRT$Location[RP500DistGRT$Location=="Hobbs"] <- "06111"
RP500DistGRT$Location[RP500DistGRT$Location=="Eunice"] <- "06210"
RP500DistGRT$Location[RP500DistGRT$Location=="Jal"] <- "06306"
RP500DistGRT$Location[RP500DistGRT$Location=="Lovington"] <- "06405"
RP500DistGRT$Location[RP500DistGRT$Location=="Tatum"] <- "06500"
RP500DistGRT$Location[RP500DistGRT$Location=="Dona Ana County"] <- "07007"
RP500DistGRT$Location[RP500DistGRT$Location=="Las Cruces, City of"] <- "07105"
RP500DistGRT$Location[RP500DistGRT$Location=="Hatch"] <- "07204"
RP500DistGRT$Location[RP500DistGRT$Location=="Mesilla, Town of"] <- "07303"
RP500DistGRT$Location[RP500DistGRT$Location=="Sunland Park"] <- "07416"
RP500DistGRT$Location[RP500DistGRT$Location=="Grant County"] <- "08008"
RP500DistGRT$Location[RP500DistGRT$Location=="Silver City"] <- "08107"
RP500DistGRT$Location[RP500DistGRT$Location=="Bayard"] <- "08206"
RP500DistGRT$Location[RP500DistGRT$Location=="Santa Clara"] <- "08305"
RP500DistGRT$Location[RP500DistGRT$Location=="Hurley"] <- "08404"
RP500DistGRT$Location[RP500DistGRT$Location=="Colfax County"] <- "09009"
RP500DistGRT$Location[RP500DistGRT$Location=="Raton"] <- "09102"
RP500DistGRT$Location[RP500DistGRT$Location=="Maxwell"] <- "09202"
RP500DistGRT$Location[RP500DistGRT$Location=="Springer"] <- "09301"
RP500DistGRT$Location[RP500DistGRT$Location=="Cimarron"] <- "09401"
RP500DistGRT$Location[RP500DistGRT$Location=="Eagle Nest"] <- "09509"
RP500DistGRT$Location[RP500DistGRT$Location=="Angel Fire"] <- "09600"
RP500DistGRT$Location[RP500DistGRT$Location=="Quay County"] <- "10010"
RP500DistGRT$Location[RP500DistGRT$Location=="Tucumcari"] <- "10117"
RP500DistGRT$Location[RP500DistGRT$Location=="San Jon"] <- "10214"
RP500DistGRT$Location[RP500DistGRT$Location=="Logan"] <- "10309"
RP500DistGRT$Location[RP500DistGRT$Location=="House"] <- "10407"
RP500DistGRT$Location[RP500DistGRT$Location=="Roosevelt County"] <- "11011"
RP500DistGRT$Location[RP500DistGRT$Location=="Portales"] <- "11119"
RP500DistGRT$Location[RP500DistGRT$Location=="Elida"] <- "11216"
RP500DistGRT$Location[RP500DistGRT$Location=="Dora"] <- "11310"
RP500DistGRT$Location[RP500DistGRT$Location=="Causey"] <- "11408"
RP500DistGRT$Location[RP500DistGRT$Location=="Floyd"] <- "11502"
RP500DistGRT$Location[RP500DistGRT$Location=="San Miguel Co"] <- "12012"
RP500DistGRT$Location[RP500DistGRT$Location=="Las Vegas"] <- "12122"
RP500DistGRT$Location[RP500DistGRT$Location=="Pecos"] <- "12313"
RP500DistGRT$Location[RP500DistGRT$Location=="Mckinley County"] <- "13013"
RP500DistGRT$Location[RP500DistGRT$Location=="McKinley County"] <- "13013"
RP500DistGRT$Location[RP500DistGRT$Location=="Gallup"] <- "13114"
RP500DistGRT$Location[RP500DistGRT$Location=="Valencia County"] <- "14014"
RP500DistGRT$Location[RP500DistGRT$Location=="Belen"] <- "14129"
RP500DistGRT$Location[RP500DistGRT$Location=="Los Lunas, Village of"] <- "14316"
RP500DistGRT$Location[RP500DistGRT$Location=="Peralta"] <- "14412"
RP500DistGRT$Location[RP500DistGRT$Location=="Bosque Farms"] <- "14505"
RP500DistGRT$Location[RP500DistGRT$Location=="Otero County"] <- "15015"
RP500DistGRT$Location[RP500DistGRT$Location=="Alamogordo"] <- "15116"
RP500DistGRT$Location[RP500DistGRT$Location=="Cloudcroft"] <- "15213"
RP500DistGRT$Location[RP500DistGRT$Location=="Tularosa"] <- "15308"
RP500DistGRT$Location[RP500DistGRT$Location=="San Juan County"] <- "16016"
RP500DistGRT$Location[RP500DistGRT$Location=="Farmington"] <- "16121"
RP500DistGRT$Location[RP500DistGRT$Location=="Aztec"] <- "16218"
RP500DistGRT$Location[RP500DistGRT$Location=="Bloomfield"] <- "16312"
RP500DistGRT$Location[RP500DistGRT$Location=="Rio Arriba County"] <- "17017"
RP500DistGRT$Location[RP500DistGRT$Location=="Chama"] <- "17118"
RP500DistGRT$Location[RP500DistGRT$Location=="Espanola"] <- "17215"
RP500DistGRT$Location[RP500DistGRT$Location=="Santa Clara Pueblo"] <- "17904"
RP500DistGRT$Location[RP500DistGRT$Location=="Union County"] <- "18018"
RP500DistGRT$Location[RP500DistGRT$Location=="Clayton"] <- "18128"
RP500DistGRT$Location[RP500DistGRT$Location=="Des Moines"] <- "18224"
RP500DistGRT$Location[RP500DistGRT$Location=="Folsom"] <- "18411"
RP500DistGRT$Location[RP500DistGRT$Location=="Luna County"] <- "19019"
RP500DistGRT$Location[RP500DistGRT$Location=="Deming"] <- "19113"
RP500DistGRT$Location[RP500DistGRT$Location=="Columbus"] <- "19212"
RP500DistGRT$Location[RP500DistGRT$Location=="Taos County"] <- "20020"
RP500DistGRT$Location[RP500DistGRT$Location=="Taos"] <- "20126"
RP500DistGRT$Location[RP500DistGRT$Location=="Questa"] <- "20222"
RP500DistGRT$Location[RP500DistGRT$Location=="Red River"] <- "20317"
RP500DistGRT$Location[RP500DistGRT$Location=="Taos Ski Valley, Village of"] <- "20414"
RP500DistGRT$Location[RP500DistGRT$Location=="Sierra County"] <- "21021"
RP500DistGRT$Location[RP500DistGRT$Location=="T or C"] <- "21124"
RP500DistGRT$Location[RP500DistGRT$Location=="Williamsburg"] <- "21220"
RP500DistGRT$Location[RP500DistGRT$Location=="Elephant Butte, City of"] <- "21319"
RP500DistGRT$Location[RP500DistGRT$Location=="Torrance County"] <- "22022"
RP500DistGRT$Location[RP500DistGRT$Location=="Mountainair"] <- "22127"
RP500DistGRT$Location[RP500DistGRT$Location=="Moriarty"] <- "22223"
RP500DistGRT$Location[RP500DistGRT$Location=="Willard"] <- "22314"
RP500DistGRT$Location[RP500DistGRT$Location=="Encino"] <- "22410"
RP500DistGRT$Location[RP500DistGRT$Location=="Estancia"] <- "22503"
RP500DistGRT$Location[RP500DistGRT$Location=="Hidalgo County"] <- "23023"
RP500DistGRT$Location[RP500DistGRT$Location=="Lordsburg"] <- "23110"
RP500DistGRT$Location[RP500DistGRT$Location=="Virden, Village of"] <- "23209"
RP500DistGRT$Location[RP500DistGRT$Location=="Guadalupe County"] <- "24024"
RP500DistGRT$Location[RP500DistGRT$Location=="Santa Rosa"] <- "24108"
RP500DistGRT$Location[RP500DistGRT$Location=="Vaughn"] <- "24207"
RP500DistGRT$Location[RP500DistGRT$Location=="Socorro County"] <- "25025"
RP500DistGRT$Location[RP500DistGRT$Location=="Socorro, City of"] <- "25125"
RP500DistGRT$Location[RP500DistGRT$Location=="Magdalena"] <- "25221"
RP500DistGRT$Location[RP500DistGRT$Location=="Lincoln County"] <- "26026"
RP500DistGRT$Location[RP500DistGRT$Location=="Ruidoso, Village of"] <- "26112"
RP500DistGRT$Location[RP500DistGRT$Location=="Capitan"] <- "26211"
RP500DistGRT$Location[RP500DistGRT$Location=="Carrizozo"] <- "26307"
RP500DistGRT$Location[RP500DistGRT$Location=="Corona"] <- "26406"
RP500DistGRT$Location[RP500DistGRT$Location=="Ruidoso Downs"] <- "26501"
RP500DistGRT$Location[RP500DistGRT$Location=="De Baca County"] <- "27027"
RP500DistGRT$Location[RP500DistGRT$Location=="Ft Sumner"] <- "27104"
RP500DistGRT$Location[RP500DistGRT$Location=="Catron County"] <- "28028"
RP500DistGRT$Location[RP500DistGRT$Location=="Reserve"] <- "28130"
RP500DistGRT$Location[RP500DistGRT$Location=="Sandoval County"] <- "29029"
RP500DistGRT$Location[RP500DistGRT$Location=="Bernalillo"] <- "29120"
RP500DistGRT$Location[RP500DistGRT$Location=="Jemez Springs"] <- "29217"
RP500DistGRT$Location[RP500DistGRT$Location=="Cuba"] <- "29311"
RP500DistGRT$Location[RP500DistGRT$Location=="San Ysidro"] <- "29409"
RP500DistGRT$Location[RP500DistGRT$Location=="Corrales"] <- "29504"
RP500DistGRT$Location[RP500DistGRT$Location=="Rio Rancho"] <- "29524"
RP500DistGRT$Location[RP500DistGRT$Location=="Sandia, Pueblo of"] <- "29912"
RP500DistGRT$Location[RP500DistGRT$Location=="Jicarilla Apache Nation"] <- "29932"
RP500DistGRT$Location[RP500DistGRT$Location=="Santa Ana Pueblo"] <- "29952"
RP500DistGRT$Location[RP500DistGRT$Location=="Cochiti Pueblo"] <- "29972"
RP500DistGRT$Location[RP500DistGRT$Location=="Santo Domingo Pueblo"] <- "29974"
RP500DistGRT$Location[RP500DistGRT$Location=="Mora County"] <- "30030"
RP500DistGRT$Location[RP500DistGRT$Location=="Wagon Mound"] <- "30115"
RP500DistGRT$Location[RP500DistGRT$Location=="Harding County"] <- "31031"
RP500DistGRT$Location[RP500DistGRT$Location=="Roy, Village Of"] <- "31109"
RP500DistGRT$Location[RP500DistGRT$Location=="Mosquero"] <- "31208"
RP500DistGRT$Location[RP500DistGRT$Location=="Los Alamos"] <- "32032"
RP500DistGRT$Location[RP500DistGRT$Location=="Cibola County"] <- "33033"
RP500DistGRT$Location[RP500DistGRT$Location=="Milan"] <- "33131"
RP500DistGRT$Location[RP500DistGRT$Location=="Grants"] <- "33227"
RP500DistGRT$Location[RP500DistGRT$Location=="Laguna, Pueblo of"] <- "33902"
RP500DistGRT$Location[RP500DistGRT$Location=="State Park & Rec Area Capital"] <- "CRSEMN"
RP500DistGRT$Location[RP500DistGRT$Location=="Office of Cultural Affairs"] <- "CRSOCA"
RP500DistGRT$Location[RP500DistGRT$Location=="NMFA Public Project Revolving Fund"] <- "CRSPPR"
RP500DistGRT$Location[RP500DistGRT$Location=="NMFA Public Project Revolving"] <- "CRSPPR"
RP500DistGRT$Location[RP500DistGRT$Location=="NM Youth Conservation Corp"] <- "CRSYCC"
RP500DistGRT$Location[RP500DistGRT$Location=="Leased Vehicle - Infrastructure"] <- "S444"
RP500DistGRT$Location[RP500DistGRT$Location=="Leased Vehicle -Infrastructure"] <- "S444"
RP500DistGRT$Location[RP500DistGRT$Location=="Leased Vehicle - County Road"] <- "S444C"
RP500DistGRT$Location[RP500DistGRT$Location=="General Fund - Gross Receipts - CRS"] <- "SGRT"
RP500DistGRT$Location[RP500DistGRT$Location=="General Fund-Gross Receipt-CRS"] <- "SGRT"
RP500DistGRT$Location[RP500DistGRT$Location=="General Fund-Gross Receipt-GRT"] <- "SGRT"
RP500DistGRT$Location[RP500DistGRT$Location=="County Supported Medicaid Fund"] <- "SMEDIC"
RP500DistGRT$Location[RP500DistGRT$Location=="Grand Total:"] <- "ZZZZZ"
RP500DistGRT$Location[RP500DistGRT$Location=="SF Indian School"] <- "01907"
RP500DistGRT$Location[RP500DistGRT$Location=="San Ildefonso Pueblo"] <- "01975"
RP500DistGRT$Location[RP500DistGRT$Location=="Anthony, City of"] <- "07507"
RP500DistGRT$Location[RP500DistGRT$Location=="Kirtland"] <- "16323"
RP500DistGRT$Location[RP500DistGRT$Location=="Ohkay Owingeh Pueblo"] <- "17942"
#RP500DistGRT$Location[RP500DistGRT$Location=="Ohkay Owingeh Puebo"] <- "17942"
RP500DistGRT$Location[RP500DistGRT$Location=="Taos Pueblo"] <- "20913"
RP500DistGRT$Location[RP500DistGRT$Location=="Picuris Pueblo"] <- "20918"
RP500DistGRT$Location[RP500DistGRT$Location=="Grenville"] <- "18315"
RP500DistGRT$Location[RP500DistGRT$Location=="Tesuque Pueblo"] <- "01953"
RP500DistGRT$Location[RP500DistGRT$Location=="19 Pueblos District"] <- "02905"
RP500DistGRT$Location[RP500DistGRT$Location=="Rio Communities"] <- "14037"
RP500DistGRT$Location[RP500DistGRT$Location=="Acoma Pueblo"] <- "33909"
RP500DistGRT$Location[RP500DistGRT$Location=="AIS Property/Nineteen Pueblo"] <- "02905"
RP500DistGRT$Location[RP500DistGRT$Location=="Zuni Pueblo"] <- "13901"
RP500DistGRT$Location[RP500DistGRT$Location=="Jemez Pueblo"] <- "29942"
RP500DistGRT$Location[RP500DistGRT$Location=="Zia Pueblo"] <- "29982"
RP500DistGRT$Location[RP500DistGRT$Location=="Village at Rio Rancho TIDD"] <- "GRTTM648"
RP500DistGRT$Location[RP500DistGRT$Location=="Los Diamantes TIDD"] <- "GRTTM530"
RP500DistGRT$Location[RP500DistGRT$Location=="Taos Ski Valley TIDD"] <- "GRTTM430"
RP500DistGRT$Location[RP500DistGRT$Location=="Lower Petroglyphs TIDD"] <- "GRTTM420"
RP500DistGRT$Location[RP500DistGRT$Location=="Las Cruces TIDD"] <- "GRTTM132"
RP500DistGRT$Location[RP500DistGRT$Location=="Stonegate TIDD"] <- "GRTTM038"
RP500DistGRT$Location[RP500DistGRT$Location=="Winrock Town TIDD Dist 2"] <- "GRTTM036"
RP500DistGRT$Location[RP500DistGRT$Location=="Winrock Town TIDD"] <- "GRTTM035"
RP500DistGRT$Location[RP500DistGRT$Location=="South Campus TIDD"] <- "GRTTM024"
RP500DistGRT$Location[RP500DistGRT$Location=="Santolina TIDD District 18"] <- "GRTTC638"
RP500DistGRT$Location[RP500DistGRT$Location=="Santolina TIDD District 12"] <- "GRTTC632"
RP500DistGRT$Location[RP500DistGRT$Location=="Santolina TIDD District 10"] <- "GRTTC630"
RP500DistGRT$Location[RP500DistGRT$Location=="Santolina TIDD District 7"] <- "GRTTC627"
RP500DistGRT$Location[RP500DistGRT$Location=="Santolina TIDD District 6"] <- "GRTTC626"
RP500DistGRT$Location[RP500DistGRT$Location=="Santolina TIDD District 5"] <- "GRTTC625"
RP500DistGRT$Location[RP500DistGRT$Location=="Santolina TIDD District 4"] <- "GRTTC624"
RP500DistGRT$Location[RP500DistGRT$Location=="Santolina TIDD District 1"] <- "GRTTC621"
RP500DistGRT$Location[RP500DistGRT$Location=="Santolina TIDD District 14"] <- "GRTTC634"
RP500DistGRT$Location[RP500DistGRT$Location=="Santolina TIDD District 2"] <- "GRTTC622"
RP500DistGRT$Location[RP500DistGRT$Location=="Santolina TIDD District 8"] <- "GRTTC628"
RP500DistGRT$Location[RP500DistGRT$Location=="Santolina TIDD District 17"] <- "GRTTC637"
RP500DistGRT$Location[RP500DistGRT$Location=="Santolina TIDD District 20"] <- "GRTTC640"
RP500DistGRT$Location[RP500DistGRT$Location=="Santolina TIDD District 11"] <- "GRTTC631"
RP500DistGRT$Location[RP500DistGRT$Location=="Santolina TIDD District 3"] <- "GRTTC623"
RP500DistGRT$Location[RP500DistGRT$Location=="Santolina TIDD District 19"] <- "02639"
RP500DistGRT$Location[RP500DistGRT$Location=="Santolina TIDD District 15"] <- "02635"
RP500DistGRT$Location[RP500DistGRT$Location=="Santolina TIDD District 16"] <- "02636"
RP500DistGRT$Location[RP500DistGRT$Location=="Santolina TIDD District 9"] <- "02629"
RP500DistGRT$Location[RP500DistGRT$Location=="Santolina TIDD District 13"] <- "02633"
RP500DistGRT$Location[RP500DistGRT$Location=="Quorum Uptown TIDD"] <- "02034"
RP500DistGRT$Location[RP500DistGRT$Location=="Valley Water and Sanitation Di"] <- "16322"

# Delete rows with unnecessary information
RP500DistGRT <- subset(RP500DistGRT, `Option`!="Option" & `Option`!="Business Activity Month:", )
# Delete rows with "N/A"
RP500DistGRT <- RP500DistGRT[!is.na(RP500DistGRT$Tax), ]

# Reorder columns
RP500DistGRT <- RP500DistGRT[c("Activity", "County", "Location", "Option", "Tax", "Penalty", "Interest", "Admin", "Contracts", "Payback", "Disbursement")]

# Format numbers to Numeric
RP500DistGRT$Tax <- as.numeric(RP500DistGRT$Tax)
RP500DistGRT$Penalty <- as.numeric(RP500DistGRT$Penalty)
RP500DistGRT$Interest <- as.numeric(RP500DistGRT$Interest)
RP500DistGRT$Admin <- as.numeric(RP500DistGRT$Admin)
RP500DistGRT$Contracts <- as.numeric(RP500DistGRT$Contracts)
RP500DistGRT$Payback <- as.numeric(RP500DistGRT$Payback)
RP500DistGRT$Disbursement <- as.numeric(RP500DistGRT$Disbursement)

write_xlsx(RP500DistGRT,"//trdecomsrv/H/Alphie/Temp/R temp/RP500DistGRT.xlsx")

# Save Dataset
save(RP500DistGRT, file = "//trdecomsrv/H/CRS Reports/R Backup Databases/RP500 GRT Distributions/RP500DistGRT_Alphie.RData")

## Export to excel
write_xlsx(RP500DistGRT,"//trdecomsrv/H/CRS Reports/R Backup Databases/RP500 GRT Distributions/RP500DistGRT_Alphie.xlsx")  

#######################
# Remove the rows with totals
# I'm removing these because they are not in the original SAS database
# However, I think they are valuable
###########################
RP500DistGRT <- subset(RP500DistGRT, `Option`!="Total Food Deductions:" & `Option`!="Total"
                       & `Option` != "Total Medical Deductions:"
                       & `Option` != "Gross GRT Distribution:"
                       & `Option` != "Food Distributions:"
                       & `Option` != "Medical Distributions"
                       & `Option` != "Municipal Equivalent Distribution:"
                       & `Option` != "Municipal Eq Contracts:"
                       & `Option` != "Total Administrative Fees:"
                       & `Option` != "Total Contracts:"
                       & `Option` != "Total Paybacks:"
                       & `Option` != "Total Distributed:", )


#save to STATA
require(foreign)
write.dta(RP500DistGRT, "//trdecomsrv/H/CRS Reports/STATA Databases/RP500DistGRT_Alphie.dta")
# Save to SAS
write.foreign(df = RP500DistGRT,
              datafile = '//trdecomsrv/H/CRS Reports/R Backup Databases/RP500 GRT Distributions/RP500Comb.RData',
              codefile = '//trdecomsrv/H/CRS Reports/R Backup Databases/RP500 GRT Distributions/RP500Comb_Alphie.sas',
              package = 'SAS')


