###################################################################################################################
# This script will generate the RP500 Distributions Compensating Tax Database
#
############################################################################

# Load libraries
libraries <- c("readxl", "writexl", "dplyr", "tidyr", "foreign", "haven", "lubridate", "sas7bdat")
lapply(libraries, library, character.only = TRUE)

#read and combine files
files <- list.files(
  path = "//trdecomsrv/H/CRS Reports/RP-500 Compensating Tax/Distribution/",
  pattern = "\\.xls$|\\.xlsx$",
  full.names = TRUE
)

#create R combined files
RP500DistComp <- sapply(files, read_excel, simplify = FALSE) %>% bind_rows(.id = "id")
#delete unnecessary columns
RP500DistComp$id <- NULL
RP500DistComp$...2 <- NULL
RP500DistComp$...10 <- NULL

#rename the columns
names(RP500DistComp)[names(RP500DistComp) == "...3"] <- "Tax"
names(RP500DistComp)[names(RP500DistComp) == "...4"] <- "Penalty"
names(RP500DistComp)[names(RP500DistComp) == "...5"] <- "Interest"
names(RP500DistComp)[names(RP500DistComp) == "...6"] <- "Admin"
names(RP500DistComp)[names(RP500DistComp) == "...7"] <- "Contracts"
names(RP500DistComp)[names(RP500DistComp) == "...8"] <- "Payback"
names(RP500DistComp)[names(RP500DistComp) == "...9"] <- "Unrecoverable"
names(RP500DistComp)[names(RP500DistComp) == "...11"] <- "Disbursement"

# Duplicate the first column
RP500DistComp$Option <- RP500DistComp$`New Mexico     Taxation and`
# Now delete the first column
RP500DistComp$`New Mexico     Taxation and` <- NULL

# Reorder columns
RP500DistComp <- RP500DistComp[c("Option", "Tax", "Penalty", "Interest", "Admin", "Contracts", "Payback", "Unrecoverable", "Disbursement")]


today <- Sys.Date()
start_date <- today %m-% months(189)
end_date <- start_date + years(30)  # Add 30 years to the current date

monthly_dates <- seq(start_date, end_date, by = "month")
formatted_dates_5 <- format(monthly_dates, "%B-%Y")
formatted_dates_5

formatted_dates_6 <- format(monthly_dates, "%Y-%m")
formatted_dates_6

#add activity month based on accrual month
Activity <- RP500DistComp %>%
  mutate(
    Activity = case_when(
      RP500DistComp$Tax %in% formatted_dates_5 ~ as.character(RP500DistComp$Tax %in% formatted_dates_6),
    )
  )  

Activity$Activity <- Activity$Tax[Activity$Activity == FALSE]

# Delete rows with unnecessary information
RP500DistComp <- subset(Activity, `Option`!="Taxation and Revenue" & `Option`!="Local Government Distribution" 
                & `Option`!="Revenue" & `Option`!="Summary" & `Option`!="Revenue Group:"
                & `Option`!="Reporting Month:" & `Option`!="Distribution Month:"
                & `Option`!="Report Run:" & `Option`!="From:"
                & `Option`!="To:", )

#fill in missing activity dates
RP500DistComp <- fill(RP500DistComp, Activity, .direction = c("down"))

#add activity month based on accrual month
County <- RP500DistComp %>% mutate(
  County = case_when(
    RP500DistComp$Option == "Santa Fe County" ~ "Santa Fe County",
    RP500DistComp$Option == "Santa Fe, City of" ~ "Santa Fe, City of",
    RP500DistComp$Option == 	"Edgewood, Town of"	 ~ 	"Edgewood, Town of",
    RP500DistComp$Option == 	"Pueblo Of Nambe"	 ~ 	"Pueblo Of Nambe",
    RP500DistComp$Option == 	"Pojoaque Pueblo"	 ~ 	"Pojoaque Pueblo",
    RP500DistComp$Option == 	"Bernalillo County"	 ~ 	"Bernalillo County",
    RP500DistComp$Option == 	"Albuquerque, City of"	 ~ 	"Albuquerque, City of",
    RP500DistComp$Option == 	"Los Ranchosde Alb"	 ~ 	"Los Ranchosde Alb",
    RP500DistComp$Option == 	"Tijeras"	 ~ 	"Tijeras",
    RP500DistComp$Option == 	"Mesa Del Sol District"	 ~ 	"Mesa Del Sol District",
    RP500DistComp$Option == 	"Eddy County"	~	"Eddy County",
    RP500DistComp$Option == 	"Carlsbad"	 ~ 	"Carlsbad",
    RP500DistComp$Option == 	"Artesia"	 ~ 	"Artesia",
    RP500DistComp$Option == 	"Hope"	 ~ 	"Hope",
    RP500DistComp$Option == 	"Loving"	 ~ 	"Loving",
    RP500DistComp$Option == 	"Chaves County"	 ~ 	"Chaves County",
    RP500DistComp$Option == 	"Roswell"	 ~ 	"Roswell",
    RP500DistComp$Option == 	"Dexter"	 ~ 	"Dexter",
    RP500DistComp$Option == 	"Hagerman"	 ~ 	"Hagerman",
    RP500DistComp$Option == 	"Lake Arthur"	 ~ 	"Lake Arthur",
    RP500DistComp$Option == 	"Curry County"	 ~ 	"Curry County",
    RP500DistComp$Option == 	"Clovis"	 ~ 	"Clovis",
    RP500DistComp$Option == 	"Village Of Grady"	 ~ 	"Village Of Grady",
    RP500DistComp$Option == 	"Texico"	~	"Texico",
    RP500DistComp$Option == 	"Melrose"	 ~ 	"Melrose",
    RP500DistComp$Option == 	"Lea County"	 ~ 	"Lea County",
    RP500DistComp$Option == 	"Hobbs"	 ~ 	"Hobbs",
    RP500DistComp$Option == 	"Eunice"	 ~ 	"Eunice",
    RP500DistComp$Option == 	"Jal"	 ~ 	"Jal",
    RP500DistComp$Option == 	"Lovington"	 ~ 	"Lovington",
    RP500DistComp$Option == 	"Tatum"	 ~ 	"Tatum",
    RP500DistComp$Option == 	"Dona Ana County"	 ~ 	"Dona Ana County",
    RP500DistComp$Option == 	"Las Cruces, City of"	 ~ 	"Las Cruces, City of",
    RP500DistComp$Option == 	"Hatch"	 ~ 	"Hatch",
    RP500DistComp$Option == 	"Mesilla, Town of"	 ~ 	"Mesilla, Town of",
    RP500DistComp$Option == 	"Sunland Park"	 ~ 	"Sunland Park",
    RP500DistComp$Option == 	"Grant County"	~	"Grant County",
    RP500DistComp$Option == 	"Silver City"	 ~ 	"Silver City",
    RP500DistComp$Option == 	"Bayard"	 ~ 	"Bayard",
    RP500DistComp$Option == 	"Santa Clara"	 ~ 	"Santa Clara",
    RP500DistComp$Option == 	"Hurley"	 ~ 	"Hurley",
    RP500DistComp$Option == 	"Colfax County"	 ~ 	"Colfax County",
    RP500DistComp$Option == 	"Raton"	 ~ 	"Raton",
    RP500DistComp$Option == 	"Maxwell"	 ~ 	"Maxwell",
    RP500DistComp$Option == 	"Springer"	 ~ 	"Springer",
    RP500DistComp$Option == 	"Cimarron"	 ~ 	"Cimarron",
    RP500DistComp$Option == 	"Eagle Nest"	 ~ 	"Eagle Nest",
    RP500DistComp$Option == 	"Angel Fire"	 ~ 	"Angel Fire",
    RP500DistComp$Option == 	"Quay County"	 ~ 	"Quay County",
    RP500DistComp$Option == 	"Tucumcari"	~	"Tucumcari",
    RP500DistComp$Option == 	"San Jon"	 ~ 	"San Jon",
    RP500DistComp$Option == 	"Logan"	 ~ 	"Logan",
    RP500DistComp$Option == 	"House"	 ~ 	"House",
    RP500DistComp$Option == 	"Roosevelt County"	 ~ 	"Roosevelt County",
    RP500DistComp$Option == 	"Portales"	 ~ 	"Portales",
    RP500DistComp$Option == 	"Elida"	 ~ 	"Elida",
    RP500DistComp$Option == 	"Dora"	 ~ 	"Dora",
    RP500DistComp$Option == 	"Causey"	 ~ 	"Causey",
    RP500DistComp$Option == 	"Floyd"	 ~ 	"Floyd",
    RP500DistComp$Option == 	"San Miguel Co"	 ~ 	"San Miguel Co",
    RP500DistComp$Option == 	"Las Vegas"	 ~ 	"Las Vegas",
    RP500DistComp$Option == 	"Pecos"	 ~ 	"Pecos",
    RP500DistComp$Option == 	"Mckinley County"	~	"Mckinley County",
    RP500DistComp$Option == 	"Gallup"	 ~ 	"Gallup",
    RP500DistComp$Option == 	"Valencia County"	 ~ 	"Valencia County",
    RP500DistComp$Option == 	"Belen"	 ~ 	"Belen",
    RP500DistComp$Option == 	"Los Lunas, Village of"	 ~ 	"Los Lunas, Village of",
    RP500DistComp$Option == 	"Peralta"	 ~ 	"Peralta",
    RP500DistComp$Option == 	"Bosque Farms"	 ~ 	"Bosque Farms",
    RP500DistComp$Option == 	"Otero County"	 ~ 	"Otero County",
    RP500DistComp$Option == 	"Alamogordo"	 ~ 	"Alamogordo",
    RP500DistComp$Option == 	"Cloudcroft"	 ~ 	"Cloudcroft",
    RP500DistComp$Option == 	"Tularosa"	 ~ 	"Tularosa",
    RP500DistComp$Option == 	"San Juan County"	 ~ 	"San Juan County",
    RP500DistComp$Option == 	"Farmington"	 ~ 	"Farmington",
    RP500DistComp$Option == 	"Aztec"	~	"Aztec",
    RP500DistComp$Option == 	"Bloomfield"	 ~ 	"Bloomfield",
    RP500DistComp$Option == 	"Rio Arriba County"	 ~ 	"Rio Arriba County",
    RP500DistComp$Option == 	"Chama"	 ~ 	"Chama",
    RP500DistComp$Option == 	"Espanola"	 ~ 	"Espanola",
    RP500DistComp$Option == 	"Santa Clara Pueblo"	 ~ 	"Santa Clara Pueblo",
    RP500DistComp$Option == 	"Union County"	 ~ 	"Union County",
    RP500DistComp$Option == 	"Clayton"	 ~ 	"Clayton",
    RP500DistComp$Option == 	"Des Moines"	 ~ 	"Des Moines",
    RP500DistComp$Option == 	"Folsom"	 ~ 	"Folsom",
    RP500DistComp$Option == 	"Luna County"	 ~ 	"Luna County",
    RP500DistComp$Option == 	"Deming"	 ~ 	"Deming",
    RP500DistComp$Option == 	"Columbus"	 ~ 	"Columbus",
    RP500DistComp$Option == 	"Taos County"	~	"Taos County",
    RP500DistComp$Option == 	"Taos"	 ~ 	"Taos",
    RP500DistComp$Option == 	"Questa"	 ~ 	"Questa",
    RP500DistComp$Option == 	"Red River"	 ~ 	"Red River",
    RP500DistComp$Option == 	"Taos Ski Valley, Village of"	 ~ 	"Taos Ski Valley, Village of",
    RP500DistComp$Option == 	"Sierra County"	 ~ 	"Sierra County",
    RP500DistComp$Option == 	"T or C"	 ~ 	"T or C",
    RP500DistComp$Option == 	"Williamsburg"	 ~ 	"Williamsburg",
    RP500DistComp$Option == 	"Elephant Butte, City of"	 ~ 	"Elephant Butte, City of",
    RP500DistComp$Option == 	"Torrance County"	 ~ 	"Torrance County",
    RP500DistComp$Option == 	"Mountainair"	 ~ 	"Mountainair",
    RP500DistComp$Option == 	"Moriarty"	 ~ 	"Moriarty",
    RP500DistComp$Option == 	"Willard"	 ~ 	"Willard",
    RP500DistComp$Option == 	"Encino"	~	"Encino",
    RP500DistComp$Option == 	"Estancia"	 ~ 	"Estancia",
    RP500DistComp$Option == 	"Hidalgo County"	 ~ 	"Hidalgo County",
    RP500DistComp$Option == 	"Lordsburg"	 ~ 	"Lordsburg",
    RP500DistComp$Option == 	"Virden, Village of"	 ~ 	"Virden, Village of",
    RP500DistComp$Option == 	"Guadalupe County"	 ~ 	"Guadalupe County",
    RP500DistComp$Option == 	"Santa Rosa"	 ~ 	"Santa Rosa",
    RP500DistComp$Option == 	"Vaughn"	 ~ 	"Vaughn",
    RP500DistComp$Option == 	"Socorro County"	 ~ 	"Socorro County",
    RP500DistComp$Option == 	"Socorro, City of"	 ~ 	"Socorro, City of",
    RP500DistComp$Option == 	"Magdalena"	 ~ 	"Magdalena",
    RP500DistComp$Option == 	"Lincoln County"	 ~ 	"Lincoln County",
    RP500DistComp$Option == 	"Ruidoso, Village of"	 ~ 	"Ruidoso, Village of",
    RP500DistComp$Option == 	"Capitan"	~	"Capitan",
    RP500DistComp$Option == 	"Carrizozo"	 ~ 	"Carrizozo",
    RP500DistComp$Option == 	"Corona"	 ~ 	"Corona",
    RP500DistComp$Option == 	"Ruidoso Downs"	 ~ 	"Ruidoso Downs",
    RP500DistComp$Option == 	"De Baca County"	 ~ 	"De Baca County",
    RP500DistComp$Option == 	"Ft Sumner"	 ~ 	"Ft Sumner",
    RP500DistComp$Option == 	"Catron County"	 ~ 	"Catron County",
    RP500DistComp$Option == 	"Reserve"	 ~ 	"Reserve",
    RP500DistComp$Option == 	"Sandoval County"	 ~ 	"Sandoval County",
    RP500DistComp$Option == 	"Bernalillo"	 ~ 	"Bernalillo",
    RP500DistComp$Option == 	"Jemez Springs"	 ~ 	"Jemez Springs",
    RP500DistComp$Option == 	"Cuba"	 ~ 	"Cuba",
    RP500DistComp$Option == 	"San Ysidro"	 ~ 	"San Ysidro",
    RP500DistComp$Option == 	"Corrales"	~	"Corrales",
    RP500DistComp$Option == 	"Rio Rancho"	 ~ 	"Rio Rancho",
    RP500DistComp$Option == 	"Sandia, Pueblo of"	 ~ 	"Sandia, Pueblo of",
    RP500DistComp$Option == 	"Jicarilla Apache Nation"	 ~ 	"Jicarilla Apache Nation",
    RP500DistComp$Option == 	"Santa Ana Pueblo"	 ~ 	"Santa Ana Pueblo",
    RP500DistComp$Option == 	"Cochiti Pueblo"	 ~ 	"Cochiti Pueblo",
    RP500DistComp$Option == 	"Santo Domingo Pueblo"	 ~ 	"Santo Domingo Pueblo",
    RP500DistComp$Option == 	"Mora County"	 ~ 	"Mora County",
    RP500DistComp$Option == 	"Wagon Mound"	 ~ 	"Wagon Mound",
    RP500DistComp$Option == 	"Harding County"	 ~ 	"Harding County",
    RP500DistComp$Option == 	"Roy, Village Of"	 ~ 	"Roy, Village Of",
    RP500DistComp$Option == 	"Mosquero"	 ~ 	"Mosquero",
    RP500DistComp$Option == 	"Los Alamos"	 ~ 	"Los Alamos",
    RP500DistComp$Option == 	"Cibola County"	~	"Cibola County",
    RP500DistComp$Option == 	"Milan"	 ~ 	"Milan",
    RP500DistComp$Option == 	"Grants"	 ~ 	"Grants",
    RP500DistComp$Option == 	"Laguna, Pueblo of"	 ~ 	"Laguna, Pueblo of",
    RP500DistComp$Option == 	"State Park & Rec Area Capital"	 ~ 	"State Park & Rec Area Capital",
    RP500DistComp$Option == 	"Office of Cultural Affairs"	 ~ 	"Office of Cultural Affairs",
    RP500DistComp$Option == 	"NMFA Public Project Revolving Fund"	 ~ 	"NMFA Public Project Revolving Fund",
    RP500DistComp$Option == 	"NMFA Public Project Revolving"	 ~ 	"NMFA Public Project Revolving",
    RP500DistComp$Option == 	"NM Youth Conservation Corp"	 ~ 	"NM Youth Conservation Corp",
    RP500DistComp$Option == 	"Leased Vehicle - Infrastructure"	 ~ 	"Leased Vehicle - Infrastructure",
    RP500DistComp$Option == 	"Leased Vehicle -Infrastructure"	 ~ 	"Leased Vehicle -Infrastructure",
    RP500DistComp$Option == 	"Leased Vehicle - County Road"	 ~ 	"Leased Vehicle - County Road",
    RP500DistComp$Option == 	"General Fund-Compensating"	 ~ 	"General Fund-Compensating",
    RP500DistComp$Option == 	"County Supported Medicaid Fund"	 ~ 	"County Supported Medicaid Fund",
    RP500DistComp$Option == 	"Grand Total"	 ~ 	"Grand Total",
    RP500DistComp$Option == 	"SF Indian School"	 ~ 	"SF Indian School",
    RP500DistComp$Option == 	"San Ildefonso Pueblo"	 ~ 	"San Ildefonso Pueblo",
    RP500DistComp$Option == 	"Anthony, City of"	 ~ 	"Anthony, City of",
    RP500DistComp$Option == 	"Kirtland"	 ~ 	"Kirtland",
    RP500DistComp$Option == 	"Ohkay Owingeh Pueblo"	 ~ 	"Ohkay Owingeh Pueblo",
    RP500DistComp$Option == 	"Ohkay Owingeh Puebo"	 ~ 	"Ohkay Owingeh Puebo",
    RP500DistComp$Option == 	"Taos Pueblo"	 ~ 	"Taos Pueblo",
    RP500DistComp$Option == 	"Picuris Pueblo"	 ~ 	"Picuris Pueblo",
    RP500DistComp$Option == 	"Grenville"	 ~ 	"Grenville",
    RP500DistComp$Option == 	"Tesuque Pueblo"	 ~ 	"Tesuque Pueblo",
    RP500DistComp$Option == 	"19 Pueblos District"	~	"19 Pueblos District",
    RP500DistComp$Option == 	"Rio Communities"	 ~ 	"Rio Communities",
    RP500DistComp$Option == 	"Acoma Pueblo"	 ~ 	"Acoma Pueblo",
    RP500DistComp$Option == 	"AIS Property/Nineteen Pueblo"	 ~ 	"AIS Property/Nineteen Pueblo",
    RP500DistComp$Option == 	"Zuni Pueblo"	 ~ 	"Zuni Pueblo",
    
  )
)

#fill in missing location names dates
RP500DistComp <- fill(County, County, .direction = c("down"))

# Duplicate County column
RP500DistComp$Location <- RP500DistComp$County
#replace location name with location code
RP500DistComp$Location[RP500DistComp$Location=="Santa Fe County"] <- "01001"
RP500DistComp$Location[RP500DistComp$Location=="Santa Fe, City of"] <- "01123"
RP500DistComp$Location[RP500DistComp$Location=="Edgewood, Town of"] <- "01320"
RP500DistComp$Location[RP500DistComp$Location=="Pueblo Of Nambe"] <- "01952"
RP500DistComp$Location[RP500DistComp$Location=="Pojoaque Pueblo"] <- "01962"
RP500DistComp$Location[RP500DistComp$Location=="Bernalillo County"] <- "02002"
RP500DistComp$Location[RP500DistComp$Location=="Albuquerque, City of"] <- "02100"
RP500DistComp$Location[RP500DistComp$Location=="Los Ranchosde Alb"] <- "02200"
RP500DistComp$Location[RP500DistComp$Location=="Tijeras"] <- "02318"
RP500DistComp$Location[RP500DistComp$Location=="Mesa Del Sol District"] <- "02606"
RP500DistComp$Location[RP500DistComp$Location=="Eddy County"] <- "03003"
RP500DistComp$Location[RP500DistComp$Location=="Carlsbad"] <- "03106"
RP500DistComp$Location[RP500DistComp$Location=="Artesia"] <- "03205"
RP500DistComp$Location[RP500DistComp$Location=="Hope"] <- "03304"
RP500DistComp$Location[RP500DistComp$Location=="Loving"] <- "03403"
RP500DistComp$Location[RP500DistComp$Location=="Chaves County"] <- "04004"
RP500DistComp$Location[RP500DistComp$Location=="Roswell"] <- "04101"
RP500DistComp$Location[RP500DistComp$Location=="Dexter"] <- "04201"
RP500DistComp$Location[RP500DistComp$Location=="Hagerman"] <- "04300"
RP500DistComp$Location[RP500DistComp$Location=="Lake Arthur"] <- "04400"
RP500DistComp$Location[RP500DistComp$Location=="Curry County"] <- "05005"
RP500DistComp$Location[RP500DistComp$Location=="Clovis"] <- "05103"
RP500DistComp$Location[RP500DistComp$Location=="Village Of Grady"] <- "015203"
RP500DistComp$Location[RP500DistComp$Location=="Texico"] <- "05302"
RP500DistComp$Location[RP500DistComp$Location=="Melrose"] <- "05402"
RP500DistComp$Location[RP500DistComp$Location=="Lea County"] <- "06006"
RP500DistComp$Location[RP500DistComp$Location=="Hobbs"] <- "06111"
RP500DistComp$Location[RP500DistComp$Location=="Eunice"] <- "06210"
RP500DistComp$Location[RP500DistComp$Location=="Jal"] <- "06306"
RP500DistComp$Location[RP500DistComp$Location=="Lovington"] <- "06405"
RP500DistComp$Location[RP500DistComp$Location=="Tatum"] <- "06500"
RP500DistComp$Location[RP500DistComp$Location=="Dona Ana County"] <- "07007"
RP500DistComp$Location[RP500DistComp$Location=="Las Cruces, City of"] <- "07105"
RP500DistComp$Location[RP500DistComp$Location=="Hatch"] <- "07204"
RP500DistComp$Location[RP500DistComp$Location=="Mesilla, Town of"] <- "07303"
RP500DistComp$Location[RP500DistComp$Location=="Sunland Park"] <- "07416"
RP500DistComp$Location[RP500DistComp$Location=="Grant County"] <- "08008"
RP500DistComp$Location[RP500DistComp$Location=="Silver City"] <- "08107"
RP500DistComp$Location[RP500DistComp$Location=="Bayard"] <- "08206"
RP500DistComp$Location[RP500DistComp$Location=="Santa Clara"] <- "08305"
RP500DistComp$Location[RP500DistComp$Location=="Hurley"] <- "08404"
RP500DistComp$Location[RP500DistComp$Location=="Colfax County"] <- "09009"
RP500DistComp$Location[RP500DistComp$Location=="Raton"] <- "09102"
RP500DistComp$Location[RP500DistComp$Location=="Maxwell"] <- "09202"
RP500DistComp$Location[RP500DistComp$Location=="Springer"] <- "09301"
RP500DistComp$Location[RP500DistComp$Location=="Cimarron"] <- "09401"
RP500DistComp$Location[RP500DistComp$Location=="Eagle Nest"] <- "09509"
RP500DistComp$Location[RP500DistComp$Location=="Angel Fire"] <- "09600"
RP500DistComp$Location[RP500DistComp$Location=="Quay County"] <- "10010"
RP500DistComp$Location[RP500DistComp$Location=="Tucumcari"] <- "10117"
RP500DistComp$Location[RP500DistComp$Location=="San Jon"] <- "10214"
RP500DistComp$Location[RP500DistComp$Location=="Logan"] <- "10309"
RP500DistComp$Location[RP500DistComp$Location=="House"] <- "10407"
RP500DistComp$Location[RP500DistComp$Location=="Roosevelt County"] <- "11011"
RP500DistComp$Location[RP500DistComp$Location=="Portales"] <- "11119"
RP500DistComp$Location[RP500DistComp$Location=="Elida"] <- "11216"
RP500DistComp$Location[RP500DistComp$Location=="Dora"] <- "11310"
RP500DistComp$Location[RP500DistComp$Location=="Causey"] <- "11408"
RP500DistComp$Location[RP500DistComp$Location=="Floyd"] <- "11502"
RP500DistComp$Location[RP500DistComp$Location=="San Miguel Co"] <- "12012"
RP500DistComp$Location[RP500DistComp$Location=="Las Vegas"] <- "12122"
RP500DistComp$Location[RP500DistComp$Location=="Pecos"] <- "12313"
RP500DistComp$Location[RP500DistComp$Location=="Mckinley County"] <- "13013"
RP500DistComp$Location[RP500DistComp$Location=="Gallup"] <- "13114"
RP500DistComp$Location[RP500DistComp$Location=="Valencia County"] <- "14014"
RP500DistComp$Location[RP500DistComp$Location=="Belen"] <- "14129"
RP500DistComp$Location[RP500DistComp$Location=="Los Lunas, Village of"] <- "14316"
RP500DistComp$Location[RP500DistComp$Location=="Peralta"] <- "14412"
RP500DistComp$Location[RP500DistComp$Location=="Bosque Farms"] <- "14505"
RP500DistComp$Location[RP500DistComp$Location=="Otero County"] <- "15015"
RP500DistComp$Location[RP500DistComp$Location=="Alamogordo"] <- "15116"
RP500DistComp$Location[RP500DistComp$Location=="Cloudcroft"] <- "15213"
RP500DistComp$Location[RP500DistComp$Location=="Tularosa"] <- "15308"
RP500DistComp$Location[RP500DistComp$Location=="San Juan County"] <- "16016"
RP500DistComp$Location[RP500DistComp$Location=="Farmington"] <- "16121"
RP500DistComp$Location[RP500DistComp$Location=="Aztec"] <- "16218"
RP500DistComp$Location[RP500DistComp$Location=="Bloomfield"] <- "16312"
RP500DistComp$Location[RP500DistComp$Location=="Rio Arriba County"] <- "17017"
RP500DistComp$Location[RP500DistComp$Location=="Chama"] <- "17118"
RP500DistComp$Location[RP500DistComp$Location=="Espanola"] <- "17215"
RP500DistComp$Location[RP500DistComp$Location=="Santa Clara Pueblo"] <- "17904"
RP500DistComp$Location[RP500DistComp$Location=="Union County"] <- "18018"
RP500DistComp$Location[RP500DistComp$Location=="Clayton"] <- "18128"
RP500DistComp$Location[RP500DistComp$Location=="Des Moines"] <- "18224"
RP500DistComp$Location[RP500DistComp$Location=="Folsom"] <- "18411"
RP500DistComp$Location[RP500DistComp$Location=="Luna County"] <- "19019"
RP500DistComp$Location[RP500DistComp$Location=="Deming"] <- "19113"
RP500DistComp$Location[RP500DistComp$Location=="Columbus"] <- "19212"
RP500DistComp$Location[RP500DistComp$Location=="Taos County"] <- "20020"
RP500DistComp$Location[RP500DistComp$Location=="Taos"] <- "20126"
RP500DistComp$Location[RP500DistComp$Location=="Questa"] <- "20222"
RP500DistComp$Location[RP500DistComp$Location=="Red River"] <- "20317"
RP500DistComp$Location[RP500DistComp$Location=="Taos Ski Valley, Village of"] <- "20414"
RP500DistComp$Location[RP500DistComp$Location=="Sierra County"] <- "21021"
RP500DistComp$Location[RP500DistComp$Location=="T or C"] <- "21124"
RP500DistComp$Location[RP500DistComp$Location=="Williamsburg"] <- "21220"
RP500DistComp$Location[RP500DistComp$Location=="Elephant Butte, City of"] <- "21319"
RP500DistComp$Location[RP500DistComp$Location=="Torrance County"] <- "22022"
RP500DistComp$Location[RP500DistComp$Location=="Mountainair"] <- "22127"
RP500DistComp$Location[RP500DistComp$Location=="Moriarty"] <- "22223"
RP500DistComp$Location[RP500DistComp$Location=="Willard"] <- "22314"
RP500DistComp$Location[RP500DistComp$Location=="Encino"] <- "22410"
RP500DistComp$Location[RP500DistComp$Location=="Estancia"] <- "22503"
RP500DistComp$Location[RP500DistComp$Location=="Hidalgo County"] <- "23023"
RP500DistComp$Location[RP500DistComp$Location=="Lordsburg"] <- "23110"
RP500DistComp$Location[RP500DistComp$Location=="Virden, Village of"] <- "23209"
RP500DistComp$Location[RP500DistComp$Location=="Guadalupe County"] <- "24024"
RP500DistComp$Location[RP500DistComp$Location=="Santa Rosa"] <- "24108"
RP500DistComp$Location[RP500DistComp$Location=="Vaughn"] <- "24207"
RP500DistComp$Location[RP500DistComp$Location=="Socorro County"] <- "25025"
RP500DistComp$Location[RP500DistComp$Location=="Socorro, City of"] <- "25125"
RP500DistComp$Location[RP500DistComp$Location=="Magdalena"] <- "25221"
RP500DistComp$Location[RP500DistComp$Location=="Lincoln County"] <- "26026"
RP500DistComp$Location[RP500DistComp$Location=="Ruidoso, Village of"] <- "26112"
RP500DistComp$Location[RP500DistComp$Location=="Capitan"] <- "26211"
RP500DistComp$Location[RP500DistComp$Location=="Carrizozo"] <- "26307"
RP500DistComp$Location[RP500DistComp$Location=="Corona"] <- "26406"
RP500DistComp$Location[RP500DistComp$Location=="Ruidoso Downs"] <- "26501"
RP500DistComp$Location[RP500DistComp$Location=="De Baca County"] <- "27027"
RP500DistComp$Location[RP500DistComp$Location=="Ft Sumner"] <- "27104"
RP500DistComp$Location[RP500DistComp$Location=="Catron County"] <- "28028"
RP500DistComp$Location[RP500DistComp$Location=="Reserve"] <- "28130"
RP500DistComp$Location[RP500DistComp$Location=="Sandoval County"] <- "29029"
RP500DistComp$Location[RP500DistComp$Location=="Bernalillo"] <- "29120"
RP500DistComp$Location[RP500DistComp$Location=="Jemez Springs"] <- "29217"
RP500DistComp$Location[RP500DistComp$Location=="Cuba"] <- "29311"
RP500DistComp$Location[RP500DistComp$Location=="San Ysidro"] <- "29409"
RP500DistComp$Location[RP500DistComp$Location=="Corrales"] <- "29504"
RP500DistComp$Location[RP500DistComp$Location=="Rio Rancho"] <- "29524"
RP500DistComp$Location[RP500DistComp$Location=="Sandia, Pueblo of"] <- "29912"
RP500DistComp$Location[RP500DistComp$Location=="Jicarilla Apache Nation"] <- "29932"
RP500DistComp$Location[RP500DistComp$Location=="Santa Ana Pueblo"] <- "29952"
RP500DistComp$Location[RP500DistComp$Location=="Cochiti Pueblo"] <- "29972"
RP500DistComp$Location[RP500DistComp$Location=="Santo Domingo Pueblo"] <- "29974"
RP500DistComp$Location[RP500DistComp$Location=="Mora County"] <- "30030"
RP500DistComp$Location[RP500DistComp$Location=="Wagon Mound"] <- "30115"
RP500DistComp$Location[RP500DistComp$Location=="Harding County"] <- "31031"
RP500DistComp$Location[RP500DistComp$Location=="Roy, Village Of"] <- "31109"
RP500DistComp$Location[RP500DistComp$Location=="Mosquero"] <- "31208"
RP500DistComp$Location[RP500DistComp$Location=="Los Alamos"] <- "32032"
RP500DistComp$Location[RP500DistComp$Location=="Cibola County"] <- "33033"
RP500DistComp$Location[RP500DistComp$Location=="Milan"] <- "33131"
RP500DistComp$Location[RP500DistComp$Location=="Grants"] <- "33227"
RP500DistComp$Location[RP500DistComp$Location=="Laguna, Pueblo of"] <- "33902"
RP500DistComp$Location[RP500DistComp$Location=="State Park & Rec Area Capital"] <- "CRSEMN"
RP500DistComp$Location[RP500DistComp$Location=="Office of Cultural Affairs"] <- "CRSOCA"
RP500DistComp$Location[RP500DistComp$Location=="NMFA Public Project Revolving Fund"] <- "CRSPPR"
RP500DistComp$Location[RP500DistComp$Location=="NMFA Public Project Revolving"] <- "CRSPPR"
RP500DistComp$Location[RP500DistComp$Location=="NM Youth Conservation Corp"] <- "CRSYCC"
RP500DistComp$Location[RP500DistComp$Location=="Leased Vehicle - Infrastructure"] <- "S444"
RP500DistComp$Location[RP500DistComp$Location=="Leased Vehicle -Infrastructure"] <- "S444"
RP500DistComp$Location[RP500DistComp$Location=="Leased Vehicle - County Road"] <- "S444C"
RP500DistComp$Location[RP500DistComp$Location=="General Fund-Compensating"] <- "FNDSCMP"
RP500DistComp$Location[RP500DistComp$Location=="County Supported Medicaid Fund"] <- "SMEDIC"
RP500DistComp$Location[RP500DistComp$Location=="Grand Total"] <- "ZZZZZ"
RP500DistComp$Location[RP500DistComp$Location=="SF Indian School"] <- "01907"
RP500DistComp$Location[RP500DistComp$Location=="San Ildefonso Pueblo"] <- "01975"
RP500DistComp$Location[RP500DistComp$Location=="Anthony, City of"] <- "07507"
RP500DistComp$Location[RP500DistComp$Location=="Kirtland"] <- "16323"
RP500DistComp$Location[RP500DistComp$Location=="Ohkay Owingeh Pueblo"] <- "17942"
RP500DistComp$Location[RP500DistComp$Location=="Ohkay Owingeh Puebo"] <- "17942"
RP500DistComp$Location[RP500DistComp$Location=="Taos Pueblo"] <- "20913"
RP500DistComp$Location[RP500DistComp$Location=="Picuris Pueblo"] <- "20918"
RP500DistComp$Location[RP500DistComp$Location=="Grenville"] <- "18315"
RP500DistComp$Location[RP500DistComp$Location=="Tesuque Pueblo"] <- "01953"
RP500DistComp$Location[RP500DistComp$Location=="19 Pueblos District"] <- "02905"
RP500DistComp$Location[RP500DistComp$Location=="Rio Communities"] <- "14037"
RP500DistComp$Location[RP500DistComp$Location=="Acoma Pueblo"] <- "33909"
RP500DistComp$Location[RP500DistComp$Location=="AIS Property/Nineteen Pueblo"] <- "02905"
RP500DistComp$Location[RP500DistComp$Location=="Zuni Pueblo"] <- "13901"


# Delete rows with unnecessary information
RP500DistComp <- subset(RP500DistComp, `Option`!="Option" & `Option`!="Business Activity Month:", )
# Delete rows with "N/A"
RP500DistComp <- RP500DistComp[!is.na(RP500DistComp$Tax), ]

# Reorder columns
RP500DistComp <- RP500DistComp[c("Activity", "County", "Location", "Option", "Tax", "Penalty", "Interest", "Admin", "Contracts", "Payback", "Unrecoverable", "Disbursement")]

# Format numbers to Numeric
RP500DistComp$Tax <- as.numeric(RP500DistComp$Tax)
RP500DistComp$Penalty <- as.numeric(RP500DistComp$Penalty)
RP500DistComp$Interest <- as.numeric(RP500DistComp$Interest)
RP500DistComp$Admin <- as.numeric(RP500DistComp$Admin)
RP500DistComp$Contracts <- as.numeric(RP500DistComp$Contracts)
RP500DistComp$Payback <- as.numeric(RP500DistComp$Payback)
RP500DistComp$Payback <- as.numeric(RP500DistComp$Unrecoverable)
RP500DistComp$Disbursement <- as.numeric(RP500DistComp$`Disbursement`)

#################################################################################################################################################################################################
# Export datasets

# Save Dataset
save(RP500DistComp, file = "//trdecomsrv/H/Alphie/Temp/R temp/RP500COMPCombined.RData")
write_xlsx(RP500DistComp,"//trdecomsrv/H/CRS Reports/R Backup Databases/RP500 COMP Distributions/RP500COMPComb_Alphie.xlsx")
#save to STATA
require(foreign)
write.dta(RP500DistComp, "//trdecomsrv/H/CRS Reports/STATA Databases/RP500COMPComb_Alphie.dta")
# Save to SAS
write.foreign(df = RP500DistComp,
              datafile = '//trdecomsrv/H/CRS Reports/R Backup Databases/RP500 COMP Distributions/RP500COMPComb_Alphie.RData',
              codefile = '//trdecomsrv/H/CRS Reports/R Backup Databases/RP500 COMP Distributions/RP500COMPComb_Alphie.sas',
              package = 'SAS')


