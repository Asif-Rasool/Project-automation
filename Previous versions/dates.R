
library(readxl)
library(writexl)
library(dplyr)
library(foreign)
library(haven)
library(sas7bdat)
library(tidyr)
library(lubridate)



last_month <- floor_date(Sys.Date(), "month") - months(2)
last_month

formatted_dates_1 <- format(last_month, "%B-%Y")
formatted_dates_1

formatted_dates_2 <- format(last_month, "%Y-%m")
formatted_dates_2

# Example start and end dates
start_date <- as.Date("2022-01-01")
end_date <- as.Date("2023-12-31")

# Generate a sequence of dates between start and end dates
date_sequence <- seq(start_date, formatted_dates_1, by = "month")
date_sequence
