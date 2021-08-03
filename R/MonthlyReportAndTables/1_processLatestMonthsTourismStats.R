#### Preparation ####

# Clear the environment
rm(list = ls())

# Load the required libraries
#install.packages("dplyr")
library(dplyr) # Manipulating data
library(openxlsx) # Reading and editing excel formatted files

# Note where VNSO code/data is on current computer
repository <- file.path(dirname(rstudioapi::getSourceEditorContext()$path), "..", "..")
setwd(repository) # Required for file.choose() function

# Load the general R functions
source(file.path(repository, "R", "functions.R"))

# Note the secure data path
secureDataFolder <- file.path(repository, "data", "secure")

# Note the open data path
openDataFolder <- file.path(repository, "data", "open")

# Read in the raw TOURISM data from secure folder of the repository 
# Note that spaces in column names have been automatically replaced with "."s
tourismStatsFile <- file.path(secureDataFolder, "Passenger_Report_2.csv")
tourismStats <- read.csv(tourismStatsFile, header = TRUE, na.strings = c("","NA","NULL","null"))

#Read in the classifications data from open folder of the repository
CountryCodesByVisitors <- file.path(openDataFolder, "CountryCodes_VisByNat.csv")
CountryCodesVisitors <- read.csv(CountryCodesByVisitors, header = TRUE, na.strings = c("","NA","NULL","null"))

CountryCodesByResidents <- file.path(openDataFolder, "CountryCodes_ResbyNat.csv")
CountryCodesResidents <- read.csv(CountryCodesByResidents, header = TRUE, na.strings = c("","NA","NULL","null"))

## Replaces all missing values to NA, NULL, null
## TourismStats <- read.csv(tourismStatsFile, Header = TRUE, na.strings=c ("","NA","NULL", "null"))

# count the missing values
numberMissing <- apply(tourismStats, MARGIN=2,
                       FUN=function(columnValues){
                         return(sum(is.na(columnValues)))
                       })

View(numberMissing)

# Remove duplicated rows from the tourism statistics data

duplicatedRows <- duplicated(tourismStats) 
TourismStatsNoDup <- tourismStats[duplicatedRows == FALSE, ]

#### Merge classifications table with the tourism dataset ####
MergedResidentsClassifications <- merge(tourismStats, CountryCodesResidents, by="COUNTRY.OF.RESIDENCE", all.x=TRUE)
MergedFinalTourism <- merge(MergedResidentsClassifications, CountryCodesVisitors, by="COUNTRY.OF.RESIDENCE", all.x=TRUE)

View(MergedFinalTourism)




## CALCULATE THE AGE ##
# Take two dates and differentiate from one another

str(MergedFinalTourism$FLIGHT.DATE)
str(MergedFinalTourism$BIRTHDATE)

# Change date formats from character to Dates
MergedFinalTourism$FLIGHT.DATE <- as.Date(MergedFinalTourism$FLIGHT.DATE)
MergedFinalTourism$BIRTHDATE <- as.Date(MergedFinalTourism$BIRTHDATE)

# Check structure after changing formats
str(MergedFinalTourism$FLIGHT.DATE)
str(MergedFinalTourism$BIRTHDATE)

### Calculate Age ###
MergedFinalTourism$AGE <- as.numeric(difftime(MergedFinalTourism$FLIGHT.DATE, MergedFinalTourism$BIRTHDATE, unit = "weeks"))/52.25

# Round of Age to no decimal place
MergedFinalTourism$AGE = round(MergedFinalTourism$AGE, digits = 0)

View(MergedFinalTourism)

### Calculate LOS ###

MergedFinalTourism$INTENDED.DEP.DATE <- as.Date((MergedFinalTourism$INTENDED.DEP.DATE))
MergedFinalTourism$LENGTH.OF.STAY <- as.numeric(MergedFinalTourism$INTENDED.DEP.DATE - MergedFinalTourism$FLIGHT.DATE, unit = "days")

MergedFinalTourism$INTENDED.DEP.DATE = round(MergedFinalTourism$INTENDED.DEP.DATE, digits = 0)

View(MergedFinalTourism)

### Rename Travel Purpose without codes ###
# Fill NA in POV with "Returning Residents" #

str(MergedFinalTourism$TravelPurpose)

MergedFinalTourism$TravelPurpose[which(is.na(MergedFinalTourism$TravelPurpose))] <- "6. Returning Residents"

# Table 1: Summary of overseas Migration

Tab1_Summary <- MergedFinalTourism %>%
  group_by(PORT,ARR.DEPART,TravelPurpose) %>%
  count()


# Table 2: Purpose of Visit

tradeStatsSubset <- tradeStatsNoBanknotes[tradeStatsNoBanknotes$Type %in% c("EX / 1","EX / 3", "IM / 4", "IM / 7", "PC / 4"), ]
tradeStatsCommodities <- tradeStatsSubset[tradeStatsSubset$CP4 %in% c(1000, 3071, 4000, 4071, 7100), ]

PortSubset <- MergedFinalTourism[]
Tab2_POV <- MergedFinalTourism %>%
  filter(%in% c (PORT == "VAIRP", ARR.DEPART == "ARRIVAL") %>%
  group_by(PORT,TravelPurpose) %>%
  count()

write.csv("Tab2_POV.csv")


# Table 3: Country of Usual Residence

# Table 4: Residents arrival by Nationality

# Table 5: 




