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

TravelPurposeByCodes <-file.path(openDataFolder, "PurposeOfVisitCodes.csv")
TravelPurposeCodes <- read.csv(TravelPurposeByCodes, header = TRUE, na.strings = c("", "NA","NULL","null"))

#### Merge classifications table with the tourism dataset ####
# Fill NA in POV with "Returning Residents" #
MergedResidentsClassifications <- merge(tourismStats, CountryCodesResidents, by="COUNTRY.OF.RESIDENCE", all.x=TRUE)
MergedFinalTourism <- merge(MergedResidentsClassifications, CountryCodesVisitors, by="COUNTRY.OF.RESIDENCE", all.x=TRUE)
MergedFinalTourism$TravelPurpose[which(is.na(MergedFinalTourism$TravelPurpose))] <- "6. Returning Residents"
TourismFINAL <- merge(MergedFinalTourism, TravelPurposeCodes, by="TravelPurpose", all.x=TRUE)

#### Data Cleaning (Missing Values, Duplicates) ####
## TourismStats <- read.csv(tourismStatsFile, Header = TRUE, na.strings=c ("","NA","NULL", "null"))

# count the missing values
numberMissing <- apply(TourismFINAL, MARGIN=2,
                       FUN=function(columnValues){
                         return(sum(is.na(columnValues)))
                       })

# Remove duplicated rows from the tourism statistics data

duplicatedRows <- duplicated(TourismFINAL) 
TourismFINALNoDups <- TourismFINAL[duplicatedRows == FALSE, ]

#### CALCULATE THE AGE ####
# Take two dates and differentiate from one another

str(TourismFINAL$FLIGHT.DATE)
str(TourismFINAL$BIRTHDATE)

# Change date formats from character to Dates
TourismFINAL$FLIGHT.DATE <- as.Date(TourismFINAL$FLIGHT.DATE)
TourismFINAL$BIRTHDATE <- as.Date(TourismFINAL$BIRTHDATE)

# Check structure after changing formats
str(TourismFINAL$FLIGHT.DATE)
str(TourismFINAL$BIRTHDATE)

### CALCULATE AGE ###
TourismFINAL$AGE <- as.numeric(difftime(TourismFINAL$FLIGHT.DATE, TourismFINAL$BIRTHDATE, unit = "weeks"))/52.25

# Round of Age to no decimal place
TourismFINAL$AGE = round(TourismFINAL$AGE, digits = 0)

#### Calculate LOS ####
TourismFINAL$INTENDED.DEP.DATE <- as.Date((TourismFINAL$INTENDED.DEP.DATE))
TourismFINAL$LENGTH.OF.STAY <- as.numeric(TourismFINAL$INTENDED.DEP.DATE - TourismFINAL$FLIGHT.DATE, unit = "days")

TourismFINAL$INTENDED.DEP.DATE = round(TourismFINAL$INTENDED.DEP.DATE, digits = 0)

#### Table 1: Summary of overseas Migration ####

Tab1_VILA <- TourismFINAL %>%
  group_by(PORT,ARR.DEPART,VisitorResident) %>%
  filter(PORT %in% c("VAIRP","VAIR","SAIRP","SAIR")) %>%
  count()

#### Table 2: Purpose of Visit ####

Tab2_POV <- TourismFINAL %>%
  group_by(TravelPurpose, VisitorResident) %>%
  filter(VisitorResident %in% c("Visitor")) %>%
  count()

#### Table 3: Country of Usual Residence ####

#### Table 4: Residents arrival by Nationality ####

#### Table 5: Average LOS ####

#### Table 6: Average Age ####




