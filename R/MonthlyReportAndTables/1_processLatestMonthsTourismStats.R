#### Preparation ####

# Clear the environment
rm(list = ls())

# Load the required libraries
#install.packages("dplyr")
#install.packages("tidyr")
#install.packages("janitor")
library(dplyr) # Manipulating data
library(openxlsx) # Reading and editing excel formatted files
library(tidyr) # Reading for pivot wide function
library(janitor) # Add totals row and column

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
tourismStatsFile <- file.path(secureDataFolder, "SEC_RAW_ASY_RawDataAndReferenceTables_31-12-19.csv")
tourismStats <- read.csv(tourismStatsFile, header = TRUE, na.strings = c("","NA","NULL","null"))

#### Data Cleaning (Missing Values, Duplicates) ####
## TourismStats <- read.csv(tourismStatsFile, Header = TRUE, na.strings=c ("","NA","NULL", "null"))

# count the missing values
numberMissing <- apply(TourismFINAL, MARGIN=2,
                       FUN=function(columnValues){
                         return(sum(is.na(columnValues)))
                       })
numberMissing


# Convert the counts to a proportion
proportionMissing <- numberMissing / nrow(TourismFINAL)

# Check for columns with high amounts of NA values
colsWithManyMissingValues <- names(proportionMissing)[proportionMissing > 0.1]
for(column in colsWithManyMissingValues){
  warning(paste0("Large amounts of missing data identified in \"", column, "\" column. View with: \n\tView(TourismFINAL[is.na(TourismFINAL[, \"",  column, "\"]), ])"))
}

# Remove duplicated rows from the tourism statistics data

duplicatedRows <- duplicated(TourismFINAL) 
TourismFINALNoDups <- TourismFINAL[duplicatedRows == FALSE, ]

#### CALCULATE THE AGE ####
# Take two dates and differentiate from one another
TourismFINAL<- tourismStats
str(TourismFINAL$FLIGHT.DATE)
str(TourismFINAL$BIRTHDATE)

TourismFINAL$FLIGHT.DATE <- as.character(TourismFINAL$FLIGHT.DATE)
str(TourismFINAL$FLIGHT.DATE)

TourismFINAL$BIRTHDATE <- as.character(TourismFINAL$BIRTHDATE)
str(TourismFINAL$BIRTHDATE)

# Change date formats from character to Dates
TourismFINAL$FLIGHT.DATE <- as.Date(TourismFINAL$FLIGHT.DATE, format = "%d/%m/%Y")
TourismFINAL$BIRTHDATE <- as.Date(TourismFINAL$BIRTHDATE, format = "%d/%m/%Y")

# Check structure after changing formats
str(TourismFINAL$FLIGHT.DATE)
str(TourismFINAL$BIRTHDATE)

### CALCULATE AGE ###
TourismFINAL$AGE <- as.numeric(difftime(TourismFINAL$FLIGHT.DATE, TourismFINAL$BIRTHDATE, unit = "weeks"))/52.25

# Round of Age to no decimal place
TourismFINAL$AGE = round(TourismFINAL$AGE, digits = 0)

TourismFINAL$AGE<- trunc(TourismFINAL$AGE)




#### Calculate LOS ####
TourismFINAL$INTENDED.DEP.DATE <- as.Date(TourismFINAL$INTENDED.DEP.DATE,  format = "%d/%m/%Y")
TourismFINAL$LENGTH.OF.STAY <- as.numeric(TourismFINAL$INTENDED.DEP.DATE - TourismFINAL$FLIGHT.DATE, unit = "days")
TourismFINAL$INTENDED.DEP.DATE = round(TourismFINAL$INTENDED.DEP.DATE, digits = 0)




#Read in the classifications data from open folder of the repository
CountryCodesByVisitors <- file.path(openDataFolder, "CountryCodes_VisByNat.csv")
CountryCodesVisitors <- read.csv(CountryCodesByVisitors, header = TRUE, na.strings = c("","NA","NULL","null"))

CountryCodesByResidents <- file.path(openDataFolder, "CountryCodes_ResbyNat.csv")
CountryCodesResidents <- read.csv(CountryCodesByResidents, header = TRUE, na.strings = c("","NA","NULL","null"))

TravelPurposeByCodes <-file.path(openDataFolder, "PurposeOfVisitCodes.csv")
TravelPurposeCodes <- read.csv(TravelPurposeByCodes, header = TRUE, na.strings = c("", "NA","NULL","null"))

OuterIslandByCodes <- file.path(openDataFolder, "tblIsland.csv")
OuterIslandCodes <- read.csv(OuterIslandByCodes, header = TRUE, na.strings = c("", "NA","NULL","null"))

VisitorOuterIsland <- file.path(openDataFolder, "Departure_OuterIslands.csv")
VisitorsOuterIslandComplete <- read.csv(VisitorOuterIsland, header = TRUE, na.strings = c("", "NA","NULL","null"))

#### Merge classifications table with the tourism dataset ####
# Fill NA in POV with "Returning Residents" #
MergedResidentsClassifications <- merge(tourismStats, CountryCodesResidents, by="COUNTRY.OF.RESIDENCE", all.x=TRUE)
MergedFinalTourism <- merge(MergedResidentsClassifications, CountryCodesVisitors, by="COUNTRY.OF.RESIDENCE", all.x=TRUE)

str(MergedFinalTourism$TravelPurpose)
MergedFinalTourism$TravelPurpose <- as.character(MergedFinalTourism$TravelPurpose)
MergedFinalTourism$TravelPurpose[which(is.na(MergedFinalTourism$TravelPurpose))] <- "6. Returning Residents"
TourismFINAL <- merge(MergedFinalTourism, TravelPurposeCodes, by="TravelPurpose", all.x=TRUE)

MergeOuterIslandCodes <- merge(OuterIslandCodes, VisitorsOuterIslandComplete, by="IslandName", all.x = TRUE)