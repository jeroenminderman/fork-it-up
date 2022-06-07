#### Preparation ####

### Some change
### Another change

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

# Note the secure data path
secureDataFolder <- file.path(repository, "data", "secure")

# Note the open data path
openDataFolder <- file.path(repository, "data", "open")

# Read in the raw TOURISM data from secure folder of the repository 
# Note that spaces in column names have been automatically replaced with "."s
tourismStatsFile <- file.path(secureDataFolder, "SEC_RAW_ASY_RawDataAndReferenceTables_31-12-19.csv")
tourismStats <- read.csv(tourismStatsFile, header = TRUE, na.strings = c("","NA","NULL","null"))

#### Explore the extent of missing data ####

# count the missing values
numberMissing <- apply(tourismStats, MARGIN=2,
                       FUN=function(columnValues){
                         return(sum(is.na(columnValues)))
                       })

# Convert the counts to a proportion
proportionMissing <- numberMissing / nrow(tourismStats)

# Check for columns with high amounts of NA values
colsWithManyMissingValues <- names(proportionMissing)[proportionMissing > 0.1]
for(column in colsWithManyMissingValues){
  warning(paste0("Large amounts of missing data identified in \"", column, "\" column. View with: \n\tView(tourismStats[is.na(tourismStats[, \"",  column, "\"]), ])"))
}

#### Clean and process the latest month's data ####

# Change flight date format from character to date
tourismStats$FLIGHT.DATE <- as.character(tourismStats$FLIGHT.DATE)
tourismStats$FLIGHT.DATE <- as.Date(tourismStats$FLIGHT.DATE, format = "%d/%m/%Y")
#tourismFINALNoDups$BIRTHDATE <- as.Date(tourismFINALNoDups$BIRTHDATE, tryFormats = c("%Y-%m-%d", "%Y/%m/%d"))

# Change birth date format from character to date
tourismStats$BIRTHDATE <- as.character(tourismStats$BIRTHDATE)
tourismStats$BIRTHDATE <- as.Date(tourismStats$BIRTHDATE, format = "%d/%m/%Y")

# Change intended depature date format from character to date
tourismStats$INTENDED.DEP.DATE <- as.character(tourismStats$INTENDED.DEP.DATE)
tourismStats$INTENDED.DEP.DATE <- as.Date(tourismStats$INTENDED.DEP.DATE,  format = "%d/%m/%Y")

# Check structure after changing formats
str(tourismStats$FLIGHT.DATE)
str(tourismStats$BIRTHDATE)
str(tourismStats$INTENDED.DEP.DATE)

# Substitute values for returning residents in empty spaces 
tourismStats$TravelPurpose <- as.character(tourismStats$TravelPurpose)
tourismStats$TravelPurpose[which(is.na(tourismStats$TravelPurpose))] <- "6. Returning Residents"

# Remove duplicated rows from the tourism statistics data
duplicatedRows <- duplicated(tourismStats) 
tourismFINALNoDups <- tourismStats[duplicatedRows == FALSE, ]

#### Calculate the age of passengers ####

### CALCULATE AGE ###
tourismFINALNoDups$AGE <- as.numeric(difftime(tourismFINALNoDups$FLIGHT.DATE, tourismFINALNoDups$BIRTHDATE, unit = "weeks"))/52.25
tourismFINALNoDups$AGE<- trunc(tourismFINALNoDups$AGE) # Extract age of passengers

#### Calculate length of stay of passengers ####

# Calculate number of days staying
tourismFINALNoDups$LENGTH.OF.STAY <- as.numeric(tourismFINALNoDups$INTENDED.DEP.DATE - tourismFINALNoDups$FLIGHT.DATE, unit = "days")
tourismFINALNoDups$INTENDED.DEP.DATE = round(tourismFINALNoDups$INTENDED.DEP.DATE, digits = 0)

#### Extract the Year, Month, and Day as separate columns ####

# Create new columns for dates
tourismFINALNoDups$Year <- format(tourismFINALNoDups$FLIGHT.DATE, format= "%Y")
tourismFINALNoDups$Month <- format(tourismFINALNoDups$FLIGHT.DATE, format= "%B")
tourismFINALNoDups$Day <- format(tourismFINALNoDups$FLIGHT.DATE, format= "%d")

#### Merge in classification tables ####

# Read in the vistors classifications into data frame
countryCodesByVisitors <- file.path(openDataFolder, "CountryCodes_VisByNat.csv")
countryVisitors <- read.csv(countryCodesByVisitors, header = TRUE, na.strings = c("","NA","NULL","null"))
tourismMergedWithClassifications <- merge(tourismFINALNoDups, countryVisitors, by="COUNTRY.OF.RESIDENCE", all.x=TRUE)

# Read in the residents classifications into data frame
countryCodesByResidents <- file.path(openDataFolder, "CountryCodes_ResbyNat.csv")
countryResidents <- read.csv(countryCodesByResidents, header = TRUE, na.strings = c("","NA","NULL","null"))
tourismMergedWithClassifications <- merge(tourismMergedWithClassifications, countryResidents, by="COUNTRY.OF.RESIDENCE", all.x=TRUE)

# Read in the purpose of visit classifications into data frame
travelPurposeByCodes <-file.path(openDataFolder, "PurposeOfVisitCodes.csv")
travelPurpose <- read.csv(travelPurposeByCodes, header = TRUE, na.strings = c("", "NA","NULL","null"))
travelPurpose$TravelPurpose <- as.character(travelPurpose$TravelPurpose)
travelPurpose$TravelPurpose[which(is.na(travelPurpose$TravelPurpose))] <- "6. Returning Residents"
tourismMergedWithClassifications <- merge(tourismMergedWithClassifications, travelPurpose, by="TravelPurpose", all.x=TRUE)
#travelPurpose$TravelPurpose <- as.character(travelPurpose$TravelPurpose)
#travelPurpose$TravelPurpose[which(is.na(travelPurpose$TravelPurpose))] <- "6. Returning Residents"

# Read in the outerisland travel classifications into data frame (have to get this data from Air Vanuatu)
outerIslandByCodes <- file.path(openDataFolder, "tblIsland.csv")
outerIsland <- read.csv(outerIslandByCodes, header = TRUE, na.strings = c("", "NA","NULL","null"))
#tourismMergedWithClassifications <- merge(tourismMergedWithClassifications, outerIsland, by="IslandName", all.x = TRUE)
#VisitorOuterIsland <- file.path(openDataFolder, "Departure_OuterIslands.csv")
#VisitorsOuterIslandComplete <- read.csv(VisitorOuterIsland, header = TRUE, na.strings = c("", "NA","NULL","null"))

#### Finish ####

# Make copy of latest month's processed data
processedTourismStats <- tourismMergedWithClassifications

# Create csv of last months processed data
outputDataFile <- file.path(secureDataFolder, paste("OUT_PROC_ASY_ProcessedRawData_31.12.19.csv"))
write.csv(processedTourismStats, outputDataFile)