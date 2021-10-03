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

str(MergedFinalTourism$TravelPurpose)
MergedFinalTourism$TravelPurpose <- as.character(MergedFinalTourism$TravelPurpose)
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

#### Calculate LOS ####
TourismFINAL$INTENDED.DEP.DATE <- as.Date(TourismFINAL$INTENDED.DEP.DATE,  format = "%d/%m/%Y")
TourismFINAL$LENGTH.OF.STAY <- as.numeric(TourismFINAL$INTENDED.DEP.DATE - TourismFINAL$FLIGHT.DATE, unit = "days")

TourismFINAL$INTENDED.DEP.DATE = round(TourismFINAL$INTENDED.DEP.DATE, digits = 0)

#### Table 1: Summary of overseas Migration ####

Tab1_VILA <- TourismFINAL %>%
  group_by(PORT,ARR.DEPART,VisitorResident) %>%
  filter(PORT %in% c("VAIRP","VAIR","SAIRP","SAIR")) %>%
  count()

#### Table 2: Purpose of Visit ####

Tab2_POV <- TourismFINAL %>%
  group_by(TravelPurposeClassifications) %>%
  filter(VisitorResident %in% c("Visitor")) %>%
  filter (ARR.DEPART %in% c("ARRIVAL")) %>%
  count()

#### Table 3: Country of Usual Residence ####
Tab3_VistorsCUR <- TourismFINAL %>%
  filter(PORT %in% c("SAIR","SAIRP","VAIR","VAIRP"))%>%
  filter(VisitorResident %in%("Visitor"))%>%
  filter(ARR.DEPART %in%("ARRIVAL"))%>%
  group_by(GROUP) %>%
  count()
  


#### Table 4: Residents arrival by Nationality ####
Tab4_ResByNat<- TourismFINAL %>%
  group_by(PORT,RESIDENTS.BY.REGION) %>%
  filter(VisitorResident == "Resident", ARR.DEPART == 'ARRIVAL',PORT %in% c("VAIRP","VAIR","SAIRP","SAIR")) %>%
  count()


#### Table 5: Average LOS ####

str(TourismFINAL$LENGTH.OF.STAY)
str(TourismFINAL$VisitorResident)

visitors_data <- TourismFINAL %>%
  filter(VisitorResident == "Visitor")


# Identify when Visitors length of stay over 120 days and correct
LOS_threshold <- 120
visitors_data <- visitors_data %>%
  mutate(visit_too_long = LENGTH.OF.STAY > LOS_threshold,
         CORRECTED.LENGTH.OF.STAY = case_when(
           LENGTH.OF.STAY > LOS_threshold ~ 120,
           TRUE ~ LENGTH.OF.STAY
         ))


visitors_data <- visitors_data %>%
  mutate(Towns = case_when(
    grepl(pattern = "VAI", visitors_data$PORT) ~ "PORT VILA",
    grepl(pattern = "SAI", visitors_data$PORT) ~ "LUGANVILLE"
  )) 


AVG_LOS_VUV <- visitors_data %>%
  filter(ARR.DEPART == "ARRIVAL") %>%
  filter(CORRECTED.LENGTH.OF.STAY != "NA") %>%
  filter(Towns %in% c("PORT VILA", "LUGANVILLE")) %>%
  group_by(Towns) %>%
  summarise(AVG_LengthOfStay = round(mean(CORRECTED.LENGTH.OF.STAY), digits = 0)) 

Vanuatu_LOS <- mean(AVG_LOS_VUV$AVG_LengthOfStay)

Towns <- "VANUATU"
AVG_LengthOfStay <- 12
Vanuatu <- data.frame(Towns,AVG_LengthOfStay)

Tab5_AVGLOS <- rbind(AVG_LOS_VUV,Vanuatu)



#### Table 6: Average Age ####

Tab6_AverageAge <- TourismFINAL %>%
  filter(VisitorResident == "Visitor") %>%
  filter(PORT %in% c("VAIR","VAIRP", "SAIR", "SAIRP")) %>%
  group_by(PORT) %>% 
  summarise(total= mean(AGE))

#### Table 7: Visitors travelling to outer islands ####

#### Table 8: Visitors usual residence arrivals by Purpose of Visit ####

#grouping the purpose of visits
POV_Tab8 <- TourismFINAL %>%
  mutate( POV_Groups = case_when(
    grepl(pattern = "Bus", TourismFINAL$TravelPurposeClassifications) ~ "Business",
    grepl(pattern = "Conf", TourismFINAL$TravelPurposeClassifications) ~ "Business",
    grepl(pattern = "Edu", TourismFINAL$TravelPurposeClassifications) ~ "All Other",
    grepl(pattern = "Oth", TourismFINAL$TravelPurposeClassifications) ~ "All Other",
    grepl(pattern = "Spo", TourismFINAL$TravelPurposeClassifications) ~ "All Other",
    grepl(pattern = "Sto", TourismFINAL$TravelPurposeClassifications) ~ "All Other",
    grepl(pattern = "Hol", TourismFINAL$TravelPurposeClassifications) ~ "Holiday",
    grepl(pattern = "Vis", TourismFINAL$TravelPurposeClassifications) ~ "Visitings Friends and relatives"
  )) 

#creating the data frame comparing visitors country of usual residence by purpose of visit
UsualResByPOV <- POV_Tab8 %>%
  group_by(GROUP, POV_Groups) %>%
  filter(VisitorResident %in% c("Visitor")) %>%
  filter (ARR.DEPART %in% c("ARRIVAL")) %>%
  count()


#manipulating the above to a wider format
UsualResByPOV2 <- UsualResByPOV %>%
  pivot_wider (names_from = POV_Groups, values_from = n)

### Add in total row and column ###
#adding a total column

UsualResByPOV2 <- UsualResByPOV2 %>%
  adorn_totals("col")

#adding a total row
UsualResByPOV2 <- UsualResByPOV2 %>%
  adorn_totals("row")

##Percentage of Total Columns

#Calculate Percentage of total
n_cols <- ncol(UsualResByPOV2)
columns_of_interest <- 2:(n_cols - 1) #ignores the first and last columns
UsualResByPOV_proportion <- UsualResByPOV2[, columns_of_interest]/UsualResByPOV2$Total
UsualResByPOV_percentage <- UsualResByPOV_proportion * 100

#Note percentage columns in their name
colnames(UsualResByPOV_percentage) <- paste(colnames(UsualResByPOV_percentage),"(% share)")

#Add percentage columns into main table
UsualResByPOV_percentage <- cbind(UsualResByPOV2, UsualResByPOV_percentage)

##Combining Value and Percentage into one single column

#Initialise a dataframe to store the combined values and percentages
Tab8_UsualResByPOV <- data.frame("GROUP" = UsualResByPOV2$GROUP)

#Note columns of interest
columns_of_interest <- colnames(UsualResByPOV2)[
  c(-1, -ncol(UsualResByPOV2))
]

#Add in each column with combined value and percentages
for(travel_type_column in columns_of_interest){
  
  #create column name with percentage
  travel_type_column_as_percentage <- paste(travel_type_column, "(% share)")
 
  #get the counts for current travel type
  #Note use of "drop" - this extracts the values as a vector instead of a data.frame
  
  travel_type_counts <- UsualResByPOV2 [, travel_type_column, drop = TRUE]
 
  #get the counts as percentages
  #Note use of "drop" - this extracts the values as a vector instead of a data.frame
  
  travel_type_percentages <- UsualResByPOV_percentage [,
                                                       travel_type_column_as_percentage, drop = TRUE]
  
  #Round as percentages
  travel_type_percentages <- round(travel_type_percentages, digits = 0)
  
  #create column with combined value and percentage
  Tab8_UsualResByPOV[, travel_type_column_as_percentage] <- paste(
    travel_type_counts, "(", travel_type_percentages, "%)", sep = ""
  )
}

#Add a total column
Tab8_UsualResByPOV$Total <- UsualResByPOV2$Total

#Remove the percentages from final row

number_rows <- nrow(Tab8_UsualResByPOV)
Tab8_UsualResByPOV[number_rows,] <- UsualResByPOV2[number_rows, ]
Tab8_UsualResByPOV <- Tab8_UsualResByPOV[ , c(1,5, 4, 3, 2, 6)]

#write.csv(UsualResByPOV, file="Table_8.csv")

