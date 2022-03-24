#### Preparation ####

# Read in the processed TOURISM data from secure folder of the repository 
tourismProcessedFile <- file.path(secureDataFolder, "OUT_PROC_ASY_ProcessedRawData_31.12.19.csv")
tourismProcessed <- read.csv(tourismProcessedFile, header = TRUE, na.strings = c("","NA","NULL","null"))

#### Table 1: Summary of overseas Migration ####

Tab1_VILA <- tourismProcessed %>%
  group_by(PORT,ARR.DEPART,VisitorResident) %>%
  filter(PORT %in% c("VAIRP","VAIR","SAIRP","SAIR")) %>%
  count()

#### Table 2: Purpose of Visit ####

Tab2_POV <- tourismProcessed %>%
  group_by(TravelPurposeClassifications) %>%
  filter(VisitorResident %in% c("Visitor")) %>%
  filter (ARR.DEPART %in% c("ARRIVAL")) %>%
  count()

#### Table 3: Country of Usual Residence ####

Tab3_VistorsCUR <- tourismProcessed %>%
  filter(PORT %in% c("SAIR","SAIRP","VAIR","VAIRP"))%>%
  filter(VisitorResident %in%("Visitor"))%>%
  filter(ARR.DEPART %in%("ARRIVAL"))%>%
  group_by(GROUP) %>%
  count()

#### Table 4: Residents arrival by Nationality ####

Tab4_ResByNat<- tourismProcessed %>%
  group_by(RESIDENTS.BY.REGION) %>%
  filter(VisitorResident == "Resident", ARR.DEPART == 'ARRIVAL',PORT %in% c("VAIRP","VAIR","SAIRP","SAIR")) %>%
  count()

#### Table 5: Average length of stay ####

# Change length of stay format from integer to character
str(tourismProcessed$LENGTH.OF.STAY)
tourismProcessed$LENGTH.OF.STAY <- as.character(tourismStats$LENGTH.OF.STAY)
str(tourismProcessed$VisitorResident)

# Filter for visiting tourists 
visitors_data <- tourismProcessed %>%
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


#### Table 6: Average age ####

str(TourismFINAL$AGE)
str(TourismFINAL$VisitorResident)

visitor_data_age <- tourismProcessed %>%
  filter(VisitorResident == "Visitor")


visitor_data_age <- visitor_data_age %>%
  mutate(cities = case_when(
    grepl(pattern = "VAI", visitor_data_age$PORT) ~ "PORT VILA",
    grepl(pattern = "SAI", visitor_data_age$PORT) ~ "LUGANVILLE"
  )) 


AVG_AGE_VUV <- visitor_data_age %>%
  filter(ARR.DEPART == "ARRIVAL") %>%
  filter(AGE != "NA") %>%
  filter(cities %in% c("PORT VILA", "LUGANVILLE")) %>%
  group_by(cities) %>%
  summarise(AVG_AGE = round(mean(AGE), digits = 0)) 

Vanuatu_AGE <- mean(AVG_AGE_VUV$AVG_AGE)

cities <- "VANUATU"
AVG_AGE <- 38.5
Vanuatu <- data.frame(cities, AVG_AGE)

Tab6_AVGAGE <- rbind(AVG_AGE_VUV,Vanuatu)

#### Table 7: Visitors to outer islands (working on data) ####

# create subset from Main dataset #

DepartureSubset <-tourismProcessed[which(tourismProcessed$ARR.DEPART=="DEPARTURE"),]

# Merge Outer Island into Departure Subset #
FINALOuterIsland <- merge(MergeOuterIslandCodes,DepartureSubset, by="PASSPORT", all.x = TRUE )

Tab7_VisitorOuterIsland <- FINALOuterIsland %>%
  group_by(PORT, MainIsland) %>%
  filter(PORT %in% c("VAIRP","VAIR","SAIRP","SAIR")) %>%
  count()

#### Table 8: Visitors usual residence arrivals by Purpose of Visit ####

#grouping the purpose of visits
POV_Tab8 <- tourismProcessed %>%
  mutate( POV_Groups = case_when(
    grepl(pattern = "Bus", tourismProcessed$TravelPurposeClassifications) ~ "Business",
    grepl(pattern = "Conf", tourismProcessed$TravelPurposeClassifications) ~ "Business",
    grepl(pattern = "Edu", tourismProcessed$TravelPurposeClassifications) ~ "All Other",
    grepl(pattern = "Oth", tourismProcessed$TravelPurposeClassifications) ~ "All Other",
    grepl(pattern = "Spo", tourismProcessed$TravelPurposeClassifications) ~ "All Other",
    grepl(pattern = "Sto", tourismProcessed$TravelPurposeClassifications) ~ "All Other",
    grepl(pattern = "Hol", tourismProcessed$TravelPurposeClassifications) ~ "Holiday",
    grepl(pattern = "Vis", tourismProcessed$TravelPurposeClassifications) ~ "Visitings Friends and relatives"
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



#### END ####