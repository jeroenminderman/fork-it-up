#### Preparation ####

# Check that 1_processLatestMonthsTradeStats.R has already been run
if(exists("repository") == FALSE || exists("secureDataFolder") == FALSE || 
   exists("openDataFolder") == FALSE || exists("processedTradeStats") == FALSE){
  stop("Necessary variables have not been created. Please ensure that all code in \"1_processLatestMonthsTradeStats.R\" has been executed successfully.")
}

# Note the outputs folder
outputsFolder <- file.path(repository, "outputs")

# Note current year and month of data
date <- max(processedTradeStats$Reg..Date, na.rm=TRUE)
month <- format(date, "%B")
year <- format(date, "%Y")

# Note the excel workbook containing the final formatted tables
finalWorkbookFileName <- file.path(outputsFolder, "SEC_FINAL_MAN_FinalTradeStatisticsTables_30-11-19_WORKING.xlsx")

# Load the excel file
finalWorkbook <- openxlsx::loadWorkbook(finalWorkbookFileName)

# Print progress
cat("Finished loading excel file with formatted tables.\n")

