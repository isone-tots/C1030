
###    THIS SCRIPT MUST BE RAN ON A NON-WINDOWS 10 MACHINE, OR TERMINAL SERVER UNTIL \\RTSMB IS UPGRADED    ###


########### ISO Initiated Audit Parameters as of 6-11-2018: ###########
#
#   1. Ignore 0 start weightings
#   2. Units with Caps > 0
#   3. Minimum MW = 2MW
#   4. Units who have not started in last 3 months
#   5. Only look at starts in last 12 monts
#   6. Ratio of LP requested to Economic Starts > .5
#   7.
#
#######################################################################

Sys.setenv(JAVA_HOME = "C:\\Program Files\\Java\\openjdk8u232\\jre") # for 64-bit version
# install packages
if (!require("pacman")) install.packages("pacman")
pacman::p_load(readr, xlsx, RJDBC, splitstackshape, readxl, lubridate, dplyr, svDialogs, tcltk)

# clear the environment pane
rm(list = ls())

# Set the time
currDate <- Sys.time()


# PARAMETERS
minMW <- 2
capMW <- 0
startRatio <- .5
oneYear <- currDate - (3600 * 24 * 365)

# Set the date represeting 3 months ago
pastDate <- currDate - (3600 * 24 * 90)

# set the file location for the weekly process sheets
filelocation <- "\\\\iso-ne.com\\shares\\tso\\Claim 10_30 Auditing\\ISO Initiated C10 Audits\\Weekly Data\\"

# set the file location for the output file
outputlocation <- "\\\\iso-ne.com\\shares\\tso\\Claim 10_30 Auditing\\ISO Initiated C10 Audits\\"

# outputlocation2 <- "\\\\moss.iso-ne.com@SSL\\sites\\rcd\\"
# outputlocation2 <- "https://moss.iso-ne.com/sites/rcd/VAR Tests/"

# setwd("\\\\moss.iso-ne.com@SSL\\sites\\rcd\\")
# setwd("\\\\iso-ne.com\\shares\\tso\\Claim 10_30 Auditing\\ISO Initiated C10 Audits\\")


# Make a list of the directory contents
list <- file.info(list.files(filelocation, full.names = TRUE), drop = FALSE)
list <- tibble::rownames_to_column(list, "Path")

# sort file modified time to newest to oldest
list <- arrange(list, desc(mtime))

# Set names for the list
names(list) <- c("Path", "size", "isdir", "mode", "mtime", "ctime", "atime", "exe")

# Get the filename for use in the import
latestfile <- list[1, "Path"]


# Get usernames and passwords

My_Username <- Sys.info()["user"]

PWLocation <- paste0("//iso-ne.com//shares//", My_Username, "//passwords.csv")

passwords <- read.csv(PWLocation)

My_Password1 <- as.character(passwords[2, 2])

My_Password2 <- as.character(passwords[1, 2])


#----------------- Import Data from Weekly Process and Contacts Sheet to dataframes named 'startdata' and 'contactsheet'

# Get unit start data and create df

startdata <-read_excel("\\\\iso-ne.com\\shares\\tso\\Claim 10_30 Auditing\\ISO Initiated C10 Audits\\Weekly Data\\C1030_Weekly_Process_Tab_Data with weightings.xlsx", 
             col_types = c("numeric", "text", "date", 
                           "date", "numeric", "numeric", "numeric", 
                           "numeric", "numeric", "numeric", 
                           "numeric", "numeric", "numeric", 
                           "numeric", "numeric", "numeric", 
                           "numeric", "numeric", "numeric", 
                           "numeric", "numeric", "numeric", 
                           "numeric", "numeric", "numeric", 
                           "numeric", "numeric", "text", "text", 
                           "numeric", "numeric", "numeric", 
                           "numeric", "numeric", "numeric"))

# Get contact sheet data and create df (this is now getting cap values from UI, but didnt rename contact)


contactsheet <- read_excel("\\\\iso-ne.com\\shares\\tso\\Claim 10_30 Auditing\\ISO Initiated C10 Audits\\Weekly Data\\OI_WEEKLYCLAIM1030_01-16-2020.xls", 
           col_types = c("text", "numeric", "date", 
                         "numeric", "numeric", "numeric", 
                         "numeric", "numeric", "numeric", 
                         "numeric", "numeric", "numeric", 
                         "numeric"))



# contactsheet <-
#   read_excel(
#     latestfile,
#     sheet = "Sheet2",
#     col_types = c(
#       "numeric",
#       "text",
#       "skip",
#       "skip",
#       "skip",
#       "skip",
#       "skip",
#       "skip",
#       "numeric",
#       "numeric",
#       "numeric",
#       "numeric",
#       "numeric",
#       "numeric",
#       "numeric",
#       "numeric",
#       "numeric",
#       "numeric",
#       "numeric",
#       "numeric",
#       "numeric",
#       "numeric",
#       "numeric",
#       "skip",
#       "numeric",
#       "numeric",
#       "skip",
#       "skip"
#     ),
#     skip = 1
#   )

#startdata$UNIT_SHORT_NAME <- factor(startdata$UNIT_SHORT_NAME)

# combine the two DFs, on Unit short Name
mergedata <- merge(startdata, contactsheet, by.x = "UNIT_SHORT_NAME", by.y = "Asset Short Name")



# Delete duplicate Unit ID column from merge
#mergedata$"Resource ID.y" <- NULL

#--------------------------------------------Filter 

# PARAMETER: on only starts with weightings 1 to 10
mergedata <- mergedata[mergedata$WEIGHTING_10 >= 1, ]


# PARAMETER: out units that have no 30 minute cap
mergedata <- mergedata[mergedata$`Claim 30` > capMW, ]

#--------------------------------------------Create new columns needed for building audit list

#-------Establish new column and set Audit Target MW
mergedata$Desired10 <- do.call(pmax, mergedata[40:41])



#-------Establish new column and set a flag indicating if the Target MW is greater than audit MW threshold (T/F) [F gets excluded]
mergedata$AuditTarget <- mergedata$Desired10 > minMW
#-------Establish new column and set a flag indicating if the audit is older than a year from todays date (T/F) [not currently used]
mergedata$Last12Month <- mergedata$TIMESTAMP_RECEIVED_LOCAL > oneYear

#Filter on only audit targets == TRUE
mergedata <- mergedata[mergedata$AuditTarget == TRUE, ]



contactInfo <- mergedata 

################################################################################

# ---- Get list of units who have started in the last three months, apply filters, get targets, then find out who has not started

# Who has started in last 3 months
threemonths <- contactInfo[contactInfo$TIMESTAMP_RECEIVED_LOCAL > pastDate, ]

# filter on only those with a 30 cap > 0
threemonths <- threemonths[threemonths$`Claim 30` > capMW, ]
threemonths <- unique(threemonths[, c("ASSET_ID", "UNIT_SHORT_NAME", "Desired10")])

# compare three-month start list to list of all FS units
threemonthsfinal <- merge(contactsheet, threemonths, by.x = "Asset Short Name", by.y = "UNIT_SHORT_NAME", all.x = TRUE)

# filter on NAs (those who have not started)
threemonthsfinal <- threemonthsfinal[is.na(threemonthsfinal$ASSET_ID), ]




threemonthsfinal <- merge(threemonthsfinal, contactInfo, by.x = "Asset Short Name", by.y = "UNIT_SHORT_NAME", all.x = TRUE)
threemonthsfinal <- unique(threemonthsfinal[, c("Asset ID.x", "Asset Short Name", "Desired10.y")])
threemonthsfinal <- threemonthsfinal[!is.na(threemonthsfinal$Desired10.y), ]


# Remove the DRRs
threemonthsfinalcap <- threemonthsfinal[!grepl("^Z", threemonthsfinal$`Asset Short Name`), ]
threemonthsfinalcap$Reason <- " > 90 days since last start"
# rename the columns
names(threemonthsfinalcap) <- c("Resource ID", "Resource Short Name", "Desired10", "Reason")


targetData <- threemonthsfinalcap


#-------------------------------------------Count the number of LP requested audits
# Create df of the unit, audit flags
auditRequests <- data.frame(mergedata[c("UNIT_SHORT_NAME", "REQUEST_AUDIT_10_FLAG")])

# count flags by putting in summary table
freqs <- table(
  auditRequests$UNIT_SHORT_NAME,
  auditRequests$REQUEST_AUDIT_10_FLAG
)

# back to a data frame
freqsDF <- as.data.frame.matrix(freqs)

# Establish new columns and calculate audit request ratio and total number of audits
freqsDF$Ratio <- (freqsDF$"1" / (freqsDF$"1" + freqsDF$"0"))
freqsDF$Total <- (freqsDF$"1" + freqsDF$"0")
# put back the column names from the earlier transform
freqsDF <- setNames(
  cbind(rownames(freqsDF), freqsDF, row.names = NULL),
  c("Resource Short Name", "N", "Y", "Ratio", "Total")
)



# Remove NAs, get rid of DRRs, get Desired10 values, and apply the reason code
freqsDFfilter <- freqsDF[freqsDF$Ratio > .5, ]

freqsDFfilter <- freqsDFfilter[!grepl("^Z", freqsDFfilter$`Resource Short Name`), ]
freqsDFfilter$Reason <- " > 50% audit to start ratio"

freqsDFfilter <- merge(freqsDFfilter, contactInfo, by.x = "Resource Short Name", by.y = "UNIT_SHORT_NAME", all.x = TRUE)
freqsDFfilter <- freqsDFfilter[, c("ASSET_ID", "Resource Short Name", "Desired10", "Reason"), drop = F]
names(freqsDFfilter) <- c("Resource ID", "Resource Short Name", "Desired10", "Reason")


freqsDFfilter <- unique((freqsDFfilter))

# Append the results to the threemonthcap list

AuditAssets <- rbind(threemonthsfinalcap, freqsDFfilter)
AuditAssets <- AuditAssets[order(AuditAssets$`Resource ID`), ]

# Create final DF and assign names that BI will be able to use
AuditList <- AuditAssets
names(AuditList) <- c("EQUIPMENT_NAME_ID", "ASSET_ID", "Desired10", "Reason") ##### GOOD TO HERE

assetList <- AuditList %>% select(1)
assetList <- as.data.frame(assetList)
assetList <- assetList[order(assetList), ]
assetList <- paste("", as.character(assetList), "", collapse = ", ", sep = "")




# ______________________________________________________________________END LIST CODE




# ______________________________________________________________________BEGIN BI CODE


# Query 1 for BI system. BIPROD for warehouse. BIPROD2 for Direct Access
BI_System_1 <- "BIPROD"

# Query 2 for BI system. BIPROD for warehouse. BIPROD2 for Direct Access
BI_System_2 <- "BIPROD2"

# Query 1 for BI - this query returns assets that are bidding as fast start from the list of units that havent started > 90 days

BI_Query_1 <- paste0("

SELECT   \"Energy_Ancillary_Services_Markets_Data\".\"- Asset Operational Data\".\"Fast Start Capable Flag\" as FAST_START_CAPABLE_FLAG, 
   \"Energy_Ancillary_Services_Markets_Data\".\"- Asset Operational Data\".\"Fast Start Qualified Flag\" as FAST_START_QUALIFIED_FLAG, 
                     \"Energy_Ancillary_Services_Markets_Data\".\"- Time Offsets\".\"Local Day Offset\" as LOCAL_DAY_OFFSET, 
                     \"Energy_Ancillary_Services_Markets_Data\".\"Asset Dimension\".\"Asset ID\" as ASSET_ID, 
                     \"Energy_Ancillary_Services_Markets_Data\".\"Asset Dimension\".\"Asset Name\" as ASSET_NAME, 
                     \"Energy_Ancillary_Services_Markets_Data\".\"Time Dimension\".\"Local Day\" as LOCAL_DAY, 
                     \"Energy_Ancillary_Services_Markets_Data\".\"- Real Time Asset Offers\".\"RT Cold Notification Time\" as RT_COLD_NOTIFICATION_TIME, 
                     \"Energy_Ancillary_Services_Markets_Data\".\"- Real Time Asset Offers\".\"RT Cold Startup Time\" as RT_COLD_STARTUP_TIME, 
                     \"Energy_Ancillary_Services_Markets_Data\".\"- Real Time Asset Offers\".\"RT Minimum Down Time\" as RT_MINIMUM_DOWN_TIME, 
                     \"Energy_Ancillary_Services_Markets_Data\".\"- Real Time Asset Offers\".\"RT Minimum Run Time\" as RT_MINIMUM_RUN_TIME
                     FROM\"Energy_Ancillary_Services_Markets_Data\" WHERE ((\"- Real Time Asset Offers\".\"RT Minimum Down Time\" <= 1) AND (\"- Real Time Asset Offers\".\"RT Minimum Run Time\" <= 1) AND (\"- Real Time Asset Offers\".\"RT Cold Notification Time\"+\"- Real Time Asset Offers\".\"RT Cold Startup Time\" < 0.5) AND (\"- Time Offsets\".\"Local Day Offset\" BETWEEN -7 AND -1) AND (\"- Asset Operational Data\".\"Fast Start Qualified Flag\" <> 'N') AND (\"- Asset Operational Data\".\"Fast Start Qualified Flag\" IS NOT NULL) AND (\"Asset Dimension\".\"Asset ID\" IN (", assetList, ")))


")

assetList2 <- AuditList %>% select(2)
assetList2 <- as.data.frame(assetList2)
assetList2 <- assetList2[order(assetList2), ]
assetList2 <- paste("'", as.character(assetList2), "'", collapse = ", ", sep = "")

# Query 2 for BI - this query returns the assets that are
BI_Query_2 <- paste0("

SELECT   \"Outage\".\"Equipment\".\"Equipment Name ID\" as EQUIPMENT_NAME_ID, 
   \"Outage\".\"Outages\".\"Actual End Date (Local)\" as ACTUAL_END_DATE, 
                     \"Outage\".\"Outages\".\"Actual Start Date (Local)\" as ACTUAL_START_DATE, 
                     \"Outage\".\"Outages\".\"Outage Category Description\" as OUTAGE_CATEGORY_DESCRIPTION, 
                     \"Outage\".\"Outages\".\"Outage Request ID\" as OUTAGE_REQUEST_ID, 
                     \"Outage\".\"Outages\".\"Outage Status Description\" as OUTAGE_STATUS_DESCRIPTION, 
                     \"Outage\".\"Outages\".\"Planned End Date (Local)\" as PLANNED_END_DATE, 
                     \"Outage\".\"Outages\".\"Planned Start Date (Local)\" as PLANNED_START_DATE, 
                     \"Outage\".\"Outages\".\"Outage Count\" as OUTAGE_COUNT
                     FROM\"Outage\" WHERE ((\"Outages\".\"Outage Category Description\" = 'Generation') AND (\"Outages\".\"Planned Start Date (Local)\" < (CURRENT_DATE+10)) AND (\"Equipment\".\"Equipment Name ID\" IS NOT NULL) AND (\"Outages\".\"Actual End Date (Local)\" IS NULL) AND (\"Outages\".\"Outage Status Description\" IN ('Approved', 'Implemented')) AND (\"Equipment\".\"Equipment Name ID\" IN (",assetList2,")))

")

# Loading JDBC to connect to Oracle
driver <- JDBC("oracle.bi.jdbc.AnaJdbcDriver", "c:/RJDBC/bijdbc.jar", identifier.quote = "'")

# Connect to BI
Connection_1 <- dbConnect(driver, paste("jdbc:oraclebi://", BI_System_1, ".iso-ne.com:9703/", sep = ""), My_Username, My_Password1)

# Query Results
Query_Results_1 <- dbGetQuery(Connection_1, BI_Query_1)

# Connect to BI for query 2
Connection_2 <- dbConnect(driver, paste("jdbc:oraclebi://", BI_System_2, ".iso-ne.com:9703/", sep = ""), My_Username, My_Password2)

# Query Results
Query_Results_2 <- dbGetQuery(Connection_2, BI_Query_2)

# Close connection 1
dbDisconnect(Connection_1)
# Close connection 2
dbDisconnect(Connection_2)


# ____________________________________________________________Final Ouput Section

# Purge all User password data from environment
rm(My_Password1)
rm(My_Password2)
rm(passwords)

offerBehavior <- AuditList
offerBehavior <- offerBehavior[order(offerBehavior$ASSET_ID), ]
offerBehavior <- as.data.frame(offerBehavior)
names(offerBehavior) <- c("ASSET_ID", "NAME", "Desired10", "Reason")





# compare list against bidding behaviors
finaloutput <- merge(offerBehavior, Query_Results_1, by.x = "ASSET_ID", by.y = "ASSET_ID", all.x = TRUE)
bidding <- finaloutput
finaloutput <- finaloutput[!is.na(finaloutput$LOCAL_DAY), ]
finaloutput <- finaloutput %>% select(1:4)
finaloutput <- unique(finaloutput[, 1:4])
bidding <- bidding[is.na(bidding$LOCAL_DAY), ]

bidding$OfferStatus <- "Not bidding as a fast start"
bidding <- bidding %>% select(1:4,14)



# compare list created as a result from checking bidding behaviors against CROW outages

finaloutput <- merge(finaloutput, Query_Results_2, by.x = "NAME", by.y = "EQUIPMENT_NAME_ID", all.x = TRUE)
finaloutput <- finaloutput[is.na(finaloutput$OUTAGE_REQUEST_ID), ]
finaloutput <- finaloutput[, c("NAME", "ASSET_ID", "Desired10", "Reason"), drop = F]
finaloutput <- finaloutput[order(finaloutput$NAME), ]
finaloutput$Desired10 <- floor(finaloutput$Desired10)
finaloutput <- unique(finaloutput[, 1:4])



# rename the columns for friendly use in report
names(finaloutput) <- c("Resource Short Name", "Asset ID", "Target MW", "Reason")

# generate report only if the finaloutput has generated a list

xlFileName <- paste0(outputlocation, "ISO Initiated Audit List - ", format(currDate, "%Y-%m-%d %H-%M-%S"), ".xlsx")
# xlFileName2 <- paste0(outputlocation2, "ISO Initiated Audit List - ", format(currDate, "%Y-%m-%d %H-%M-%S"),".xlsx")

rowcheck <- nrow(finaloutput)
if (rowcheck > 0) {
  write.xlsx(x = finaloutput, file = xlFileName, sheetName = "Audit_List", col.names = TRUE, row.names = FALSE)
  # write.xlsx(x=finaloutput, file=xlFileName2, sheetName = "Audit_List", col.names=TRUE, row.names = FALSE)
  msgBox <- tkmessageBox(
    title = "All done!",
    message = "A file was created and saved to //isofilpd1/tso/Claim 10_30 Auditing/ISO Initiated C10 Audits", icon = "info", type = "ok"
  )
} else {
  msgBox <- tkmessageBox(
    title = "All done!",
    message = "There were no units that meet the auditing criteria this month. No report was generated.", icon = "info", type = "ok"
  )
}

xlFileName2 <- paste0(outputlocation, "Assets selected for ISO Audit not offering FS - ", format(currDate, "%Y-%m-%d %H-%M-%S"), ".xlsx")
# xlFileName2 <- paste0(outputlocation2, "ISO Initiated Audit List - ", format(currDate, "%Y-%m-%d %H-%M-%S"),".xlsx")



rowcheck2 <- nrow(bidding)
if (rowcheck2 > 0) {
  write.xlsx(x = bidding, file = xlFileName2, sheetName = "Non FS Bidding", col.names = TRUE, row.names = FALSE)
  # write.xlsx(x=finaloutput, file=xlFileName2, sheetName = "Audit_List", col.names=TRUE, row.names = FALSE)
  msgBox <- tkmessageBox(
    title = "Processing complete.",
    message = "A file was created with units that are not offering Fast Start and saved to //isofilpd1/tso/Claim 10_30 Auditing/ISO Initiated C10 Audits", icon = "info", type = "ok"
  )
} else {
  msgBox <- tkmessageBox(
    title = "Processing complete.",
    message = "All units are offering Fast Start! No report was generated.", icon = "info", type = "ok"
  )
}
# If script does not seem to finish, look for the pop-up box - it might be under another window!
# __________________________________________________________________________END
