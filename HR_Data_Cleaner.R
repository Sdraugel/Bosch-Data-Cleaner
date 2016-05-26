##Name: Data Cleaner
##Developer: Steven Draugel
##Description: Combines cleans, and formats pager data for the finance department
##01/26/2016

#Points the rJava library to the JVM location. Uncomment the one you need
#x64
Sys.setenv(JAVA_HOME='C:\\Program Files\\Java\\jre8')
#x86
#Sys.setenv(JAVA_HOME='C:\\Program Files (x86)\\Java\\jre7')

###Install block -- Uncomment the below lines to install necassary packages automatically
##Installs the necessary packages to format the data
install.packages("data.table")
install.packages("dplyr")
install.packages("xlsx")
install.packages("XLConnect")
install.packages("stringr")


#Import Statements
library(data.table)
library(dplyr)
library(tcltk)
library(xlsx)
library(tcltk)
library(tcltk2)
library(stringr)

###########################################
#Start of data cleaning
###########################################

#Set the directory where the data is stored
#####Put the folder name below#####
directory = tk_choose.dir(default = "", caption = "Select the working directory")

setwd(directory)

Last_Month_Report = tk_choose.files(default = "", caption = "Please select last month's Breakdown report", filters = NULL, index = 1)

HR_Info = tk_choose.files(default = "", caption = "Please select a current HR report", filters = NULL, index = 1)

Last_Month_Report_xlsx <- read.xlsx(Last_Month_Report[1], 1)
HR_Info_xlsx = read.xlsx(HR_Info[1], 1)

df_x = as.data.frame(Last_Month_Report_xlsx)
df_y = as.data.frame(HR_Info_xlsx)

write.csv(df_x, "Old.csv")
write.csv(df_y, "HR.csv")

#rm(Last_Month_Report,HR_Info,Last_Month_Report_xlsx,HR_Info_xlsx,df_x,df_y)

#Read in and change .txt files into CSV variables
OLD = read.csv("Old.csv", stringsAsFactors = FALSE)
USE = read.csv(tk_choose.files(default = "", caption = "Please select a current the usage report .txt", filters = NULL, index = 1), stringsAsFactors = FALSE)
#USE = read.csv("USAGE.txt", stringsAsFactors = FALSE)
DET = read.csv(tk_choose.files(default = "", caption = "Please select a current the detail report .txt", filters = NULL, index = 1), stringsAsFactors = FALSE)
HR = read.csv("HR.csv", stringsAsFactors = FALSE)

#Choose the files to run
#OLD = read.csv(file.choose(tk_choose.files(caption = "Select an old report", multi = FALSE)))
#USE = read.csv(file.choose(tk_choose.files(caption = "Select the usage file", multi = FALSE)))
#DET = read.csv(file.choose(tk_choose.files(caption = "Select the new detail file", multi = FALSE)))
#HR = read.csv(file.choose(tk_choose.files(caption = "Select the HR reference", multi = FALSE)))

#Format the csv variables into dataframes
df_OLD = data.frame(OLD)
df_USE = data.frame(USE)
df_DET = data.frame(DET)
df_HR = data.frame(HR)

#Convert columns to string
df_USE$Holder <- as.character(df_USE$Holder)
df_DET$Holder <- as.character(df_DET$Holder)
df_HR$Empl..Appl.Name <- as.character(df_HR$Empl..Appl.Name)
  
#The columns to be kept from the original data
keeps = c(13, 14, 34, 49, 50)

#Unwanted columns removed
df_USE <- subset(df_USE, select = keeps, stringsAsFactors = FALSE)
df_DET <- subset(df_DET, select = keeps, stringsAsFactors = FALSE)
df_HR <- subset(df_HR, select = c(1,3), stringsAsFactors = FALSE)

#Fixed the change in names from HR
holders <- as.data.frame(str_split_fixed(df_HR$Empl..Appl.Name, ",", 2))
new_holders <- as.data.frame(str_split_fixed(holders$V2, " ", 3))
holders <- as.data.frame(paste(holders$V1, new_holders$V2, sep = ", "))
colnames(holders)[1] <- "ReferenceField3"
df_HR$Empl..Appl.Name <- holders$ReferenceField3

#Data frames combined into one data frame
Combined_Data = rbind(df_DET, df_USE)

#Convert all data in column TotalAmount to an int
Combined_Data$TotalAmount <- as.numeric(Combined_Data$TotalAmount)
Combined_Data$PagerPhoneNumber <- as.numeric(Combined_Data$PagerPhoneNumber)

#Remove all rows where PagerPhoneNumber is 0
Combined_Data$PagerPhoneNumber[Combined_Data$PagerPhoneNumber == 0 & Combined_Data$TotalAmount > 0] <- "Finance"
Combined_Data$Holder[Combined_Data$PagerPhoneNumber == "Finance"] <- "Finance Charge"
Combined_Data$ReferenceField2[Combined_Data$PagerPhoneNumber == "Finance"] <- "Finance"
Combined_Data$ReferenceField3[Combined_Data$PagerPhoneNumber == "Finance"] <- "Finance"
Clean_Data = setDT(Combined_Data)[, if(!any(PagerPhoneNumber == 0)) .SD, by = PagerPhoneNumber]

#Removes the old data objects to clear working memory
#rm(df_DET, df_USE, DET, Combined_Data, keeps, HR)

#Change df_HR column name Empl..Appl.Name to Holder 
colnames(df_HR)[1] <- "ReferenceField3"
df_HR$ReferenceField3 <- as.character(df_HR$ReferenceField3)
df_HR$ReferenceField3 <- as.numeric(df_HR$ReferenceField3)
Clean_Data$ReferenceField3 <- as.character(Clean_Data$ReferenceField3)
Clean_Data$ReferenceField3 <- as.numeric(Clean_Data$ReferenceField3)

#Merging the use data with the reference data
merge_df = merge(x = Clean_Data, y = df_HR, by = "ReferenceField3", all=TRUE)
colnames(merge_df)[6] = "Actual.Cost.Center"

#Adding cost center to all 9999999 personnel numbers
merge_df$Actual.Cost.Center[merge_df$ReferenceField3 == 999999] <- 927347
merge_df$Actual.Cost.Center[merge_df$ReferenceField3 >= 92000000] <- 927347
merge_df$Actual.Cost.Center[merge_df$ReferenceField3 == 9999999] <- 927347

#Reorganize the columns for correct output
subset_df <- subset(merge_df, select = c(3,2,4,5,1,6))
subset_df <- subset_df[order(Actual.Cost.Center),]

#Remove all rows where PagerPhoneNumber is 0
subset_df = setDT(subset_df)[, if(!any(is.na(PagerPhoneNumber))) .SD, by = PagerPhoneNumber]

#Removing the unneeded columns from the old report
subset_old <- subset(df_OLD, select = c("Holder", "ReferenceField2", "ReferenceField3", "Actual.Cost.Center"))
subset_old$Actual.Cost.Center <- sub(" .*","",subset_old$Actual.Cost.Center)
subset_old <- setDT(subset_old)[, if(!any(is.na(Holder))) .SD, by = Holder]

field_check <- ""
field_check[1] <- as.data.frame(df_OLD$ReferenceField2)
field_check[2] <- as.data.frame(df_OLD$ReferenceField3)
field_check[3] <- as.data.frame(df_OLD$Actual.Cost.Center)
field_check <- as.data.frame(field_check)
colnames(field_check)[1] <- "ReferenceField2"
colnames(field_check)[2] <- "ReferenceField3"
colnames(field_check)[3] <- "Actual.Cost.Center"

HR$Empl..Appl.Name <- sub(",,", ",", HR$Empl..Appl.Name)

subset_df$Holder <- sub(" ", ", ", toupper(subset_df$Holder))
subset_df$Holder <- sub(",,", ",", subset_df$Holder)

#~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~Good to here~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#Add the correct ReferenceField3 and Actual.Cost.Center from last months report
#subset_df$ReferenceField3 <- subset_old$ReferenceField3[match(subset_df$Holder, subset_old$Holder)] ####This is where the problem is!!###
subset_df$Actual.Cost.Center <- subset_old$Actual.Cost.Center[match(subset_df$Holder, subset_old$Holder)]
subset_df$ReferenceField3 <- field_check$ReferenceField3[match(subset_df$ReferenceField2, field_check$ReferenceField2)]
subset_df$Actual.Cost.Center <- field_check$Actual.Cost.Center[match(subset_df$ReferenceField2, field_check$ReferenceField2)]
subset_df[is.na(subset_df)] <- "0"
#subset_df$Actual.Cost.Center <- HR$Cost.Center[match(subset_df$Holder, HR$Empl..Appl.Name)]

#Creates a table with totals for each Actual.Cost.Center
totals <- aggregate(subset_df$TotalAmount, by=list(Actual.Cost.Center = subset_df$Actual.Cost.Center), FUN=sum)
colnames(totals)[2] = "Totals"
totals$Totals <- as.numeric(as.character(as.factor(totals$Totals)))
totals$Actual.Cost.Center <- as.numeric(as.character(as.factor(totals$Actual.Cost.Center)))

#Get a grand total for the total cost
GrandTotal <- colSums(totals, na.rm = FALSE)
GrandTotal <- as.data.frame(GrandTotal)

#remove the uneeded row
GrandTotal <- GrandTotal[-1,]
GrandTotal <- as.data.frame(GrandTotal)
colnames(GrandTotal)[1] = "Totals"

#Merge totals and subset_df together
merge_subset_df <- rbind(subset_df, totals, fill = TRUE)
merge_subset_df <- rbind(merge_subset_df, GrandTotal, fill = TRUE)

#Reorder the output so that the columns are in the correct order and totaled rows are in the correct place
Output <- subset(merge_subset_df, select = c(3,2,4,5,1,6,7))
Output <- Output[order(-Actual.Cost.Center, PagerPhoneNumber),]
Output <- as.data.frame(Output)
colnames(Output)[7] = "Totals"

#Remove uneeded data sets to free up RAM
#rm(merge_subset_df, subset_df, totals, subset_old, GrandTotal)

#Add the string Total to the last row for each Actual.Cost.Center
Output$Actual.Cost.Center[is.na(Output$Holder)] <- "TOTAL"
Output$Actual.Cost.Center[Output$Actual.Cost.Center == 123456] <- "Finance"
Output <- sapply(Output, as.character)
Output <- as.data.frame(Output)
Output$Actual.Cost.Center <- as.character(Output$Actual.Cost.Center)
Output$Actual.Cost.Center[nrow(Output)] <- "GRAND TOTAL"

Final <- subset(Output, select = c("Holder", 
                                   "PagerPhoneNumber", 
                                   "TotalAmount", 
                                   "ReferenceField2", 
                                   "ReferenceField3", 
                                   "Actual.Cost.Center",
                                   "Totals"))

Final <- as.data.frame(Final)

#Set each column to the correct data type
Final$Holder <- as.character(Final$Holder)
Final$PagerPhoneNumber <- as.character(Final$PagerPhoneNumber)
Final$TotalAmount <- as.numeric(as.character(Final$TotalAmount))
Final$ReferenceField2 <- as.character(Final$ReferenceField2)
Final$ReferenceField3 <- as.character(Final$ReferenceField3)
Final$Actual.Cost.Center <- as.numeric(as.character(Final$Actual.Cost.Center))
Final$Totals <- as.numeric(as.character(Final$Totals))
Final$Actual.Cost.Center[Final$ReferenceField3 == "Finance"] <- "Finance"

Final$Actual.Cost.Center[is.na(Output$Holder)] <- "TOTAL"
Final$Actual.Cost.Center[nrow(Output)] <- "GRAND TOTAL"

#Get rid of NA
Final[is.na(Final)] <- ""

colnames(Final)[6] = "Actual Cost Center"

#Final$`Actual Cost Center` <- Final$`Actual Cost Center`[match(subset_df$Holder, subset_old$Holder)]

#Creates a XLSX file and writes Final to it
write.xlsx(Final, file="Output Report.xlsx", row.names = FALSE)

############################################
#End of data cleaning
###########################################

#rm(FIN, Final, Output)

###########################################
#Start of Excel formatting
###########################################
#Read in the cleaned data
data <- read.xlsx("Output Report.xlsx", 1, header=TRUE, colClasses = NA)
data <- as.data.frame(data)

#Create work book objects
wb <- createWorkbook(type="xlsx")
sheet <- createSheet(wb, sheetName = "DETAIL")
rows <- getRows(sheet)
cells <- getCells(rows)

#Deifing the cell style and table column style
CellStyle(wb, dataFormat = NULL, alignment = NULL, border = NULL, fill = NULL, font = NULL)
TABLE_COLNAMES_STYLE <- CellStyle(wb) + Font(wb, heightInPoints = 10, name = "Arial", isBold=TRUE) +
  Alignment(wrapText=TRUE, horizontal="ALIGN_CENTER") +
  Border(color="black", position=c("TOP", "BOTTOM"), 
         pen=c("BORDER_THIN", "BORDER_THICK"))

#Write the cleaned data to the workbook
addDataFrame(data, sheet, row.names = FALSE, colnamesStyle = TABLE_COLNAMES_STYLE)
#attributes(wb)$row.names <- NULL

#Auto resize the columns in the excel document
autoSizeColumn(sheet, colIndex = c(1:ncol(data)))

subset_use <- subset(USE, select = c("Holder", "TotalPages"))

usage <- createSheet(wb, sheetName = "USAGE")
rows_usage <- getRows(usage)
cells_usage <- getCells(rows_usage)

addDataFrame(subset_use, usage, row.names = FALSE, colnamesStyle = TABLE_COLNAMES_STYLE)

autoSizeColumn(usage, colIndex = c(1:ncol(USE)))



#rm(data)

#Save and output the finished report
saveWorkbook(wb, "Finished_Report.xlsx")
#rm(list=ls(all=TRUE))




















