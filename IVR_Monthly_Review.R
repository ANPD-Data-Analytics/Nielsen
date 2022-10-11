# Future updates - Adjust Barcodes in IVR Files to have leading zeros

#sink("//FCRPDFile02/WRKStApps1$/Retail Sales/Retail Reports/8F/log.txt", append=FALSE, split=FALSE)

# Load Libraries ####
library(readxl)
library(tidyverse)
library(rlist)
library(xlsx)
library(openxlsx)
library(odbc)
library(readtext)
library(compareDF)
library(pryr)
#library(rJava)
#install.packages('lambda.r')

# Define Username / Password ####
user = Sys.getenv("username")

# Define Username/ Password ####
UID <- readtext(paste0("C:/Users/",user,"/Login_Info/UID_RetailVelocity.txt"))$text
PWD <- readtext(paste0("C:/Users/",user,"/Login_Info/PWD_RetailVelocity.txt"))$text

# Connect to SQL Sandbox Data ####
odbcConnStr <- paste0("Driver={ODBC Driver 17 for SQL Server};
                 Server=tcp:abt-sqlmi03-test.public.14690b0458e7.database.windows.net,3342;
                 database=AbbottSandbox;
                 Uid={", UID, "};
                 Pwd={", PWD, "};
                 Encrypt=yes;
                 TrustServerCertificate=no;
                 Connection Timeout=1000;
                 Authentication=ActiveDirectoryPassword;")

# Function to query the sandbox ####
qrySandbox <- function(sqlText1) {
  con <- odbc::dbConnect(odbc::odbc(), 
                         .connection_string = odbcConnStr)
  outputData <- odbc::dbGetQuery(con, sqlText1)
  return(outputData)
}

# Define Today's Date ####
date <- Sys.Date()

# ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~####

# Define IVR and IVR_prev files in the directory and name as below ####
file_path1 <- "//FCRPDFile02/WRKStApps1$/Retail Sales/Retail Reports/8F/Connect Validation/Items and Values_LN.xlsx"
file_path2 <- "//FCRPDFile02/WRKStApps1$/Retail Sales/Retail Reports/8F/Connect Validation/Items and Values_LN_Prev.xlsx"
file_path3 <- "//FCRPDFile02/WRKStApps1$/Retail Sales/Retail Reports/8F/Connect Validation/Items and Values_ITN.xlsx"
file_path4 <- "//FCRPDFile02/WRKStApps1$/Retail Sales/Retail Reports/8F/Connect Validation/Items and Values_ITN_Prev.xlsx"

# Load the IVR files ####
LN <- read_excel(file_path1)
LN_Prev <- read_excel(file_path2)
ITN <- read_excel(file_path3)
ITN_Prev <- read_excel(file_path4)

# ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~####

# LN IVR Exceptions      ####

LN1 <- LN %>% 
  select(`BARCODE`, 
         `#US LOC PRO ITEM DESCRIPTION`, 
         `Calc EQ`,
         `ABT_MEGA CATEGORY`,
         `ABT_CATEGORY`,
         `ABT_SEGMENT`,
         `ABT_SUBSEGMENT`,
         `ABT_BRAND`,
         `ABT_SUBBRAND`,
         `#US LOC BRAND`,
         `ABT_MANUFACTURER`,
         `ABT_FLAVOR`,
         `ABT_FLAVOR GROUP`,
         #`ABT_FLAVOR DETAIL`,
         `ABT_FORM`,
         `ABT_FORM TYPE`,
         `ABT_COUNT SIZE`,
         #`ABT_COUNT GROUP (FORM)`,
         `ABT_PACK`,
         #`ABT_HYDRATION`,
         #`ABT_INGREDIENT`,
         `ABT_DIET GROUP`,
         `#US LOC BASE SIZE`,
         `ABT_SIZE (LIQUID)`,
         `ABT_CONTAINER (LIQUID)`,
         `#US LOC FLAVOR`,
         `#US LOC FORM`,
         `#US LOC DERIVED BRAND OWNER HIGH`,
         `#US LOC BRAND OWNER`,
         `#US LOC COMPARE TO CLAIM`,
         `#US LOC IM MULTI CHAR`,
         `#US LOC ORGANIC CLAIM`,
         `#US LOC PACKAGE GENERAL SHAPE`,
         `#US LOC SERVING PER CONTAINER`,
         #`#US LOC PI KETOGENIC DIET`
  )         

LN_Prev1 <- LN_Prev %>% 
  select(`BARCODE`, 
         `#US LOC PRO ITEM DESCRIPTION`, 
         `Calc EQ`,
         `ABT_MEGA CATEGORY`,
         `ABT_CATEGORY`,
         `ABT_SEGMENT`,
         `ABT_SUBSEGMENT`,
         `ABT_BRAND`,
         `ABT_SUBBRAND`,
         `#US LOC BRAND`,
         `ABT_MANUFACTURER`,
         `ABT_FLAVOR`,
         `ABT_FLAVOR GROUP`,
         #`ABT_FLAVOR DETAIL`,
         `ABT_FORM`,
         `ABT_FORM TYPE`,
         `ABT_COUNT SIZE`,
         #`ABT_COUNT GROUP (FORM)`,
         `ABT_PACK`,
         #`ABT_HYDRATION`,
         #`ABT_INGREDIENT`,
         `ABT_DIET GROUP`,
         `#US LOC BASE SIZE`,
         `ABT_SIZE (LIQUID)`,
         `ABT_CONTAINER (LIQUID)`,
         `#US LOC FLAVOR`,
         `#US LOC FORM`,
         `#US LOC DERIVED BRAND OWNER HIGH`,
         `#US LOC BRAND OWNER`,
         `#US LOC COMPARE TO CLAIM`,
         `#US LOC IM MULTI CHAR`,
         `#US LOC ORGANIC CLAIM`,
         `#US LOC PACKAGE GENERAL SHAPE`,
         `#US LOC SERVING PER CONTAINER`,
         #`#US LOC PI KETOGENIC DIET`
  ) 

# LN Full File ####
LNIVRFULL <- LN1

# Find items in one and not the other ####
UPCsInBoth <- intersect(LN1$`BARCODE`, LN_Prev1$`BARCODE`)

newUPC <- LN1 %>% filter(!LN1$`BARCODE` %in% UPCsInBoth) %>% arrange(`ABT_MANUFACTURER`)
UPCdropout <- LN_Prev1 %>% filter(!LN_Prev1$`BARCODE` %in% UPCsInBoth)

LN1 <- LN1 %>% filter(LN1$`BARCODE` %in% UPCsInBoth) %>% arrange(`ABT_MANUFACTURER`)
LN_Prev1 <- LN_Prev1 %>% filter(LN_Prev1$`BARCODE` %in% UPCsInBoth)

#round EQ Factor once added to sample files
#LN1$`EQ Calc` <- round(as.numeric(LN1$`EQ Calc`), 4)
#LN_Prev1$`EQ Calc` <- round(as.numeric(LN_Prev1$`EQ Calc`), 4)

# All other differences ####
outputdf <- NULL
thekeylist <- "BARCODE"
nameslist <- names(LN_Prev1)
namesdiff <- nameslist[nameslist != thekeylist]
for (i in namesdiff)
{
  thekeylistcp <- thekeylist
  vars <- list.append(thekeylistcp, i)
  df_subseta <- LN1 %>% select(vars)
  df_subsetb <- LN_Prev1 %>% select(vars)
  #converting to char avoid type conversion issues
  df_subseta[] <- as.data.frame(lapply(df_subseta, as.character))
  df_subsetb[] <- as.data.frame(lapply(df_subsetb, as.character))
  diffs1 <- anti_join(df_subseta, df_subsetb, by = vars)
  diffs2 <- anti_join(df_subsetb, df_subseta, by = vars)
  outputdftemp <- diffs1
  outputdftemp <- bind_rows(outputdftemp, diffs2)
  #key, field, new, old
  names(df_subseta) <- list.append(thekeylist, "a_val")
  names(df_subsetb) <- list.append(thekeylist, "b_val")
  outputdftemp <- outputdftemp %>% select(thekeylist) %>% mutate(fieldchanged = i)
  outputdftemp <- left_join(outputdftemp, df_subseta, by = thekeylist)
  outputdftemp <- left_join(outputdftemp, df_subsetb, by = thekeylist)
  if(!is.null(outputdf)){
    outputdf <- bind_rows(outputdf, outputdftemp)
  }else{
    outputdf <- outputdftemp
  }
}

rm(outputdftemp,i,diffs1,diffs2,df_subseta,df_subsetb)

names(outputdf) <- c("BARCODE", "Field Changed", "New", "Previous")

LN_Catalog_Changes <- outputdf %>% distinct(.) %>% arrange(`BARCODE`)


#Join on Current IVR values
JOIN_Current<- LN1[, c("BARCODE", "#US LOC PRO ITEM DESCRIPTION", "ABT_SUBSEGMENT")]
LN_Catalog_Changes1<-left_join(LN_Catalog_Changes,JOIN_Current,by="BARCODE")

# Create Workbook for IVR to IVR diffs with LN Catalog ####
wb <- openxlsx::createWorkbook()
openxlsx::addWorksheet(wb, "NEW LN IVR")
openxlsx::writeData(wb,
                    sheet = "NEW LN IVR",
                    as.data.frame(LNIVRFULL),
                    rowNames = FALSE)

wb <- openxlsx::createWorkbook()
openxlsx::addWorksheet(wb, "LN UPC Carryover")
openxlsx::writeData(wb,
                    sheet = "LN UPC Carryover",
                    as.data.frame(LN1),
                    rowNames = FALSE)

# Add tab for New LN UPCs
openxlsx::addWorksheet(wb, "New LN UPCs")
openxlsx::writeData(wb,
                    sheet = "New LN UPCs",
                    as.data.frame(newUPC),
                    rowNames = FALSE)

# Add tab for LN UPC Dropout
openxlsx::addWorksheet(wb, "LN UPC Dropout")
openxlsx::writeData(wb,
                    sheet = "LN UPC Dropout",
                    as.data.frame(UPCdropout),
                    rowNames = FALSE)

# Add tab for other LN Changes
openxlsx::addWorksheet(wb, "LN Other Changes")
openxlsx::writeData(wb,
                    sheet = "LN Other Changes",
                    as.data.frame(LN_Catalog_Changes1),
                    rowNames = FALSE)

# Save LN Changes Workbook - For IVR to IVR ####
openxlsx::saveWorkbook(wb, file = paste0("\\\\FCRPDFile02\\WRKStApps1$\\Retail Sales\\Retail Reports\\8F\\Connect Validation\\LNDiffs", date, ".xlsx"), overwrite = TRUE)

# ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~####

# LN - Compare Current Sandbox Item Master to New LN IVR ####

# Query Sandbox ####
# Adjust Column below to add fields ####
input <- qrySandbox("SELECT   
			                         IT.NIELSEN_UPC
			                        ,IT.NIELSEN_DESCRIPTION
                              ,IT.EQ_FACTOR
                              ,IT.NIELSEN_SEGMENT
			                        ,IT.NIELSEN_BRAND
FROM     AbbottSandbox.dbo.vvItem IT ")

# Sandbox Column options ####
# [IDX_ITEM]
# ,[ITEM_DESC]
# ,[NIELSEN_DESCRIPTION]
# ,[FULL_UPC]
# ,[NIELSEN_UPC]
# ,[EQ_FACTOR]
# ,[ABBOTT_PRODUCT]
# ,[AN_ITEM_NBR]
# ,[AN_BRAND_FAMILY]
# ,[AN_BRAND]
# ,[AN_SEGMENT]
# ,[AN_FORM]
# ,[AN_FLAVOR]
# ,[AN_WIC_REBATE_PRODUCT]
# ,[NIELSEN_ABT_MEGA_CATEGORY]
# ,[NIELSEN_ABT_SUBSEGMENT]
# ,[NIELSEN_ABT_MANUFACTURER]
# ,[NIELSEN_ABT_FORM]
# ,[NIELSEN_ABT_SIZE]
# ,[NIELSEN_ABT_REBATE]
# ,[NIELSEN_ABT_CONTAINER]
# ,[NIELSEN_ABT_FLAVOR]
# ,[NIELSEN_ABT_FLAVOR_GROUP]
# ,[NIELSEN_ABT_PACK]
# ,[NIELSEN_ABT_FORMULATION]
# ,[NIELSEN_MEGA_CATG]
# ,[NIELSEN_ABT_FORMULA_TYPE]
# ,[NIELSEN_CATEGORY]
# ,[NIELSEN_SEGMENT]
# ,[NIELSEN_SUBSEGMENT]
# ,[NIELSEN_MANUFACTURER]
# ,[NIELSEN_BRAND]
# ,[NIELSEN_SUBBRAND]
# ,[NIELSEN_BRANDLOW]
# ,[NIELSEN_FORM]
# ,[NIELSEN_PACK_SIZE]
# ,[NIELSEN_FLAVOR]
# ,[NIELSEN_FLAVOR_GROUP]
# ,[NIELSEN_ORGANIC_CLAIM]
# ,[NIELSEN_FORMULA_TYPE]
# ,[NIELSEN_FORMULASUBTYPE]
# ,[NIELSEN_FORMULATION]
# ,[NIELSEN_VARIANT]
# ,[NIELSEN_ABT_NAT_ORG]
# ,[EACHES_PER_SCANNED_ITEM]
# ,[EACHES_PER_CASE]
# ,[ABT_SIZE_RANGE]
# ,[VENDOR_CODE]
# ,[NIELSEN_ABT_BRAND_FAMILY]
# ,[NIELSEN_ABT_DIET_GROUP]
# ,[NIELSEN_ABT_BRAND_TYPE]
# ,[NIELSEN_ABT_SUBBRAND_GROUP]
# ,[NIELSEN_ABT_BASE_SIZE]
# ,[ABT_COUNT_SIZE]
# ,[ABT_MULTI]

# Save SQL data as DF for Comparison Item Master to IVR ####
# Adjust Sandbox Column name below to match IVR ####
CurrentItemValues <- input 
colnames(CurrentItemValues)<-c('BARCODE','#US LOC PRO ITEM DESCRIPTION','Calc EQ','ABT_SEGMENT','ABT_BRAND')

# Check Items that changed ####
# Select columns from IVR ####
NewCheckValues <- subset(LN1, select = c('BARCODE','#US LOC PRO ITEM DESCRIPTION','Calc EQ','ABT_SEGMENT','ABT_BRAND'))
CheckValuesBoth <- as.data.frame(intersect(NewCheckValues$`BARCODE`, CurrentItemValues$`BARCODE`))
colnames(CheckValuesBoth)<-c('BARCODE')

# Filter DFs to LN barcodes ####
NewCheckValues1 <- as.data.frame(NewCheckValues %>% filter(NewCheckValues$`BARCODE` %in% CheckValuesBoth$`BARCODE`))
CurrentItemValues1 <- as.data.frame(CurrentItemValues %>% filter(CurrentItemValues$`BARCODE` %in% CheckValuesBoth$`BARCODE`))

# Compare the data frames ####
CurrentVSNew <- compare_df(NewCheckValues1, CurrentItemValues1, c("BARCODE"),keep_unchanged_rows = FALSE, keep_unchanged_cols = TRUE)

# Save comparison file, Current to New Changes Workbook - For IVR to Sandbox ####
create_output_table(CurrentVSNew, output_type = 'xlsx', 
                    file_name = paste0("\\\\FCRPDFile02\\WRKStApps1$\\Retail Sales\\Retail Reports\\8F\\Connect Validation\\Changes2Current\\","LN Check Changes to Current ", date, ".xlsx"))

# ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~####

# ITN IVR Exceptions       ####

ITN1 <- ITN %>% 
  select(`BARCODE`, 
         `#US LOC PRO ITEM DESCRIPTION`, 
         `Calc EQ`,
         `ABT_MEGA CATEGORY`,
         `ABT_CATEGORY`,
         `ABT_SEGMENT`,
         `ABT_SUBSEGMENT`,
         `ABT_MANUFACTURER`,
         `ABT_BRAND FAMILY`,
         `ABT_BRAND`,
         `#US LOC BRAND`,
         `ABT_SUBBRAND`,
         `ABT_SUBBRAND GROUP`,
         `ABT_FORMULA TYPE`,
         `ABT_FORMULA SUBTYPE`,
         `ABT_FORMULA FORMULATION`,
         `ABT_VARIANT`,
         `ABT_PACK SIZE`,
         `ABT_MULTI`,
         `ABT_COUNT SIZE`,
         `ABT_PACK`,
         `ABT_FLAVOR`,
         `ABT_FORM`,
         `ABT_NATURAL/ORGANIC`,
         `ABT_FORMULA REB/NON REBATED`,
         `ABT_RCO`,
         `ABT_REFILL`,
         `ABT_OES TYPE`,
         `ABT_FIBER TN`,
         `ABT_SIZE RANGE`,
         `ABT_SIZE OVER 25`,
         `ABT_SIZE OVER 16`,
         `ABT_IF BRANDS EXCL SPCLTY`,
         `ABT_UP AGE INCL SPECIALTY`,
         `#US LOC BASE SIZE`,
         `#US LOC DERIVED BRAND OWNER HIGH`,
         `#US LOC COMPARE TO CLAIM`,
         `#US LOC FLAVOR`,
         `#US LOC FORM`,
         `#US LOC IM MULTI CHAR`,
         `#US LOC ORGANIC CLAIM`,
         `#US LOC PACKAGE GENERAL SHAPE`,
         `#US LOC SERVING PER CONTAINER`,
         `#US LOC YIELD`,
         `#US LOC BRAND OWNER`,
  )         

ITN_Prev1 <- ITN_Prev %>% 
  select(`BARCODE`, 
         `#US LOC PRO ITEM DESCRIPTION`, 
         `Calc EQ`,
         `ABT_MEGA CATEGORY`,
         `ABT_CATEGORY`,
         `ABT_SEGMENT`,
         `ABT_SUBSEGMENT`,
         `ABT_MANUFACTURER`,
         `ABT_BRAND FAMILY`,
         `ABT_BRAND`,
         `#US LOC BRAND`,
         `ABT_SUBBRAND`,
         `ABT_SUBBRAND GROUP`,
         `ABT_FORMULA TYPE`,
         `ABT_FORMULA SUBTYPE`,
         `ABT_FORMULA FORMULATION`,
         `ABT_VARIANT`,
         `ABT_PACK SIZE`,
         `ABT_MULTI`,
         `ABT_COUNT SIZE`,
         `ABT_PACK`,
         `ABT_FLAVOR`,
         `ABT_FORM`,
         `ABT_NATURAL/ORGANIC`,
         `ABT_FORMULA REB/NON REBATED`,
         `ABT_RCO`,
         `ABT_REFILL`,
         `ABT_OES TYPE`,
         `ABT_FIBER TN`,
         `ABT_SIZE RANGE`,
         `ABT_SIZE OVER 25`,
         `ABT_SIZE OVER 16`,
         `ABT_IF BRANDS EXCL SPCLTY`,
         `ABT_UP AGE INCL SPECIALTY`,
         `#US LOC BASE SIZE`,
         `#US LOC DERIVED BRAND OWNER HIGH`,
         `#US LOC COMPARE TO CLAIM`,
         `#US LOC FLAVOR`,
         `#US LOC FORM`,
         `#US LOC IM MULTI CHAR`,
         `#US LOC ORGANIC CLAIM`,
         `#US LOC PACKAGE GENERAL SHAPE`,
         `#US LOC SERVING PER CONTAINER`,
         `#US LOC YIELD`,
         `#US LOC BRAND OWNER`,
  )

# ITN Full File ####
ITNIVRFULL <- ITN1

# Find items in one and not the other ####
UPCsInBoth <- intersect(ITN1$`BARCODE`, ITN_Prev1$`BARCODE`)

newUPC <- ITN1 %>% filter(!ITN1$`BARCODE` %in% UPCsInBoth)%>% arrange(`ABT_MANUFACTURER`)
UPCdropout <- ITN_Prev1 %>% filter(!ITN_Prev1$`BARCODE` %in% UPCsInBoth)

ITN1 <- ITN1 %>% filter(ITN1$`BARCODE` %in% UPCsInBoth) %>% arrange(`ABT_MANUFACTURER`)
ITN_Prev1 <- ITN_Prev1 %>% filter(ITN_Prev1$`BARCODE` %in% UPCsInBoth)

#round EQ Factor once added to sample files
#ITN1$`EQ Calc` <- round(as.numeric(ITN1$`EQ Calc`), 4)
#ITN_Prev1$`EQ Calc` <- round(as.numeric(ITN_Prev1$`EQ Calc`), 4)

# All other differences ####
outputdf <- NULL
thekeylist <- "BARCODE"
nameslist <- names(ITN_Prev1)
namesdiff <- nameslist[nameslist != thekeylist]
for (i in namesdiff)
{
  thekeylistcp <- thekeylist
  vars <- list.append(thekeylistcp, i)
  df_subseta <- ITN1 %>% select(vars)
  df_subsetb <- ITN_Prev1 %>% select(vars)
  #converting to char avoid type conversion issues
  df_subseta[] <- as.data.frame(lapply(df_subseta, as.character))
  df_subsetb[] <- as.data.frame(lapply(df_subsetb, as.character))
  diffs1 <- anti_join(df_subseta, df_subsetb, by = vars)
  diffs2 <- anti_join(df_subsetb, df_subseta, by = vars)
  outputdftemp <- diffs1
  outputdftemp <- bind_rows(outputdftemp, diffs2)
  #key, field, new, old
  names(df_subseta) <- list.append(thekeylist, "a_val")
  names(df_subsetb) <- list.append(thekeylist, "b_val")
  outputdftemp <- outputdftemp %>% select(thekeylist) %>% mutate(fieldchanged = i)
  outputdftemp <- left_join(outputdftemp, df_subseta, by = thekeylist)
  outputdftemp <- left_join(outputdftemp, df_subsetb, by = thekeylist)
  if(!is.null(outputdf)){
    outputdf <- bind_rows(outputdf, outputdftemp)
  }else{
    outputdf <- outputdftemp
  }
}

rm(outputdftemp,i,diffs1,diffs2,df_subseta,df_subsetb)

names(outputdf) <- c("BARCODE", "Field Changed", "New", "Previous")

ITN_Catalog_Changes <- outputdf %>% distinct(.) %>% arrange(`BARCODE`)

#Join on Current IVR values
JOIN_Current<- ITN1[, c("BARCODE", "#US LOC PRO ITEM DESCRIPTION", "ABT_SUBSEGMENT")]
ITN_Catalog_Changes1<-left_join(ITN_Catalog_Changes,JOIN_Current,by="BARCODE")

# Identify duplicate UPCs between ITN and LN
Nielsen_dupes <- data.frame(intersect(ITN$BARCODE, LN$BARCODE))
colnames(Nielsen_dupes)<-c('BARCODE')

# Create Workbook with ITN diffs IVR vs IVR ####
wb <- openxlsx::createWorkbook()
openxlsx::addWorksheet(wb, "NEW ITN IVR")
openxlsx::writeData(wb,
                    sheet = "NEW ITN IVR",
                    as.data.frame(ITNIVRFULL),
                    rowNames = FALSE)

openxlsx::addWorksheet(wb, "ITN UPC Carryover")
openxlsx::writeData(wb,
                    sheet = "ITN UPC Carryover",
                    as.data.frame(ITN1),
                    rowNames = FALSE)

# Add tab for New ITN UPCs
openxlsx::addWorksheet(wb, "New ITN UPCs")
openxlsx::writeData(wb,
                    sheet = "New ITN UPCs",
                    as.data.frame(newUPC),
                    rowNames = FALSE)

# Add tab for ITN UPC Dropout
openxlsx::addWorksheet(wb, "ITN UPC Dropout")
openxlsx::writeData(wb,
                    sheet = "ITN UPC Dropout",
                    as.data.frame(UPCdropout),
                    rowNames = FALSE)

# Add tab for other ITN Changes
openxlsx::addWorksheet(wb, "ITN Other Changes")
openxlsx::writeData(wb,
                    sheet = "ITN Other Changes",
                    as.data.frame(ITN_Catalog_Changes1),
                    rowNames = FALSE)

# Add tab for other ITN Changes
openxlsx::addWorksheet(wb, "UPC in ITN-LN")
openxlsx::writeData(wb,
                    sheet = "UPC in ITN-LN",
                    as.data.frame(Nielsen_dupes),
                    rowNames = FALSE)

# Save ITN IVR to IVR Changes Workbook ####
openxlsx::saveWorkbook(wb, file = paste0("\\\\FCRPDFile02\\WRKStApps1$\\Retail Sales\\Retail Reports\\8F\\Connect Validation\\ITNDiffs", date, ".xlsx"), overwrite = TRUE)

# ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~####

# ITN - Compare Sandbox Current Item Master to New IVR  ####

# Check Items that changed ####
NewCheckValues <- subset(ITN1, select = c('BARCODE','#US LOC PRO ITEM DESCRIPTION','Calc EQ','ABT_SEGMENT','ABT_BRAND'))
CheckValuesBoth <- as.data.frame(intersect(NewCheckValues$`BARCODE`, CurrentItemValues$`BARCODE`))
colnames(CheckValuesBoth)<-c('BARCODE')

# Filter DFs to ITN barcodes ####
NewCheckValues1 <- as.data.frame(NewCheckValues %>% filter(NewCheckValues$`BARCODE` %in% CheckValuesBoth$`BARCODE`))
CurrentItemValues1 <- as.data.frame(CurrentItemValues %>% filter(CurrentItemValues$`BARCODE` %in% CheckValuesBoth$`BARCODE`))

# Compare DFs ####
CurrentVSNew <- compare_df(CurrentItemValues1,NewCheckValues1, c("BARCODE"),keep_unchanged_rows = FALSE, keep_unchanged_cols = TRUE)

# Save comparison file - Current Sandbox Item Master vs New IVR Changes Workbook ####
create_output_table(CurrentVSNew, output_type = 'xlsx', 
                    file_name = paste0("\\\\FCRPDFile02\\WRKStApps1$\\Retail Sales\\Retail Reports\\8F\\Connect Validation\\Changes2Current\\","ITN Check Changes to Current ", date, ".xlsx"))

# ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~####

# Old Code ####
# UID <- readtext("C:/Users/galloax1/OneDrive - Abbott/LoginInfo/Sandbox_UID.txt")$text
# PWD <- readtext("C:/Users/galloax1/OneDrive - Abbott/LoginInfo/Sandbox_PWD.txt")$text
# 
# odbcConnStr <- paste0("Driver={ODBC Driver 17 for SQL Server};
#                  Server=tcp:abt-sqlmi03-test.public.14690b0458e7.database.windows.net,3342;
#                  database=AbbottSandbox;
#                  Uid={", UID, "};
#                  Pwd={", PWD, "};
#                  Encrypt=yes;
#                  TrustServerCertificate=no;
#                  Connection Timeout=1000;
#                  Authentication=ActiveDirectoryPassword;")
# 
# qrySandbox <- function(sqlText1) {
#   con <- odbc::dbConnect(odbc::odbc(), 
#                          .connection_string = odbcConnStr)
#   outputData <- odbc::dbGetQuery(con, sqlText1)
#   return(outputData)
# }
# 
# input <- qrySandbox("select	 
# 				it.ITEM_DESCRIPTION AS ITEM_DESC
# 				,it.U_NIELSEN_DESCRIPTION
# 				,it.U_FRIENDLY_DESCRIPTION
# 				,CAST(it.U_BARCODE AS varchar(12))  AS NIELSEN_UPC
# 				,it.U_FULL_BARCODE AS FULL_UPC
# 				,it.U_NIELSEN_EQ_FACTOR_ADULT
# 				,it.U_NIELSEN_EQ_FACTOR_ITN
# 				,it.U_NIELSEN_EQ
# 		,max(case when itatt.IDX_ITEM_ATTRIBUTE_CLASS =  99 then itatt.VALUE_NAME end)  as Nielsen_Segment
# 		,max(case when itatt.IDX_ITEM_ATTRIBUTE_CLASS =  169 then itatt.VALUE_NAME end)  as Nielsen_ABT_SubSegment
# 		,max(case when itatt.IDX_ITEM_ATTRIBUTE_CLASS =  98 then itatt.VALUE_NAME end)  as Nielsen_Brand
# 		,max(case when itatt.IDX_ITEM_ATTRIBUTE_CLASS = 160 then itatt.VALUE_NAME end)  as Nielsen_Sub_Brand		
# 		,max(case when itatt.IDX_ITEM_ATTRIBUTE_CLASS = 191 then itatt.VALUE_NAME end)  as Nielsen_Brand_Low
# 		,max(case when itatt.IDX_ITEM_ATTRIBUTE_CLASS = 104 then itatt.VALUE_NAME end)  as Nielsen_Form
# 		,max(case when itatt.IDX_ITEM_ATTRIBUTE_CLASS = 105 then itatt.VALUE_NAME end)  as Nielsen_Flavor
# 		 from	 AbbottMain.Abbott.ITEM it
# 		 left join 
# 
# 	(SELECT  ial.IDX_ITEM 
# 			,ial.IDX_ITEM_ATTRIBUTE_CLASS
# 			,iac.CLASS_NAME
# 			,IAV.VALUE_NAME 
# 		 
# 	 from	 AbbottMain.Abbott.Item_Attribute_Link ial
# 			,AbbottMain.Abbott.Item_Attribute_Class iac
# 			,AbbottMain.Abbott.Item_Attribute_Value iav
# 
# 	 Where	iav.IDX_ITEM_ATTRIBUTE_CLASS = ial.IDX_ITEM_ATTRIBUTE_CLASS
# 	 and	iac.IDX_ITEM_ATTRIBUTE_CLASS = ial.IDX_ITEM_ATTRIBUTE_CLASS
# 	 AND	ial.IDX_ITEM_ATTRIBUTE_VALUE = iav.IDX_ITEM_ATTRIBUTE_VALUE
# 	 AND	iac.IDX_ITEM_ATTRIBUTE_CLASS = iav.IDX_ITEM_ATTRIBUTE_CLASS) itatt
# 
# on  it.IDX_ITEM = itatt.IDX_ITEM
# group by it.ITEM_DESCRIPTION
# 				,it.U_NIELSEN_DESCRIPTION
# 				,it.U_FRIENDLY_DESCRIPTION
# 				,it.U_BARCODE
# 				,it.U_FULL_BARCODE
# 				,it.U_NIELSEN_EQ_FACTOR_ADULT
# 				,it.U_NIELSEN_EQ_FACTOR_ITN
# 				,it.U_NIELSEN_EQ")
# 
# 
# Item_Master_VQuery <- input %>% select(-U_NIELSEN_EQ) %>% select(NIELSEN_UPC, U_NIELSEN_DESCRIPTION, U_NIELSEN_EQ_FACTOR_ADULT, 
#                                                                  U_NIELSEN_EQ_FACTOR_ITN, Nielsen_Segment, Nielsen_ABT_SubSegment, 
#                                                                  Nielsen_Brand, Nielsen_Sub_Brand, Nielsen_Brand_Low, Nielsen_Form, 
#                                                                  Nielsen_Flavor) %>% 
#   filter(!NIELSEN_UPC == "000000000000") %>% distinct(.)
# #filter(!is.na(Nielsen_ABT_SubSegment)|!is.na(Nielsen_Brand)|!is.na(Nielsen_Sub_Brand)) %>% distinct(.)
# names(Item_Master_VQuery) <- c("Nielsen_UPC", "Nielsen_description", "Nielsen_EQ_Factor_Adult", "Nielsen_EQ_Factor_ITN", 
#                                "Nielsen_Segment", "Nielsen_ABT_SubSegment", "Nielsen_Brand", "Nielsen_Sub_Brand", "Nielsen_Brand_Low", 
#                                "Nielsen_Form", "Nielsen_Flavor")
# 
# 
# #compare item master to current 8F list
# VelocityUPC <- unique(Item_Master_VQuery$Nielsen_UPC)
# ITN_item_list <- intersect(ITN1$`BARCODE`, VelocityUPC)
# LN_item_list <- intersect(LN1$`BARCODE`, VelocityUPC)
# union_list <- union(ITN1$`BARCODE`, LN1$`BARCODE`)
# Overlap_item_list <- intersect(ITN1$`BARCODE`, LN1$`BARCODE`)
# 
# ITN_item_list <- ITN_item_list[which(!ITN_item_list %in% Overlap_item_list)]
# LN_item_list <- LN_item_list[which(!LN_item_list %in% Overlap_item_list)]
# 
# ITN_item_list_test <- setdiff(ITN1$`BARCODE`, VelocityUPC)
# LN_item_list_test <- setdiff(LN1$`BARCODE`, VelocityUPC)
# 
# inITNnotVelocity <- setdiff(ITN_item_list_test, VelocityUPC) %>% as_data_frame(.) %>% 
#   rename("Nielsen_UPC" = value) %>% left_join(.,ITN1, by = c("Nielsen_UPC" = "BARCODE"))
# inLNnotVelocity <- setdiff(LN_item_list_test, VelocityUPC) %>% as_data_frame(.) %>% 
#   rename("Nielsen_UPC" = value) %>% left_join(.,LN1, by = c("Nielsen_UPC" = "BARCODE"))
# 
# 
# Overlap_item_list <- Overlap_item_list %>% as_data_frame(Overlap_item_list) %>% rename("BARCODE" = value) 
# 
# Velocity_overlap <- Overlap_item_list %>% left_join(., Item_Master_VQuery, by = c("BARCODE" = "Nielsen_UPC"))
# 
# #need to left join other EQ and reorder & rename then search for differences
# 
# temp_ITN_data <- ITN1 %>% select(`BARCODE`, `Calc EQ`, "#US LOC PRO ITEM DESCRIPTION", `ABT_SEGMENT`, `ABT_SUBSEGMENT`, `ABT_BRAND`, 
#                                                `ABT_SUBBRAND`, `#US LOC BRAND`, ABT_FORM, ABT_FLAVOR)
# 
# temp_LN_data <- LN1 %>% select(`BARCODE`, `Calc EQ`, "#US LOC PRO ITEM DESCRIPTION", `ABT_SEGMENT`, `ABT_SUBSEGMENT`, `ABT_BRAND`, 
# `ABT_SUBBRAND`, `#US LOC BRAND`, ABT_FORM, ABT_FLAVOR)
# 
# 
# temp_ITN_data$BARCODE <- as.character(temp_ITN_data$BARCODE)
# temp_LN_data$BARCODE <- as.character(temp_LN_data$BARCODE)
# 
# temp_ITN_data$`Calc EQ` <- as.character(temp_ITN_data$`Calc EQ`)
# temp_LN_data$`Calc EQ` <- as.character(temp_LN_data$`Calc EQ`)
# 
# 
# Overlap_item_list <- Overlap_item_list %>%
#   left_join(.,temp_ITN_data) %>%
#   left_join(.,temp_LN_data)
# 
# 
# Overlap_item_list <- Overlap_item_list %>% select(`BARCODE`, `Calc EQ`, "#US LOC PRO ITEM DESCRIPTION", `ABT_SEGMENT`, `ABT_SUBSEGMENT`, `ABT_BRAND`, 
#                                `ABT_SUBBRAND`, `#US LOC BRAND`, ABT_FORM, ABT_FLAVOR)
# 
# ####
# ###Round in here ###
# ######
# 
# eqs <- input %>% select(NIELSEN_UPC, U_NIELSEN_EQ) %>% distinct(.)
# names(eqs) <- c("BARCODE", "Nielsen_EQ")
# 
# ############################################################# Above works
# #############################################################
# #############################################################
# #############################################################
# #############################################################
# 
# Velocity_overlap <- Velocity_overlap %>% select(`BARCODE`, Nielsen_description, Nielsen_Segment, Nielsen_ABT_SubSegment, Nielsen_Brand, 
#                                                 Nielsen_Sub_Brand, Nielsen_Brand_Low, Nielsen_Form, Nielsen_Flavor) %>%
#   left_join(., eqs)
# Velocity_overlap <- Velocity_overlap %>% select(`BARCODE`, Nielsen_EQ, Nielsen_description, Nielsen_Segment, Nielsen_ABT_SubSegment, Nielsen_Brand, 
#                                                 Nielsen_Sub_Brand, Nielsen_Brand_Low, Nielsen_Form, Nielsen_Flavor)
# names(Velocity_overlap) <- c("BARCODE", "Calc EQ", "ITEM DESCRIPTION", "ABT_SEGMENT(C)", "ABT_SUBSEGMENT(C)", "ABT_BRAND(C)", 
#                              "ABT_SUBBRAND(C)", "BRAND LOW", "FORM", "FLAVOR") 
# 
# names(Overlap_item_list) <- c("BARCODE", "Calc EQ", "ITEM DESCRIPTION", "ABT_SEGMENT(C)", "ABT_SUBSEGMENT(C)", "ABT_BRAND(C)", 
#                              "ABT_SUBBRAND(C)", "BRAND LOW", "FORM", "FLAVOR") 
# 
# #round EQ
# Overlap_item_list$`Calc EQ` <- round(as.numeric(Overlap_item_list$`Calc EQ`), 4)
# Velocity_overlap$`Calc EQ` <- round(as.numeric(Velocity_overlap$`Calc EQ`), 4)
# #differences
# outputdf <- NULL
# thekeylist <- "BARCODE"
# nameslist <- names(Overlap_item_list)
# namesdiff <- nameslist[nameslist != thekeylist]
# for (i in namesdiff)
# {
#   thekeylistcp <- thekeylist
#   vars <- list.append(thekeylistcp, i)
#   df_subseta <- Overlap_item_list %>% select(vars)
#   df_subsetb <- Velocity_overlap %>% select(vars)
#   #converting to char avoid type conversion issues
#   df_subseta[] <- as.data.frame(lapply(df_subseta, as.character))
#   df_subsetb[] <- as.data.frame(lapply(df_subsetb, as.character))
#   diffs1 <- anti_join(df_subseta, df_subsetb, by = vars)
#   diffs2 <- anti_join(df_subsetb, df_subseta, by = vars)
#   outputdftemp <- diffs1
#   outputdftemp <- bind_rows(outputdftemp, diffs2)
#   #key, field, new, old
#   names(df_subseta) <- list.append(thekeylist, "a_val")
#   names(df_subsetb) <- list.append(thekeylist, "b_val")
#   outputdftemp <- outputdftemp %>% select(thekeylist) %>% mutate(fieldchanged = i)
#   outputdftemp <- left_join(outputdftemp, df_subseta, by = thekeylist)
#   outputdftemp <- left_join(outputdftemp, df_subsetb, by = thekeylist)
#   if(!is.null(outputdf)){
#     outputdf <- bind_rows(outputdf, outputdftemp)
#   }else{
#     outputdf <- outputdftemp
#   }
# }
# 
# rm(outputdftemp,i,diffs1,diffs2,df_subseta,df_subsetb)
# 
# 
# names(outputdf) <- c("BARCODE", "Field Changed", "IVR", "Velocity")
# ITN1Desc <- ITN1 %>% select(`BARCODE`,`ITN ITEM DESCRIPTION` = "#US LOC PRO ITEM DESCRIPTION")
# LN1Desc <- LN1 %>% select(`BARCODE`,`LN ITEM DESCRIPTION` = "#US LOC PRO ITEM DESCRIPTION")
# Overlap_Out <-
#   outputdf %>% distinct(.) %>% arrange(`BARCODE`) %>% 
#   left_join(., ITN1Desc) %>% 
#   left_join(., LN1Desc) %>% mutate("ITN ITEM DESCRIPTION" = case_when(`Field Changed` == "ITEM DESCRIPTION"~`ITN ITEM DESCRIPTION`,
#                                                                                     TRUE~NA_character_),
#                                                  "LN ITEM DESCRIPTION" = case_when(`Field Changed` == "ITEM DESCRIPTION"~`LN ITEM DESCRIPTION`,
#                                                                                    TRUE~NA_character_),
#   )
# 
# 
# ITN1 <- as_data_frame(ITN_item_list) %>% rename("BARCODE" = value) %>% 
#   inner_join(., ITN1, by = "BARCODE")
# 
# LN1 <- as_data_frame(LN_item_list) %>% rename("BARCODE" = value) %>% 
#   inner_join(., LN1, by = "BARCODE")
# 
# outersect <- function(x, y) {
#   sort(c(x[!x%in%y],
#          y[!y%in%x]))
# }
# 
# #all the UPCs in Velocity not in the 8F
# unmatched_list <- outersect(union_list, VelocityUPC) %>% as_data_frame(.) %>% rename("Nielsen_UPC" = value)
# unmatched_list <- left_join(unmatched_list, Item_Master_VQuery)
# eqs <- input %>% select(NIELSEN_UPC, U_NIELSEN_EQ) %>% distinct(.)
# names(eqs) <- c("Nielsen_UPC", "Nielsen_EQ")
# desc <- input %>% select(NIELSEN_UPC, U_FRIENDLY_DESCRIPTION)
# names(desc) <- c("Nielsen_UPC", "Friendly_description")
# unmatched_list <- unmatched_list %>% select(Nielsen_UPC, Nielsen_description, Nielsen_Segment, Nielsen_ABT_SubSegment, Nielsen_Brand, 
#                                             Nielsen_Sub_Brand, Nielsen_Brand_Low, Nielsen_Form, Nielsen_Flavor) %>%
#   left_join(., eqs) %>% left_join(., desc)
# unmatched_list <- unmatched_list %>% select(Nielsen_UPC, Nielsen_description, Friendly_description, Nielsen_Segment, Nielsen_ABT_SubSegment, Nielsen_Brand, 
#                                             Nielsen_Sub_Brand, Nielsen_Brand_Low, Nielsen_Form, Nielsen_Flavor, Nielsen_EQ)
# 
# #otherwise compare either ln or itn
# #ITN
# ITN_item_list <- as_data_frame(ITN_item_list) %>% rename("Nielsen_UPC" = value)
# Item_Master_VQuery_ITN <- inner_join(Item_Master_VQuery, ITN_item_list) %>% 
#   select(Nielsen_UPC, Nielsen_description, Nielsen_EQ_Factor_ITN, Nielsen_Segment, Nielsen_ABT_SubSegment, 
#          Nielsen_Brand, Nielsen_Sub_Brand, Nielsen_Brand_Low, Nielsen_Form, Nielsen_Flavor)
# 
# ####### NEW CODE
# ITN2 <- ITN1 %>% select(BARCODE,"#US LOC PRO ITEM DESCRIPTION","Calc EQ",ABT_SEGMENT,ABT_SUBSEGMENT,ABT_BRAND,ABT_SUBBRAND,"#US LOC BRAND",ABT_FORM,ABT_FLAVOR)
# 
# names(Item_Master_VQuery_ITN) <- names(ITN2)
# #round EQ
# Item_Master_VQuery_ITN$`Nielsen_EQ_Factor_ITN` <- round(as.numeric(Item_Master_VQuery_ITN$`Nielsen_EQ_Factor_ITN`), 2)
# ITN2$`EQ Calc` <- round(as.numeric(ITN1$`Calc EQ`), 2)
# 
# #differences
# outputdf <- NULL
# thekeylist <- "BARCODE"
# nameslist <- names(Item_Master_VQuery_ITN)
# namesdiff <- nameslist[nameslist != thekeylist]
# for (i in namesdiff)
# {
#   thekeylistcp <- thekeylist
#   vars <- list.append(thekeylistcp, i)
#   df_subseta <- ITN1 %>% select(vars)
#   df_subsetb <- Item_Master_VQuery_ITN %>% select(vars)
#   #converting to char avoid type conversion issues
#   df_subseta[] <- as.data.frame(lapply(df_subseta, as.character))
#   df_subsetb[] <- as.data.frame(lapply(df_subsetb, as.character))
#   diffs1 <- anti_join(df_subseta, df_subsetb, by = vars)
#   diffs2 <- anti_join(df_subsetb, df_subseta, by = vars)
#   outputdftemp <- diffs1
#   outputdftemp <- bind_rows(outputdftemp, diffs2)
#   #key, field, new, old
#   names(df_subseta) <- list.append(thekeylist, "a_val")
#   names(df_subsetb) <- list.append(thekeylist, "b_val")
#   outputdftemp <- outputdftemp %>% select(thekeylist) %>% mutate(fieldchanged = i)
#   outputdftemp <- left_join(outputdftemp, df_subseta, by = thekeylist)
#   outputdftemp <- left_join(outputdftemp, df_subsetb, by = thekeylist)
#   if(!is.null(outputdf)){
#     outputdf <- bind_rows(outputdf, outputdftemp)
#   }else{
#     outputdf <- outputdftemp
#   }
# }
# 
# rm(outputdftemp,i,diffs1,diffs2,df_subseta,df_subsetb)
# 
# names(outputdf) <- c("BARCODE", "Field Changed", "IVR", "Velocity")
# ITNdf <- outputdf %>% distinct(.) %>% arrange(`BARCODE`)
# 
# 
# #LN
# 
# LN_item_list <- as_data_frame(LN_item_list) %>% rename("Nielsen_UPC" = value)
# Item_Master_VQuery_LN <- inner_join(Item_Master_VQuery, LN_item_list) %>% 
#   select(Nielsen_UPC, Nielsen_description, Nielsen_EQ_Factor_Adult, Nielsen_Segment, Nielsen_ABT_SubSegment, 
#          Nielsen_Brand, Nielsen_Sub_Brand, Nielsen_Brand_Low, Nielsen_Form, Nielsen_Flavor)
# 
# LN2 <- LN1 %>% select(BARCODE,"#US LOC PRO ITEM DESCRIPTION","Calc EQ",ABT_SEGMENT,ABT_SUBSEGMENT,ABT_BRAND,ABT_SUBBRAND,"#US LOC BRAND",ABT_FORM,ABT_FLAVOR)
# 
# names(Item_Master_VQuery_LN) <- names(LN2)
# #round EQ
# Item_Master_VQuery_LN$`Calc EQ` <- round(as.numeric(Item_Master_VQuery_LN$`Calc EQ`), 2)
# LN1$`Calc EQ` <- round(as.numeric(LN1$`Calc EQ`), 2)
# 
# #differences
# outputdf <- NULL
# thekeylist <- "BARCODE"
# nameslist <- names(Item_Master_VQuery_LN)
# namesdiff <- nameslist[nameslist != thekeylist]
# for (i in namesdiff)
# {
#   thekeylistcp <- thekeylist
#   vars <- list.append(thekeylistcp, i)
#   df_subseta <- LN1 %>% select(vars)
#   df_subsetb <- Item_Master_VQuery_LN %>% select(vars)
#   #converting to char avoid type conversion issues
#   df_subseta[] <- as.data.frame(lapply(df_subseta, as.character))
#   df_subsetb[] <- as.data.frame(lapply(df_subsetb, as.character))
#   diffs1 <- anti_join(df_subseta, df_subsetb, by = vars)
#   diffs2 <- anti_join(df_subsetb, df_subseta, by = vars)
#   outputdftemp <- diffs1
#   outputdftemp <- bind_rows(outputdftemp, diffs2)
#   #key, field, new, old
#   names(df_subseta) <- list.append(thekeylist, "a_val")
#   names(df_subsetb) <- list.append(thekeylist, "b_val")
#   outputdftemp <- outputdftemp %>% select(thekeylist) %>% mutate(fieldchanged = i)
#   outputdftemp <- left_join(outputdftemp, df_subseta, by = thekeylist)
#   outputdftemp <- left_join(outputdftemp, df_subsetb, by = thekeylist)
#   if(!is.null(outputdf)){
#     outputdf <- bind_rows(outputdf, outputdftemp)
#   }else{
#     outputdf <- outputdftemp
#   }
# }
# 
# rm(outputdftemp,i,diffs1,diffs2,df_subseta,df_subsetb)
# 
# names(outputdf) <- c("BARCODE", "Field Changed", "IVR", "Velocity")
# LNdf <- outputdf %>% distinct(.) %>% arrange(`BARCODE`)
# 
# 
# # output_list <- c("LNdf", "ITNdf", "unmatched_list")
# # rm(list=setdiff(ls(), output_list))
# 
# wb <- openxlsx::createWorkbook()
# openxlsx::addWorksheet(wb, "Items in ITN and LN")
# openxlsx::writeData(wb,
#                     sheet = "Items in ITN and LN",
#                     as.data.frame(Overlap_Out),
#                     rowNames = FALSE)
# openxlsx::addWorksheet(wb, "New ITN Items")
# openxlsx::writeData(wb,
#                     sheet = "New ITN Items",
#                     as.data.frame(inITNnotVelocity),
#                     rowNames = FALSE)
# openxlsx::addWorksheet(wb, "New LN Items")
# openxlsx::writeData(wb,
#                     sheet = "New LN Items",
#                     as.data.frame(inLNnotVelocity),
#                     rowNames = FALSE)
# openxlsx::addWorksheet(wb, "Changed ITN")
# openxlsx::writeData(wb,
#                     sheet = "Changed ITN",
#                     as.data.frame(ITNdf),
#                     rowNames = FALSE)
# openxlsx::addWorksheet(wb, "Changed LN")
# openxlsx::writeData(wb,
#                     sheet = "Changed LN",
#                     as.data.frame(LNdf),
#                     rowNames = FALSE)
# openxlsx::addWorksheet(wb, "Items in Velocity not IVR")
# openxlsx::writeData(wb,
#                     sheet = "Items in Velocity not IVR",
#                     as.data.frame(unmatched_list),
#                     rowNames = FALSE)
# openxlsx::saveWorkbook(wb, file = paste0("\\\\FCRPDFile02\\WRKStApps1$\\Retail Sales\\Retail Reports\\8F\\Connect Validation\\VelocityDiffs", date, ".xlsx"), overwrite = TRUE)
