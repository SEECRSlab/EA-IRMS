### INITIAL PROCESSING OF EA-IRMS CN DATA ###


# Script info -------------------------------------------------------------

# Original script by Miriam Gieske
# Jan. 9, 2020

# Modified by Stella Woeltjen (Pey) and Jess Gutknecht
# April 2023

# This script will extract the columns you need from the raw data, 
#   put them in the correct order,
#   and extract the data for the standards

# Each time you use this script to process a new batch of data,
#   you will need to change the following:
# The file path telling R which folder the data are in
# The names of the weigh sheet, EA, and IonVantage data files
# The name of the Excel file R saves at the end

# When switching from one batch of data to another, you should also clear the old
#   data by clicking the broom in the "Environment" window and clear the old code
#   by clicking in the "Console" window and hitting Ctrl+L.



# Before running code: ----------------------------------------------------
## Weight sheet and EA data should be copied into the IRMS batch file as additional tabs
## - Check weight sheet - ensure all rows have weight entered, even for blanks (enter 1 for blanks); 
##                      - ensure the correct description shows up in the "description" column for all rows
## - Check EA data - often, a line of zeros will be printed at bottom of EA data. This should usually be deleted.


#### LOAD PACKAGES #### (M. Gieske)

library(readxl) # Reads in .xls and .xlsx files
library(openxlsx) # Writes .xlsx files
library(dplyr) # Lets you rename columns (among other things)
library(ggplot2)
library(openxlsx)
library(tidyverse)


# Define folder and file names --------------------------------------------

### In this case there is an EA-IRMS_processing/Data folder that holds the IRMS output in .xls files
project_name <- "20230410_Gutknecht_AB_SARElegume2022_run3_raw" # copy/paste project name here
filepath <- here::here("data", "raw", project_name) # can also copy/paste filepath of folder here, if your script is not housed within a project
setwd("C:/Users/mars2199/Downloads/Data/Raw") # Tell R where to find the data (i.e. the file path)

file_name <- paste(project_name,".xls", sep="") # Check whether the file is .xlsx or .xls, and update this line accordingly

# Note: If you copy and paste the file path from the top of the Windows file explorer window, you will need to change the backslashes ("\")
# to forward slashes ("/").


# Import data --------------------------------------------------------

# Read in the weigh sheet
weigh_sheet <- readxl::read_excel(file_name, 
                                  sheet = "Weight Sheet",
                                  skip = 5, 
                                  col_names = c("Analysis order", "Tray position",  
                                                "Sample ID", "description", "Target wt (mg)", "Actual wt (mg)")) %>%
  .[!is.na(.$`Actual wt (mg)`), ] # Remove any rows where there is no weight data

#Read in the raw EA data from the VarioPYRO Cube and select the desired columns
ea_data <- readxl::read_excel(file_name, sheet = "EA data")

# Look in the Environment window (upper right).  You should see the two files that were read in.
# Check to make sure they have the same number of "observations" (rows)

# Read in the IonVantage data, select the desired columns, and rename them
# This chunk of code also removes a hidden row with alternative variable names and reorders the columns
N_data <- readxl::read_excel(file_name, sheet = "F1N2-report", skip = 12)
N1 <- N_data[-1 , ] %>%
  dplyr::select("Sample Number", "Name", "Height (nA)", "N15") %>%
  dplyr::rename("N analysis number" = 1, "Sample ID" = 2, "N peak height" = 3, "d15N vs At Air" = 4)

C_data <- readxl::read_excel(file_name, sheet = "F2CO2-report", skip = 12)
C1 <- C_data[-1 , ] %>%
  dplyr::select("Sample Number", "Name", "Height (nA)", "13C") %>%
  dplyr::rename("C analysis number" = 1, "Sample ID" = 2, "C peak height" = 3, "d13C vs VPDB" = 4)


# Extract the data you need -----------------------------------------------
# Combine the weigh sheet and EA data while dropping columns you don't need
raw1 <- cbind(weigh_sheet[ , c("Analysis order", "Sample ID", "description", "Actual wt (mg)")], 
              ea_data[ , c("Weight [mg]", "Name", "N [%]", "C [%]")]) %>%  # If this line of code throws an error, check that the column names are written here exactly as they appear in the dataframe
  dplyr::rename("%N" = 7, "%C" = 8) #rename columns 7 and 8

# Check whether the sample weights and descriptions line up. If all the values in the selected columns match up, it will say "TRUE" in the window below.
# "FALSE" means that one or more values do not match up.  You should look and decide whether the mismatch is a real problem.  (Hint: click on the file name in the environment
# window to open it so you can scroll through it). Reminder: sometimes when loading the data into varioPYRO, the sample ID is copy/pasted
# into varioPYRO and other times the description is copy/pasted.
identical(raw1$`Actual wt (mg)`, raw1$`Weight [mg]`) 
identical(raw1$description, raw1$Name)

# Check each IonVantage dataset to see whether there are any rows with no sample name. 
# That would indicate multiple peaks per sample.
# If there are any, look at the dataset to determine which peaks are valid
sum(is.na(N1$`Sample ID`)) # Number of rows with no name in the N dataset
sum(is.na(C1$`Sample ID`)) # Number of rows with no name in the C dataset

# Combine the IonVantage datasets
raw2 <- merge(N1, C1, by = "Sample ID", all = TRUE)

# The sample IDs in the IonVantage datasets have ".raw" added to the end
# Remove ".raw" so the IonVantage data can be merged with the other data
raw2$`Sample ID` <- sub(".raw", "", raw2$`Sample ID`)

# Combine the IonVantage data with the weigh sheet and EA data and organize dataframe.
# Get the data you want for the "Processed data" sheet in your Excel workbook and put the columns in the correct order
raw3 <- merge(raw1, raw2, by = "Sample ID") %>%
  dplyr::select("Analysis order", "Sample ID", "description", "Weight [mg]",  
                "N analysis number", "C analysis number", 
                "N peak height", "%N", "d15N vs At Air", 
                "C peak height", "%C", "d13C vs VPDB") %>% # select and order columns
  dplyr::rename("EA analysis order" = "Analysis order",
                d15N = "d15N vs At Air", 
                d13C = "d13C vs VPDB") %>% # rename "Analysis order", "d15N vs At Air" and "d13C vs VPDB column 
  mutate_at(c(1, 4:12), as.numeric) # mutate columns 1, 4:12 to numeric class


# Extract the standards and calculate quality control parameters ---------------------------------------------------

#Extract the standard data
peach <- raw3[grepl("Peach leaves", raw3$`Sample ID`, fixed=TRUE)==TRUE, ]
rsmt <- raw3[grepl("rsmt", raw3$`Sample ID`, fixed=TRUE)==TRUE, ]
acet <- raw3[grepl("ACETANILIDE", raw3$`Sample ID`, fixed=TRUE)==TRUE, ]
usgs <- raw3[grepl("USGS 40", raw3$`Sample ID`, fixed=TRUE)==TRUE, ]

# Order the standard data by the EA Analysis order column
peach <- peach[order(peach$'EA analysis order'),]
rsmt <- rsmt[order(rsmt$'EA analysis order'),]
acet <- acet[order(acet$'EA analysis order'),]
usgs <- usgs[order(usgs$'EA analysis order'),]

# Put all the ordered standards together.
stds <- rbind(peach, rsmt, acet, usgs)


#### calculate the R2, STD DEV, SLOPE, MEAN, EXPECTED VALUE, CALIBRATION VALUE (S. Pey, 3.24.2020) 

# Create a matrix to store the expected standard values, mean standard values and calculated
# calibration values for each of the four standards (Peach leaves, rsmt, acentanilide,
# and USGS 40)

qc <- matrix(NA, 24, 4)
row.names(qc) <- c("Peach leaves (r squared)", "Peach leaves (slope)", "Peach leaves (std dev)", 
                   "Peach leaves (mean)", "Peach leaves (expected)","Peach leaves (calibration)",
                   "rsmt (r squared)", "rsmt (slope)", "rsmt (std dev)", 
                   "rsmt (mean)", "rsmt (expected)","rsmt (calibration)",
                   "Acetanilide (r squared)", "Acetanilide (slope)", "Acetanilide (std dev)", 
                   "Acetanilide (mean)", "Aetenalide (expected)","Acetanilide (calibration)",
                   "USGS 40 (r squared)", "USGS 40 (slope)", "USGS 40 (std dev)", 
                   "USGS 40 (mean)", "USGS 40 (expected)","USGS 40 (calibration)")
colnames(qc) <- c("%N", "d15N", "%C", "d13C") 

## [PEACH LEAVES] Calculate the r-squared, standard deviation, slope and mean standard values 
## for %N, d15N, %C and d13C. Add the expected %N, d15N, %C and d13C values to the matrix.
## Calculate the calibration value and add it to the matrix.
qc[1, 1] <- summary(lm(`%N` ~ `N analysis number`, data = peach))$r.squared
qc[1, 2] <- summary(lm(d15N ~ `N analysis number`, data = peach))$r.squared
qc[1, 3] <- summary(lm(`%C` ~ `C analysis number`, data = peach))$r.squared
qc[1, 4] <- summary(lm(d13C ~ `C analysis number`, data = peach))$r.squared

qc[2, 1] <- summary(lm(`%N` ~ `N analysis number`, data = peach))$coefficients[2, 1]
qc[2, 2] <- summary(lm(d15N ~ `N analysis number`, data = peach))$coefficients[2, 1]
qc[2, 3] <- summary(lm(`%C` ~ `C analysis number`, data = peach))$coefficients[2, 1]
qc[2, 4] <- summary(lm(d13C ~ `C analysis number`, data = peach))$coefficients[2, 1]

qc[3, 1] <- sd(peach$`%N`)
qc[3, 2] <- sd(peach$d15N)
qc[3, 3] <- sd(peach$`%C`)
qc[3, 4] <- sd(peach$d13C)

qc[4, 1] <- mean(peach$`%N`)
qc[4, 2] <- mean(peach$d15N)
qc[4, 3] <- mean(peach$`%C`)
qc[4, 4] <- mean(peach$d13C)

qc[5,1] <- 2.28   # Expected %N, peach leaves
qc[5,2] <- NA     # Expected d15N, peach leaves
qc[5,3] <- 50.40  # Expected %C, peach leaves
qc[5,4] <- NA     # Expected d13C, peach leaves

qc[6, 1] <- qc[4, 1] / qc[5,1]  # %N calibration, peach leaves
qc[6, 2] <- NA                  # d15N calibration, peach leaves
qc[6, 3] <- qc[4, 3] / qc[5,3]  # %C calibration, peach leaves
qc[6, 4] <- NA                  # d13C calibration, peach leaves

## [Rosemount] Calculate the r-squared, standard deviation, slope and mean standard values 
## for %N, d15N, %C and d13C. Add the expected %N, d15N, %C and d13C values to the matrix.
## Calculate the calibration value and add it to the matrix.

qc[7, 1] <- summary(lm(`%N` ~ `N analysis number`, data = rsmt))$r.squared
qc[7, 2] <- summary(lm(d15N ~ `N analysis number`, data = rsmt))$r.squared
qc[7, 3] <- summary(lm(`%C` ~ `C analysis number`, data = rsmt))$r.squared
qc[7, 4] <- summary(lm(d13C ~ `C analysis number`, data = rsmt))$r.squared

qc[8, 1] <- summary(lm(`%N` ~ `N analysis number`, data = rsmt))$coefficients[2, 1]
qc[8, 2] <- summary(lm(d15N ~ `N analysis number`, data = rsmt))$coefficients[2, 1]
qc[8, 3] <- summary(lm(`%C` ~ `C analysis number`, data = rsmt))$coefficients[2, 1]
qc[8, 4] <- summary(lm(d13C ~ `C analysis number`, data = rsmt))$coefficients[2, 1]

qc[9, 1] <- sd(rsmt$`%N`)
qc[9, 2] <- sd(rsmt$d15N)
qc[9, 3] <- sd(rsmt$`%C`)
qc[9, 4] <- sd(rsmt$d13C)

qc[10, 1] <- mean(rsmt$`%N`)
qc[10, 2] <- mean(rsmt$d15N)
qc[10, 3] <- mean(rsmt$`%C`)
qc[10, 4] <- mean(rsmt$d13C)

qc[11,1] <- 0.21   # Expected %N, Rosemount soil
qc[11,2] <- NA     # Expected d15N, Rosemount soil
qc[11,3] <- 2.22   # Expected %C, Rosemount soil
qc[11,4] <- NA     # Expected d13C, Rosemount soil

qc[12, 1] <- qc[10,1] / qc[11,1] # %N calibration, Rosemount soil
qc[12, 2] <- NA                  # d15N calibration, Rosemount soil
qc[12, 3] <- qc[10,3] / qc[11,3] # %C calibration, Rosemount soil
qc[12, 4] <- NA                  # d13C calibration, Rosemount soil


## [ACETANILIDE] Calculate the r-squared, standard deviation, slope and mean standard values 
## for %N, d15N, %C and d13C. Add the expected %N, d15N, %C and d13C values to the matrix.
## Calculate the calibration value and add it to the matrix.

qc[13, 1] <- summary(lm(`%N` ~ `N analysis number`, data = acet))$r.squared
qc[13, 2] <- summary(lm(d15N ~ `N analysis number`, data = acet))$r.squared
qc[13, 3] <- summary(lm(`%C` ~ `C analysis number`, data = acet))$r.squared
qc[13, 4] <- summary(lm(d13C ~ `C analysis number`, data = acet))$r.squared

qc[15, 1] <- sd(acet$`%N`)
qc[15, 2] <- sd(acet$d15N)
qc[15, 3] <- sd(acet$`%C`)
qc[15, 4] <- sd(acet$d13C)

qc[14, 1] <- summary(lm(`%N` ~ `N analysis number`, data = acet))$coefficients[2, 1]
qc[14, 2] <- summary(lm(d15N ~ `N analysis number`, data = acet))$coefficients[2, 1]
qc[14, 3] <- summary(lm(`%C` ~ `C analysis number`, data = acet))$coefficients[2, 1]
qc[14, 4] <- summary(lm(d13C ~ `C analysis number`, data = acet))$coefficients[2, 1]

qc[16, 1] <- mean(acet$`%N`)
qc[16, 2] <- mean(acet$d15N)
qc[16, 3] <- mean(acet$`%C`)
qc[16, 4] <- mean(acet$d13C)

qc[17,1] <- 10.36 # Expected %N, acetanilide
qc[17,2] <- NA    # Expected d15N, acetanilide
qc[17,3] <- 71.09 # Expected %N, acetanilide
qc[17,4] <- NA    # Expected d13C, acetanilide

qc[18, 1] <- qc[16,1] / qc[17,1] # %N calibration, acetanilide
qc[18, 2] <- NA                  # d15N calibration, acetanilide
qc[18, 3] <- qc[16,3] / qc[17,3] # %C calibration, acetanilide
qc[18, 4] <- NA                  # d13C calibration, acetanilide


## [USGS 40] Calculate the r-squared, standard deviation, slope and mean standard values 
## for %N, d15N, %C and d13C. Add the expected %N, d15N, %C and d13C values to the matrix.
## Calculate the calibration value and add it to the matrix.
qc[19, 1] <- summary(lm(`%N` ~ `N analysis number`, data = usgs))$r.squared
qc[19, 2] <- summary(lm(d15N ~ `N analysis number`, data = usgs))$r.squared
qc[19, 3] <- summary(lm(`%C` ~ `C analysis number`, data = usgs))$r.squared
qc[19, 4] <- summary(lm(d13C ~ `C analysis number`, data = usgs))$r.squared

qc[20, 1] <- summary(lm(`%N` ~ `N analysis number`, data = usgs))$coefficients[2, 1]
qc[20, 2] <- summary(lm(d15N ~ `N analysis number`, data = usgs))$coefficients[2, 1]
qc[20, 3] <- summary(lm(`%C` ~ `C analysis number`, data = usgs))$coefficients[2, 1]
qc[20, 4] <- summary(lm(d13C ~ `C analysis number`, data = usgs))$coefficients[2, 1]

qc[21, 1] <- sd(usgs$`%N`)
qc[21, 2] <- sd(usgs$d15N)
qc[21, 3] <- sd(usgs$`%C`)
qc[21, 4] <- sd(usgs$d13C)

qc[22, 1] <- mean(usgs$`%N`)
qc[22, 2] <- mean(usgs$d15N)
qc[22, 3] <- mean(usgs$`%C`)
qc[22, 4] <- mean(usgs$d13C)

qc[23,1] <- 9.52   #Expected %N, USGS 40
qc[23,2] <- -4.52  #Expected d15N, USGS 40
qc[23,3] <- 40.8   #Expected %C, USGS 40
qc[23,4] <- -26.39 #Expected d13C, USGS 40

qc[24, 1] <- NA                      # %N calibration, USGS 40
qc[24, 2] <- qc[22, 2] - qc[23,2]    # d15N calibration, USGS 40 
qc[24, 3] <- NA                      # %C calibration, USGS 40
qc[24, 4] <- qc[22, 4] - qc[23,4]    # d13C calibration, USGS 40


### ADD QC VALUES TO STANDARDS DATAFRAME ###

# Convert the "Std" dataframes to matrices for each standard, and create a new matrix 
# for each standard containing the QC values.

peach.matrix <- as.matrix(peach)
rsmt.matrix <- as.matrix(rsmt)
acet.matrix <- as.matrix(acet)
usgs.matrix <- as.matrix(usgs)

peach.qc <- qc[1:6, 1:4]
rsmt.qc <- qc[7:12, 1:4]
acet.qc <- qc[13:18, 1:4]
usgs.qc <- qc[19:24, 1:4]

# Combine the two matrices into one matrix (which contains the standard values and QC values for each standard)

peach.qc <- as.data.frame(cbind(peach.qc, matrix(data=NA, ncol=9, nrow = 6)))
peach.qc <- as.matrix(peach.qc[,c("V5", "V6", "V7", "V8", "V9", "V10", "V11", "%N","d15N", "V12", "%C", "d13C")])
peach.qc.std <- as.data.frame(rbind(peach.matrix, peach.qc))

rsmt.qc <- as.data.frame(cbind(rsmt.qc, matrix(data=NA, ncol=9, nrow = 6)))
rsmt.qc <- as.matrix(rsmt.qc[,c("V5", "V6", "V7", "V8", "V9", "V10", "V11", "%N","d15N", "V12", "%C", "d13C")])
rsmt.qc.std <- as.data.frame(rbind(rsmt.matrix, rsmt.qc))

acet.qc <- as.data.frame(cbind(acet.qc, matrix(data=NA, ncol=9, nrow = 6)))
acet.qc <- as.matrix(acet.qc[,c("V5", "V6", "V7", "V8", "V9", "V10", "V11", "%N","d15N", "V12", "%C", "d13C")])
acet.qc.std <- as.data.frame(rbind(acet.matrix, acet.qc))

usgs.qc <- as.data.frame(cbind(usgs.qc, matrix(data=NA, ncol=9, nrow = 6)))
usgs.qc <- as.matrix(usgs.qc[,c("V5", "V6", "V7", "V8", "V9", "V10", "V11", "%N","d15N", "V12", "%C", "d13C")])
usgs.qc.std <- as.data.frame(rbind(usgs.matrix, usgs.qc))
    
# Merge all of the matrices containing both standard and QC values into one dataframe.
stds.qc <- rbind(peach.qc.std, rsmt.qc.std, acet.qc.std, usgs.qc.std)


# Apply corrections to data -----------------------------------------------

## Below, comment out (add a "#" symbol in front of the line) the standards that 
## you would NOT like to use for calibration (the standard you wish to use for 
## calibration should appear in black text). 

# %N calibration
#N.cal <- c(qc[6,1]) # [PEACH LEAVES]
#N.cal <- c(qc[12,1]) # [ROSEMOUNT SOIL]
N.cal <- c(qc[18,1]) # [ACETANILIDE]
#N.cal <- c(qc[24,1]) # [USGS 40]
print(N.cal) # Check that the calibration value looks right

# d15N calibration
#d15N.cal <- c(qc[6,2]) # [PEACH LEAVES]
#d15N.cal <- c(qc[12,2]) # [ROSEMOUNT SOIL]
#d15N.cal <- c(qc[18,2]) # [ACETANILIDE]
d15N.cal <- c(qc[24,2]) # [USGS 40]
print(d15N.cal) # Check that the calibration value looks right

# %C calibration
#C.cal <- c(qc[6,3]) # [PEACH LEAVES]
#C.cal <- c(qc[12,3]) # [ROSEMOUNT SOIL]
C.cal <- c(qc[18,3]) # [ACETANILIDE]
#C.cal <- c(qc[24,3]) # [USGS 40]
print(C.cal) # Check that the calibration value looks right

# %d13C calibration
#d13C.cal <- c(qc[6,4]) # [PEACH LEAVES]
#d13C.cal <- c(qc[12,4]) # [ROSEMOUNT SOIL]
#d13C.cal <- c(qc[18,4]) # [ACETANILIDE]
d13C.cal <- c(qc[24,4]) # [USGS 40]
print(d13C.cal) # Check that the calibration value looks right


## Apply calibration standards to the data and add the corrected data values to 
## the raw5 dataframe.
raw4 <- raw3 %>%
  mutate(corrected.N = `%N` / N.cal,
         corrected.d15N = d15N - d15N.cal,
         corrected.C = `%C` / C.cal,
         corrected.d13C = d13C - d13C.cal,
         corrected.CNratio = corrected.C / corrected.N) %>%
  dplyr::rename("N corrected" = corrected.N, "d15N corrected" = corrected.d15N, 
                "C corrected" = corrected.C, "d13C corrected" = corrected.d13C, 
                "Corrected C/N ratio" = corrected.CNratio) %>%
  .[order(.$'EA analysis order'),] # order the sheet by EA analysis order

# EXTRACT THE REPLICATES --------------------------------------------------

# Pull out the replicates. Add original sample and replicate data to same sheet. Order them based on "Sample ID".
reps_rep <- raw4 %>%
  filter(description == "unknown_replicate") # isolate replicate samples in new dataframe
rep_ids <- unlist(strsplit(reps_rep$`Sample ID`, "_rep")) # create vector containing Sample IDs of the replicate samples

reps_original <- raw4[raw4$`Sample ID` %in% rep_ids, ] # dataframe containing the 'original' samples that have replicates associated with them

reps <- rbind(reps_rep, reps_original) %>% # bind the reps dataframes into one df containing the original and replicate samples
  .[order(.$'Sample ID'),] # reorder dataframe by sample ID
rm(reps_rep, reps_original)


#Calculate the standard deviation between the original sample and replicate sample.
reps2 <- cbind(as.matrix(reps), matrix(data=NA, ncol=5, nrow = nrow(reps)))

#Rep 1
reps2[1,19] <- sd(c((reps2[1,14]),(reps2[2,14])))
reps2[1,20] <- sd(c((reps2[1,15]),(reps2[2,15])))
reps2[1,21] <- sd(c((reps2[1,16]),(reps2[2,16])))
reps2[1,22] <- sd(c((reps2[1,17]),(reps2[2,17])))

#Rep 2
reps2[3,19] <- sd(c((reps2[3,14]),(reps2[4,14])))
reps2[3,20] <- sd(c((reps2[3,15]),(reps2[4,15])))
reps2[3,21] <- sd(c((reps2[3,16]),(reps2[4,16])))
reps2[3,22] <- sd(c((reps2[3,17]),(reps2[4,17])))

#Rep 3
reps2[5,19] <- sd(c((reps2[5,14]),(reps2[6,14])))
reps2[5,20] <- sd(c((reps2[5,15]),(reps2[6,15])))
reps2[5,21] <- sd(c((reps2[5,16]),(reps2[6,16])))
reps2[5,22] <- sd(c((reps2[5,17]),(reps2[6,17])))

#Rep 4
reps2[7,19] <- sd(c((reps2[7,14]),(reps2[8,14])))
reps2[7,20] <- sd(c((reps2[7,15]),(reps2[8,15])))
reps2[7,21] <- sd(c((reps2[7,16]),(reps2[8,16])))
reps2[7,22] <- sd(c((reps2[7,17]),(reps2[8,17])))

#Rep 5
reps2[9,19] <- sd(c((reps2[9,14]),(reps2[10,14])))
reps2[9,20] <- sd(c((reps2[9,15]),(reps2[10,15])))
reps2[9,21] <- sd(c((reps2[9,16]),(reps2[10,16])))
reps2[9,22] <- sd(c((reps2[9,17]),(reps2[10,17])))

#Rep 6
reps2[11,19] <- sd(c((reps2[11,14]),(reps2[12,14])))
reps2[11,20] <- sd(c((reps2[11,15]),(reps2[12,15])))
reps2[11,21] <- sd(c((reps2[11,16]),(reps2[12,16])))
reps2[11,22] <- sd(c((reps2[11,17]),(reps2[12,17])))

#Rep 7
reps2[13,19] <- sd(c((reps2[13,14]),(reps2[14,14])))
reps2[13,20] <- sd(c((reps2[13,15]),(reps2[14,15])))
reps2[13,21] <- sd(c((reps2[13,16]),(reps2[14,16])))
reps2[13,22] <- sd(c((reps2[13,17]),(reps2[14,17])))

#Rep 8
reps2[15,19] <- sd(c((reps2[15,14]),(reps2[16,14])))
reps2[15,20] <- sd(c((reps2[15,15]),(reps2[16,15])))
reps2[15,21] <- sd(c((reps2[15,16]),(reps2[16,16])))
reps2[15,22] <- sd(c((reps2[15,17]),(reps2[16,17])))

#Rename new columns containing standard deviation calculations.
reps2 <- as.data.frame(reps2) %>%
  dplyr::rename("Std Dev N" = 19, "Std Dev C" = 20, "Std Dev d15N" =21,"Std Dev d13C" =22)



# Apply drift corrections -------------------------------------------------
dc.raw <- raw4

# Create dataframes to store drift correction values using acetenalide for %N and %C and USGS40 for d15N adn d13C
acet2 <- acet %>% 
  mutate("Drift corrected %N" = NA,  
         "Drift corrected d15N" = NA,
         "Drift corrected %C" = NA,
         "Drift corrected d13C" = NA)
usgs2 <- usgs %>% 
  mutate("Drift corrected %N" = NA,  
         "Drift corrected d15N" = NA,
         "Drift corrected %C" = NA,
         "Drift corrected d13C" = NA)

#%N DRIFT CORRECTION - Using acetenalide
if (acet.qc[1,8] > 0.7) {
  acet2$`Drift corrected %N`   <- acet2$`%N` - (acet2$`EA analysis order` * acet.qc[2,8])
  dc.raw$`Drift corrected %N` <- dc.raw$`%N` - (dc.raw$`EA analysis order` * acet.qc[2,8])
  acet.cal.N <- (mean(acet2$`Drift corrected %N`) / acet.qc[5,8])
  dc.raw$`Final drift corrected %N` <- dc.raw$`Drift corrected %N` / acet.cal.N } else {
    dc.raw$`Drift corrected %N` <- "r-squared < 0.7"
    dc.raw$`Final drift corrected %N`  <- "r-squared < 0.7"}

#d15N DRIFT CORRECTION - Using USGS 40
if (usgs.qc[1,9] > 0.7) {
  usgs2$`Drift corrected d15N` <- usgs2$d15N - (usgs2$`EA analysis order` * usgs.qc[2,9])
  dc.raw$`Drift corrected d15N` <- dc.raw$d15N - (dc.raw$`EA analysis order` * usgs.qc[2,9])
  usgs.cal.d15N <- (mean(usgs2$`Drift corrected d15N`) - usgs.qc[5,9])
  dc.raw$`Final drift corrected d15N` <- dc.raw$`Drift corrected d15N` - usgs.cal.d15N} else {
    dc.raw$`Drift corrected d15N` <- "r-squared < 0.7"
    dc.raw$`Final drift corrected d15N`  <- "r-squared < 0.7"}

#%C DRIFT CORRECTIONS - using acetenalide
if (acet.qc[1,11] > 0.7) {
  acet2$`Drift corrected %C`   <- acet2$`%C` - (acet2$`EA analysis order` * acet.qc[2,11])
  dc.raw$`Drift corrected %C` <- dc.raw$`%C` - (dc.raw$`EA analysis order` * acet.qc[2,11])
  acet.cal.C <- (mean(acet2$`Drift corrected %C`) / acet.qc[5,11])
  dc.raw$`Final drift corrected %C` <- dc.raw$`Drift corrected %C` / acet.cal.C } else {
    dc.raw$`Drift corrected %C` <- "r-squared < 0.7"
    dc.raw$`Final drift corrected %C`  <- "r-squared < 0.7"}
        
#d13C DRIFT CORRECTION - Using USGS 40
if (usgs.qc[1,12] > 0.7) {
  usgs2$`Drift corrected d13C` <- usgs2$d13C - (usgs2$`EA analysis order` * usgs.qc[2,12])
  dc.raw$`Drift corrected d13C` <- dc.raw$d13C - (dc.raw$`EA analysis order` * usgs.qc[2,12])
  usgs.cal.d13C <- (mean(usgs2$`Drift corrected d13C`) - usgs.qc[5,12])
  dc.raw$`Final drift corrected d13C`  <- dc.raw$`Drift corrected d13C` - usgs.cal.d13C} else {
    dc.raw$`Drift corrected d13C` <- "r-squared < 0.7"
    dc.raw$`Final drift corrected d13C`  <- "r-squared < 0.7"}

  
# Remove all columns excpet those containing "corrected" data
dc.raw2 <- dc.raw[,c(1:3,18:ncol(dc.raw))]
#dc.raw2 <- dc.raw2[,c(1,2,3,
                 #   4,9,10,
                 #   5,11,12,6,
                 #   7,13,14,8)]

dc.raw$runID <- project_name

# Add columns to peach and Rosemount standard dataframes (even if these are not used for drift corrections,
# all dataframes need to have the same number of columns to be bound into single frame).
peach2 <- peach %>% 
  mutate("Drift corrected %N" = NA,  
         "Drift corrected d15N" = NA,
         "Drift corrected %C" = NA,
         "Drift corrected d13C" = NA)

rsmt2 <- rsmt %>% 
  mutate("Drift corrected %N" = NA,  
         "Drift corrected d15N" = NA,
         "Drift corrected %C" = NA,
         "Drift corrected d13C" = NA)


# Re-bind the standard dataframe together so they all appear on one sheet.
dc.stds <- rbind(peach2,
                 rsmt2,
                 acet2,
                 usgs2)


#### SAVE THE DATA #### (M. Gieske)

sheet_list <- list("weight sheet" = weigh_sheet,
                   "EA data" = ea_data,
                   "Processed data" = raw4,
                   "Standards" = stds,
                   "QC" = qc,
                   "Standards and QC" = stds.qc,
                   "Replicates" = reps2,
                   "Corrected data" = raw4,
                   "Drift corrected standards" = dc.stds,
                   "Drift corrected data" = dc.raw)

setwd(here::here("data", "processed"))
openxlsx::write.xlsx(sheet_list, file = paste(project_name, "_fromR_processed.xlsx", sep = ""), rowNames = TRUE)




