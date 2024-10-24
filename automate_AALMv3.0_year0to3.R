# DIRECTONS

#IN AALM: 
# 1. delete (comment out) pop-up dialog message boxes in VBA. (see photo) cannot automate with those
# 2. change "no" to "yes" for all dropdowns on Simulation Control page EXCEPT "other"
# also choose Male or Female (affects HCTA value on Time Ind Parameter Tab)
# 3. make sure you close app before running code from start

# IN R: 
# 1. create COM objects excel, workbooks, and sheets you need to access 
# 2. initialize input and output dataframe 
# 3. Delete folders created by previous runs
# START LOOP:
# 4. index desired cells in desired location for each parameter (parameter and media sheets)
# 5. replace them one-by-one with a random selection from the normal distribution for that cell
# 6. once all cells have been replaced, run AALM_RunSimulation
# 7. store input data (all parameter values) for that iteration
# 8. locate output folder, read csv file for detailed ouputs
# 9. store output data (time in yrs, blood pb, bone pb) for that iteration
#    continue loop for desired amount
# END LOOP 
#10. close workbook and excel app, release all COM objects
#11. combine and plot data
#12. sensitivity analysis

#clear workspace and close excel if needed
rm(list = ls())
rm(list = grep("COM", ls(), value = TRUE)) 
#workbook$Close(FALSE)
#excel$Quit()

library(pacman)
p_load(RDCOMClient, readxl, openxlsx, tibble, readr, tidyverse)

#### RUN THIS FOR AGE 0-3

#BEFORE MODEL: read text file

text_path <- "<file_path_to_input_file>" #Define the filepath for input parameters, 0-2 year

par_df <- read_table(text_path, col_names = FALSE)
print(head(par_df))
#240 rows, 24 variables (columns) for year 0-3

# Define the file path for model
file_path <- "<file_path_to_model>" # Define the file path for model (ends in /AALMv3-0_030124.xlsm)

# Create a COM object for Excel
excel <- COMCreate("Excel.Application")

# Access the Workbook
workbooks <- excel[["Workbooks"]]

# Open the workbook
workbook <- workbooks$Open(file_path)


excel[["Visible"]] <- TRUE

output_df <- data.frame(Iteration = numeric(),
                        Time_years = numeric(),
                        Blood_Pb = numeric(),
                        Bone_Pb = numeric())

input_df <- data.frame(Iteration = numeric(), 
                       bldmot = numeric(),
                       rdiff = numeric(),
                       rcs2df = numeric(),
                       rcort_e = numeric(),
                       rcort_f = numeric(),
                       rcort_g = numeric(),
                       rcort_h = numeric(),
                       rrbc_e = numeric(),
                       rrbc_f = numeric(),
                       rrbc_g = numeric(),
                       rrbc_h = numeric(),  
                       soil_f = numeric(),
                       soil_g = numeric(),
                       soil_h = numeric(),
                       soil_i = numeric(),
                       soil_j = numeric(),
                       soil_k = numeric(),
                       soil_l = numeric(),
                       dust_f = numeric(),
                       dust_g = numeric(),
                       dust_h = numeric(),
                       dust_i = numeric(),
                       dust_j = numeric(),
                       dust_k = numeric(),
                       dust_l = numeric(),
                       water_f = numeric(),
                       water_g = numeric(),
                       water_h = numeric(),
                       water_i = numeric(),
                       water_j = numeric(),
                       water_k = numeric(),
                       water_l = numeric(),
                       air_f = numeric(),
                       air_g = numeric(),
                       air_h = numeric(),
                       air_i = numeric(),
                       air_j = numeric(),
                       air_k = numeric(),
                       air_l = numeric(),
                       food_f = numeric(),
                       food_g = numeric(),
                       food_h = numeric(),
                       food_i = numeric(),
                       food_j = numeric(),
                       food_k = numeric(),
                       food_l = numeric())

##### Check if folders already exist and delete 

directory <- "<directory_path_of_AALM>" # Define directory path (folder path of the model; runs automatically save to this folder)

# List all directories in the specified directory
subdirectories <- list.dirs(directory, full.names = TRUE)

# Filter directories that start with "Iteration"
iteration_folders <- subdirectories[grep("^Iteration", basename(subdirectories))]

# Delete each folder
for (folder in iteration_folders) {
  unlink(folder, recursive = TRUE)
}

#### START LOOP ################################## 

#COLUMN ORDER FOR TEXT FILE 0-3 YEAR:

# 1 BLDMOT
# 2 RDIFF
# 3 RCS2DF

# 4 RCORT0
# 5 RCORT03
# 6 RCORT1

# 7 RRBC0
# 8 RRBC03
# 9 RRBC1 

# 10 SOIL0
# 11 SOIL1
# 12 SOIL2

# 13 DUST0
# 14 DUST1
# 15 DUST2

# 16 WATER0
# 17 WATER1
# 18 WATER2

# 19 AIR0
# 20 AIR1
# 21 AIR2

# 22 FOOD0
# 23 FOOD1
# 24 FOOD2

## Simulation Control
sim_sheet <- workbook$Worksheets("Simulation Control")

#Enter Gender and Age
age_range <- sim_sheet$Range("E13")
gender_range <- sim_sheet$Range("E14")

age_range[["Value"]] <- paste("3") #change
gender_range[["Value"]] <- paste("Male")

for (x in 1:240) { #for each row of the df, run the model
  
  cat("Row", x, ":\n")
  params <- par_df[x, ] #index current row
  
  # Enter simulation name
  sim_range <- sim_sheet$Range("E6")
  
  if (is.null(sim_range)) {
    stop("Failed to reference the range E6 in Simulation Control sheet")
  }
  
  sim_range[["Value"]] <- paste("Iteration", x, sep = "")
  
  #### Time Ind Phys Params - BLDMOT #########################
  time_ind_sheet <- workbook$Worksheets("Time Ind Phys Params")
  # 1. BLDMOT
  # 2. RDIFF
  # 3. RCS2DF
  
  #BLDMOT (E8: 0.62)
  
  #BLDMOT value from first column
  bldmot <- params[1] #only one time
  
  time_ind_range <- time_ind_sheet$Range("E8")
  time_ind_range[["Value"]] <- paste(bldmot)
  
  #### Time Dep Phys Params  ##################################
  time_dep_sheet <- workbook$Worksheets("Time Dep Phys Params")
  
  #RDIFF (19)
  #E,F,G,H: 0.02311
  
  #RDIFF value from second column
  rdiff <- params[2] #all the same
  
  e <- time_dep_sheet$Range("E19")
  f <- time_dep_sheet$Range("F19")
  g <- time_dep_sheet$Range("G19")
  h <- time_dep_sheet$Range("H19")
  
  e[["Value"]] <- paste(rdiff)
  f[["Value"]] <- paste(rdiff)
  g[["Value"]] <- paste(rdiff)
  h[["Value"]] <- paste(rdiff)
  
  #RCS2DF (18)
  #E,F,G,H: 0.65
  
  #RCS2DF value from third column
  rcs2df <- params[3] #all the same
  
  e <- time_dep_sheet$Range("E18")
  f <- time_dep_sheet$Range("F18")
  g <- time_dep_sheet$Range("G18")
  h <- time_dep_sheet$Range("H18")
  
  e[["Value"]] <- paste(rcs2df)
  f[["Value"]] <- paste(rcs2df)
  g[["Value"]] <- paste(rcs2df)
  h[["Value"]] <- paste(rcs2df)
  
  
  #RCORT (16)
  #E: 0.02040 Age 0 ***
  #F: 0.01644 Age 0.27 ***
  #G: 0.00576 Age 1 ***
  #H: 0.00308 Age 5
  
  # 4. RCORT0 - 0 to 3 months (E)
  # 5. RCORT03 - 3 to 12 months (F)
  # 6. RCORT1 - 13 to 24 months (G)
  
  rcort_e <- params[4]
  rcort_f <- params[5]
  rcort_g <- params[6]
  
  #others stay the same
  rcort_h <- 0.00308
  
  #apply changes to cells
  e <- time_dep_sheet$Range("E16")
  f <- time_dep_sheet$Range("F16")
  g <- time_dep_sheet$Range("G16")
  h <- time_dep_sheet$Range("H16")
  
  e[["Value"]] <- paste(rcort_e)
  f[["Value"]] <- paste(rcort_f)
  g[["Value"]] <- paste(rcort_g)
  h[["Value"]] <- paste(rcort_h)
  
  #RRBC (22)
  #E: 0.46200 Age 0 ***
  #F: 0.46200 Age 0.27 ***
  #G: 0.78540 Age 1 ***
  #H: 0.49860 Age 5
  
  # 7. RRBC0 - 0 to 3 months (E)
  # 8. RRBC03 - 3 to 12 months (F)
  # 9. RRBC1 - 13 to 24 months (G)
  
  rrbc_e <- params[7]
  rrbc_f <- params[8]
  rrbc_g <- params[9]
  
  #others stay the same
  rrbc_h <- 0.4986
  
  #apply changes to cells
  e <- time_dep_sheet$Range("E22")
  f <- time_dep_sheet$Range("F22")
  g <- time_dep_sheet$Range("G22")
  h <- time_dep_sheet$Range("H22")
  
  e[["Value"]] <- paste(rrbc_e)
  f[["Value"]] <- paste(rrbc_f)
  g[["Value"]] <- paste(rrbc_g)
  h[["Value"]] <- paste(rrbc_h)
  
  #### Media - Intake rates ######################################
  media_sheet <- workbook$Worksheets("Media")
  
  #SOIL (28)
  
  #FOR High-CONC scenarios: change concentration
  
  sc <- media_sheet$Range("F12")
  
  #soil_conc <- 400 #high 
  soil_conc <- 200 #default
  sc[["Value"]] <- paste(soil_conc)
  
  #F: 0.03870 Age 0 ***
  #G: 0.04230 Age 1 ***
  #H: 0.03015 Age 2 ***
  #I: 0.02835 Age 3
  #J: 0.03015 Age 4
  #K: 0.02340 Age 5
  #L: 0.02475 Age 6
  
  # 10 SOIL0 (F)
  # 11 SOIL1 (G)
  # 12 SOIL2 (H)
  
  soil_f <- params[10]
  soil_g <- params[11]
  soil_h <- params[12]
  
  #others stay the same
  soil_i <- 0.02835
  soil_j <- 0.03015
  soil_k <- 0.02340
  soil_l <- 0.02475
  
  #apply changes to cells
  f <- media_sheet$Range("F28")
  g <- media_sheet$Range("G28")
  h <- media_sheet$Range("H28")
  i <- media_sheet$Range("I28")
  j <- media_sheet$Range("J28")
  k <- media_sheet$Range("K28")
  l <- media_sheet$Range("L28")
  
  f[["Value"]] <- paste(soil_f)
  g[["Value"]] <- paste(soil_g)
  h[["Value"]] <- paste(soil_h)
  i[["Value"]] <- paste(soil_i)
  j[["Value"]] <- paste(soil_j)
  k[["Value"]] <- paste(soil_k)
  l[["Value"]] <- paste(soil_l)
  
  #DUST (55)
  
  #FOR High-CONC scenarios: change concentration
  
  dc <- media_sheet$Range("F39")
  
  #dust_conc <- 1000 #high 
  dust_conc <- 150 #default
  dc[["Value"]] <- paste(dust_conc)
  
  #F: 0.04730 Age 0 ***
  #G: 0.05170 Age 1 ***
  #H: 0.03685 Age 2 ***
  #I: 0.03465 Age 3
  #J: 0.03685 Age 4
  #K: 0.02860 Age 5
  #L: 0.03025 Age 6
  
  # 13 DUST0 (F)
  # 14 DUST1 (G)
  # 15 DUST2 (H)
  
  dust_f <- params[13]
  dust_g <- params[14]
  dust_h <- params[15]
  
  #others stay the same
  dust_i <- 0.03465
  dust_j <- 0.03465
  dust_k <- 0.02860
  dust_l <- 0.03025
  
  #apply changes to cells
  f <- media_sheet$Range("F55")
  g <- media_sheet$Range("G55")
  h <- media_sheet$Range("H55")
  i <- media_sheet$Range("I55")
  j <- media_sheet$Range("J55")
  k <- media_sheet$Range("K55")
  l <- media_sheet$Range("L55")
  
  f[["Value"]] <- paste(dust_f)
  g[["Value"]] <- paste(dust_g)
  h[["Value"]] <- paste(dust_h)
  i[["Value"]] <- paste(dust_i)
  j[["Value"]] <- paste(dust_j)
  k[["Value"]] <- paste(dust_k)
  l[["Value"]] <- paste(dust_l)
  
  #WATER (82)
  
  #FOR High-CONC scenarios: change concentration
  
  wc <- media_sheet$Range("F66")
  
  #water_conc <- 15 #high 
  water_conc <- 0.9 #default
  wc[["Value"]] <- paste(water_conc)
  
  #F: 0.40 Age 0 ***
  #G: 0.43 Age 1 ***
  #H: 0.51 Age 2 ***
  #I: 0.54 Age 3
  #J: 0.57 Age 4
  #K: 0.60 Age 5
  #L: 0.63 Age 6
  
  # 16 WATER0 (F)
  # 17 WATER1 (G)
  # 18 WATER2 (H)
  
  water_f <- params[16]
  water_g <- params[17]
  water_h <- params[18]
  
  #others stay the same
  water_i <- 0.54
  water_j <- 0.57
  water_k <- 0.60
  water_l <- 0.63
  
  #apply changes to cells
  f <- media_sheet$Range("F82")
  g <- media_sheet$Range("G82")
  h <- media_sheet$Range("H82")
  i <- media_sheet$Range("I82")
  j <- media_sheet$Range("J82")
  k <- media_sheet$Range("K82")
  l <- media_sheet$Range("L82")
  
  f[["Value"]] <- paste(water_f)
  g[["Value"]] <- paste(water_g)
  h[["Value"]] <- paste(water_h)
  i[["Value"]] <- paste(water_i)
  j[["Value"]] <- paste(water_j)
  k[["Value"]] <- paste(water_k)
  l[["Value"]] <- paste(water_l)
  
  #AIR (109)
  
  #FOR High-CONC scenarios: change concentration
  
  ac <- media_sheet$Range("F93")
  
  #air_conc <- 30 #high
  air_conc <- 0.1 #default
  ac[["Value"]] <- paste(air_conc)
  
  # INTAKE RATE PARAMETERS
  #F: 3.22 Age 0 ***
  #G: 4.97 Age 1 ***
  #H: 6.09 Age 2 ***
  #I: 6.95 Age 3
  #J: 7.68 Age 4
  #K: 8.32 Age 5
  #L: 8.89 Age 6
  
  # 19 AIR0 (F)
  # 20 AIR1 (G)
  # 21 AIR2 (H)
  
  air_f <- params[19]
  air_g <- params[20]
  air_h <- params[21]
  
  #others stay the same
  air_i <- 6.95
  air_j <- 7.68
  air_k <- 8.32
  air_l <- 8.89
  
  #apply changes to cells
  f <- media_sheet$Range("F109")
  g <- media_sheet$Range("G109")
  h <- media_sheet$Range("H109")
  i <- media_sheet$Range("I109")
  j <- media_sheet$Range("J109")
  k <- media_sheet$Range("K109")
  l <- media_sheet$Range("L109")
  
  f[["Value"]] <- paste(air_f)
  g[["Value"]] <- paste(air_g)
  h[["Value"]] <- paste(air_h)
  i[["Value"]] <- paste(air_i)
  j[["Value"]] <- paste(air_j)
  k[["Value"]] <- paste(air_k)
  l[["Value"]] <- paste(air_l)
  
  #FOOD (120)
  #F: 2.66 Age 0 ***
  #G: 5.03 Age 1 ***
  #H: 5.21 Age 2 ***
  #I: 5.38 Age 3
  #J: 5.64 Age 4
  #K: 6.04 Age 5
  #L: 5.95 Age 6
  
  # 22 FOOD0 (F)
  # 23 FOOD1 (G)
  # 24 FOOD2 (H)
  
  food_f <- params[22]
  food_g <- params[23]
  food_h <- params[24]
  
  #others stay the same
  food_i <- 5.38
  food_j <- 5.64
  food_k <- 6.04
  food_l <- 5.95
  
  #apply changes to cells
  f <- media_sheet$Range("F120")
  g <- media_sheet$Range("G120")
  h <- media_sheet$Range("H120")
  i <- media_sheet$Range("I120")
  j <- media_sheet$Range("J120")
  k <- media_sheet$Range("K120")
  l <- media_sheet$Range("L120")
  
  f[["Value"]] <- paste(food_f)
  g[["Value"]] <- paste(food_g)
  h[["Value"]] <- paste(food_h)
  i[["Value"]] <- paste(food_i)
  j[["Value"]] <- paste(food_j)
  k[["Value"]] <- paste(food_k)
  l[["Value"]] <- paste(food_l)
  
  ###############################################
  
  # Create the current input data frame
  input_current <- data.frame(Iteration = x, 
                              bldmot = bldmot,
                              rdiff = rdiff,
                              rcs2df = rcs2df,
                              rcort_e = rcort_e,
                              rcort_f = rcort_f,
                              rcort_g = rcort_g,
                              rcort_h = rcort_h,
                              rrbc_e = rrbc_e,
                              rrbc_f = rrbc_f,
                              rrbc_g = rrbc_g,
                              rrbc_h = rrbc_h,  
                              soil_f = soil_f,
                              soil_g = soil_g,
                              soil_h = soil_h,
                              soil_i = soil_i,
                              soil_j = soil_j,
                              soil_k = soil_k,
                              soil_l = soil_l,
                              dust_f = dust_f,
                              dust_g = dust_g,
                              dust_h = dust_h,
                              dust_i = dust_i,
                              dust_j = dust_j,
                              dust_k = dust_k,
                              dust_l = dust_l,
                              water_f = water_f,
                              water_g = water_g,
                              water_h = water_h,
                              water_i = water_i,
                              water_j = water_j,
                              water_k = water_k,
                              water_l = water_l,
                              air_f = air_f,
                              air_g = air_g,
                              air_h = air_h,
                              air_i = air_i,
                              air_j = air_j,
                              air_k = air_k,
                              air_l = air_l,
                              food_f = food_f,
                              food_g = food_g,
                              food_h = food_h,
                              food_i = food_i,
                              food_j = food_j,
                              food_k = food_k,
                              food_l = food_l)
  
  # Add the current input to the input dataframe                           
  input_df <- rbind(input_df, input_current)
  
  ######### DONE EDITING; RUN ###########################
  excel$Run("AALM_RunSimulation")
  
  #optional - wait for output file to be generated
  Sys.sleep(5) #adjust if necessary
  
  ######## COLLECT OUTPUT ###############################
  folder_path <- paste0("Iteration", x)
  
  #                     Path to folder with AALM
  full_path <- (paste0("<directory_path_of_AALM>", folder_path, "/Out_Iteration", x, ".csv"))   #Build and index the directory path of AALM, where the output folder is created
  
  #Read output data
  if (file.exists(full_path)) {
    out_data <- read_csv(full_path, show_col_types = FALSE)
    
    output_current <- out_data[, c("Years", "Cblood", "Cbone")]
    colnames(output_current) <- c("Time_years", "Blood_Pb", "Bone_Pb")
    output_current$Iteration <- x
    
    output_df <- rbind(output_df, output_current)
    
  } else {
    warning(paste("Output file for iteration", x, "not found."))
  }
  
} # END LOOP

# Close the workbook and quit the app:
workbook$Close(FALSE)
excel$Quit()

rm(list = grep("COM", ls(), value = TRUE)) #release all COM objects

#display
print(input_df)
print(output_df)

#COMBINE DATA 
# Merge the data frames by "Iteration"
combined_df <- merge(output_df, input_df, by = "Iteration", all.x = TRUE)

# Print the combined data frame to verify
print(combined_df)

library(tidyverse)
#desired output data
output_year0to3 <- combined_df %>% # year 0 to 3
  filter(Time_years == 3.000) %>%
  select(Iteration, Time_years, Blood_Pb, Bone_Pb, X1, X2, X3, X4, X5, X6, X7, X8, X9, X10, X11, X12, X13, X14, X15, X16, X17, X18, X19, X20, X21, X22, X23, X24)

############ Write dataframe to a TXT file, CHANGE NAME HERE ################
write.table(output_year0to3, "<output_file_name>", row.names = FALSE, sep = "\t") #write output file name

