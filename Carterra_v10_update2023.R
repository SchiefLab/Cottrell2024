# Author: Oleksandr Kalyuzhniy
# Creation Date: 03/11/2022
# modified on 06/10/2022. Added checks for KD_fix with Koff > 0.5. Added "estimated" and "drop" columns

#install.packages("openxlsx",dependencies=TRUE)
#install.packages("tidyverse")
library("openxlsx")
library("tidyverse")

#choose working directory

setwd("/Users/oleks/Documents/LabBook/p1602-LSA2")

#constants with names of used files
batch_file <- "p1602-3-koff-fast.xlsx"
process_file <- "p1602-3-koff-slow.xlsx"
AnalytesMW_file <- "p1602-analytes.txt" #one file for the whole project

#constants to update every time
projectName = "p1602-3"             #change run count with each run
NumberOfAnalytes <- 3       #number of analytes in the run. Usually 1 or 2 or 3
NumberOfPrints <- 1         #number of ligand prints for each analyte. Usually 1 or 2 or 3

#constants to update only on special occasions
AbsRowMax <- 96 * NumberOfPrints * NumberOfAnalytes + 3   #this is the last row of Abs report points
SigRowMin <- AbsRowMax + 5                                #this is the first row of analyte report points
LigandValency = 2
rmax_mat_low = 0.1
rmax_mat_high = 10.0
signal_ratio = 0.1
koff_boundary = 0.009
LigandMW = 150000
outName = "_SPR_Kinetics_Summary"
Level_low = 25 #usually 75. For trimers 25
RmaxExp_low = 25 #usually 25. For trimers 25



Batch_data_in <- read.xlsx(xlsxFile = batch_file, sheet = 1, startRow = 2, colNames = FALSE, cols = 1:11)
names(Batch_data_in) = c(
  "ROI",
  "Bay",
  "Position",
  "Ligand",
  "Analyte",
  "Flagged Data",
  "batch_kon",
  "batch_koff",
  "batch_KD",
  "batch_Rmax",
  "batch_Res"
)

Batch_data_in <- Batch_data_in[-c(2, 3, 6)] 

Batch_data_in$ROI = as.numeric(Batch_data_in$ROI)

Batch_data_in$batch_kon = as.numeric(Batch_data_in$batch_kon)

Batch_data_in$batch_koff = as.numeric(Batch_data_in$batch_koff)

Batch_data_in$batch_KD = as.numeric(Batch_data_in$batch_KD)

Batch_data_in$batch_Rmax = as.numeric(Batch_data_in$batch_Rmax)

Batch_data_in$batch_Res = as.numeric(Batch_data_in$batch_Res)

#get signal data and format it

signal.one.batch <- read.xlsx(xlsxFile = batch_file, sheet = 12, startRow = SigRowMin, colNames = FALSE, cols = 1:11)

names(signal.one.batch) = c("ROI",
                            "Bay",
                            "Position",
                            "Ligand",
                            "Track",
                            "Analyte",
                            "Conc_M_batch",
                            "Link_Group",
                            "Signal_batch",
                            "SD_batch",
                            "Slope_batch")

signal.one.batch <- signal.one.batch[-c(2, 3, 5, 8, 10, 11)]

signal.one.batch$ROI = as.numeric(signal.one.batch$ROI)

signal.one.batch$Conc_M_batch = as.numeric(signal.one.batch$Conc_M_batch)

signal.one.batch$Signal_batch = as.numeric(signal.one.batch$Signal_batch)

#find max signal for each ROI and analyte in signal batch table

signal.two.batch = signal.one.batch[order(signal.one.batch$Signal_batch, decreasing = TRUE), ]

signal.batch = signal.two.batch[!duplicated(sprintf("%s-%s", signal.two.batch$ROI, signal.two.batch$Analyte)), ]

#combine Batch_data_in with signal.batch

Batch_data_signal = merge(Batch_data_in, signal.batch, by = c("ROI", "Ligand", "Analyte"))

#find negat_contr_max.batch for each analyte

negat_contr_batch = Batch_data_signal[str_detect(Batch_data_signal$Ligand, "buffer") | str_detect(Batch_data_signal$Ligand, "Ligand"),]

negat_contr_batch = negat_contr_batch[-c(1,2,4,5,6,7,8,9)]

negat_contr_batch = negat_contr_batch[order(negat_contr_batch$Signal_batch, decreasing = TRUE), ]

negat_contr_max.batch = negat_contr_batch[!duplicated(negat_contr_batch$Analyte), ]

names(negat_contr_max.batch) = c("Analyte", "Signal_batch_negat_control")

#add Signal_batch_negat_control and Signal_batch_less_NC columns to Batch_data_signal. This way we would know where signal is less then negative control (NC)

Batch_data_signal = merge(Batch_data_signal, negat_contr_max.batch, by = "Analyte")

Batch_data_signal$Signal_batch_less_NC = Batch_data_signal$Signal_batch - Batch_data_signal$Signal_batch_negat_control * 1.1

#set parameters to NA in batch table if signal is less then or equal to max of negative control for analyte1
if (length(Batch_data_signal[Batch_data_signal$Signal_batch_less_NC <= 0, ]$Signal_batch_less_NC) != 0) {
  Batch_data_signal[Batch_data_signal$Signal_batch_less_NC <= 0, ]$batch_kon = NA
  
}
if (length(Batch_data_signal[Batch_data_signal$Signal_batch_less_NC <= 0, ]$Signal_batch_less_NC) != 0) {
  Batch_data_signal[Batch_data_signal$Signal_batch_less_NC <= 0, ]$batch_koff = NA
  
}
if (length(Batch_data_signal[Batch_data_signal$Signal_batch_less_NC <= 0, ]$Signal_batch_less_NC) != 0) {
  Batch_data_signal[Batch_data_signal$Signal_batch_less_NC <= 0, ]$batch_KD = NA
  
}
if (length(Batch_data_signal[Batch_data_signal$Signal_batch_less_NC <= 0, ]$Signal_batch_less_NC) != 0) {
  Batch_data_signal[Batch_data_signal$Signal_batch_less_NC <= 0, ]$batch_Rmax = NA
  
}
if (length(Batch_data_signal[Batch_data_signal$Signal_batch_less_NC <= 0, ]$Signal_batch_less_NC) != 0) {
  Batch_data_signal[Batch_data_signal$Signal_batch_less_NC <= 0, ]$batch_Res = NA
  
}

#set parameters to "NA" in batch table if koff is less then or equal to koff_boundary
if (length(Batch_data_signal[!is.na(Batch_data_signal$batch_koff) &
                                      (Batch_data_signal$batch_koff <= koff_boundary), ]$batch_koff) != 0) {
  Batch_data_signal[!is.na(Batch_data_signal$batch_koff) &
                               (Batch_data_signal$batch_koff <= koff_boundary), ]$batch_kon = NA
  
}
if (length(Batch_data_signal[!is.na(Batch_data_signal$batch_koff) &
                                      (Batch_data_signal$batch_koff <= koff_boundary), ]$batch_koff) != 0) {
  Batch_data_signal[!is.na(Batch_data_signal$batch_koff) &
                               (Batch_data_signal$batch_koff <= koff_boundary), ]$batch_KD = NA
  
}
if (length(Batch_data_signal[!is.na(Batch_data_signal$batch_koff) &
                                      (Batch_data_signal$batch_koff <= koff_boundary), ]$batch_koff) != 0) {
  Batch_data_signal[!is.na(Batch_data_signal$batch_koff) &
                               (Batch_data_signal$batch_koff <= koff_boundary), ]$batch_Rmax = NA
  
}
if (length(Batch_data_signal[!is.na(Batch_data_signal$batch_koff) &
                                      (Batch_data_signal$batch_koff <= koff_boundary), ]$batch_koff) != 0) {
  Batch_data_signal[!is.na(Batch_data_signal$batch_koff) &
                               (Batch_data_signal$batch_koff <= koff_boundary), ]$batch_Res = NA
  
}
if (length(Batch_data_signal[!is.na(Batch_data_signal$batch_koff) &
                                      (Batch_data_signal$batch_koff <= koff_boundary), ]$batch_koff) != 0) {
  Batch_data_signal[!is.na(Batch_data_signal$batch_koff) &
                               (Batch_data_signal$batch_koff <= koff_boundary), ]$batch_koff = NA
  
}

#make table with top concentration used for batch data
signal.three.batch = signal.one.batch[order(signal.one.batch$Conc_M_batch, decreasing = TRUE), ]

signal.topconc.batch = signal.three.batch[!duplicated(sprintf(
  "%s-%s",
  signal.three.batch$ROI,
  signal.three.batch$Analyte
)), ]

names(signal.topconc.batch) = c(
  "ROI",
  "Ligand",
  "Analyte",
  "Top_Conc_M_batch",
  "Top_conc_Signal_batch"
)

#read signal_process_data file and find max signal for negative control.
signal.one.process      = read.xlsx(xlsxFile = process_file, sheet = 11, startRow = SigRowMin, colNames = FALSE, cols = 1:11)

names(signal.one.process) = c("ROI",
                              "Bay",
                              "Position",
                              "Ligand",
                              "Track",
                              "Analyte",
                              "Conc_M_process",
                              "Link_Group",
                              "Signal_process",
                              "SD_process",
                              "Slope_process")

signal.one.process <- signal.one.process[-c(2, 3, 5, 8, 10, 11)]

signal.one.process$ROI = as.numeric(signal.one.process$ROI)

signal.one.process$Conc_M_process = as.numeric(signal.one.process$Conc_M_process)

signal.one.process$Signal_process = as.numeric(signal.one.process$Signal_process)

#find max signal for each ROI in signal.one.process.analyte1

signal.two.process = signal.one.process[order(signal.one.process$Signal_process, decreasing = TRUE), ]

signal.process = signal.two.process[!duplicated(sprintf("%s-%s", signal.two.process$ROI, signal.two.process$Analyte)), ]

#get process table and add signals to it
process_data_in = read.xlsx(xlsxFile = process_file, sheet = 1, startRow = 2, colNames = FALSE, cols = 1:11)

names(process_data_in) = c(
  "ROI",
  "Bay",
  "Position",
  "Ligand",
  "Analyte",
  "Flagged Data",
  "process_kon",
  "process_koff",
  "process_KD",
  "process_Rmax",
  "process_Res"
)

process_data_in <- process_data_in[-c(2, 3, 6)] 

process_data_in$ROI = as.numeric(process_data_in$ROI)

process_data_in$process_kon = as.numeric(process_data_in$process_kon)

process_data_in$process_koff = as.numeric(process_data_in$process_koff)

process_data_in$process_KD = as.numeric(process_data_in$process_KD)

process_data_in$process_Rmax = as.numeric(process_data_in$process_Rmax)

process_data_in$process_Res = as.numeric(process_data_in$process_Res)

process_data_signal = merge(process_data_in, signal.process, by = c("ROI", "Ligand", "Analyte"))

negat_contr_process = process_data_signal[str_detect(process_data_signal$Ligand, "buffer") | str_detect(process_data_signal$Ligand, "Ligand"),]

negat_contr_process = negat_contr_process[-c(1,2,4,5,6,7,8,9)]

negat_contr_process = negat_contr_process[order(negat_contr_process$Signal_process, decreasing = TRUE), ]

negat_contr_max.process = negat_contr_process[!duplicated(negat_contr_process$Analyte), ]

names(negat_contr_max.process) = c("Analyte", "Signal_process_negat_control")

#add Signal_process_negat_control and Signal_process_less_NC columns to process_data_signal. This way we would know where signal is less then negative control (NC)

process_data_signal = merge(process_data_signal, negat_contr_max.process, by = "Analyte")

process_data_signal$Signal_process_less_NC = process_data_signal$Signal_process - process_data_signal$Signal_process_negat_control * 1.1

#set parameters to NA in process table if signal is less then or equal to max of negative control
if (length(process_data_signal[!is.na(process_data_signal$Signal_process_less_NC) &
                                        (process_data_signal$Signal_process_less_NC <= 0), ]$Signal_process_less_NC) != 0) {
  process_data_signal[!is.na(process_data_signal$Signal_process_less_NC) &
                                 (process_data_signal$Signal_process_less_NC <= 0), ]$process_kon = NA
  
}
if (length(process_data_signal[!is.na(process_data_signal$Signal_process_less_NC) &
                                        (process_data_signal$Signal_process_less_NC <= 0), ]$Signal_process_less_NC) != 0) {
  process_data_signal[!is.na(process_data_signal$Signal_process_less_NC) &
                                 (process_data_signal$Signal_process_less_NC <= 0), ]$process_koff = NA
  
}
if (length(process_data_signal[!is.na(process_data_signal$Signal_process_less_NC) &
                                        (process_data_signal$Signal_process_less_NC <= 0), ]$Signal_process_less_NC) != 0) {
  process_data_signal[!is.na(process_data_signal$Signal_process_less_NC) &
                                 (process_data_signal$Signal_process_less_NC <= 0), ]$process_KD = NA
  
}
if (length(process_data_signal[!is.na(process_data_signal$Signal_process_less_NC) &
                                        (process_data_signal$Signal_process_less_NC <= 0), ]$Signal_process_less_NC) != 0) {
  process_data_signal[!is.na(process_data_signal$Signal_process_less_NC) &
                                 (process_data_signal$Signal_process_less_NC <= 0), ]$process_Rmax = NA
  
}
if (length(process_data_signal[!is.na(process_data_signal$Signal_process_less_NC) &
                                        (process_data_signal$Signal_process_less_NC <= 0), ]$Signal_process_less_NC) != 0) {
  process_data_signal[!is.na(process_data_signal$Signal_process_less_NC) &
                                 (process_data_signal$Signal_process_less_NC <= 0), ]$process_Res = NA
  
}

print(str(process_data_signal))

#set parameters to NA in process table if koff is more then  koff_boundary
 if (length(process_data_signal[!is.na(process_data_signal$process_koff) &
                                         (process_data_signal$process_koff > koff_boundary), ]$process_koff) != 0) {
   process_data_signal[!is.na(process_data_signal$process_koff) &
                                  (process_data_signal$process_koff > koff_boundary), ]$process_kon = NA
   
 }
 if (length(process_data_signal[!is.na(process_data_signal$process_koff) &
                                         (process_data_signal$process_koff > koff_boundary), ]$process_koff) != 0) {
   process_data_signal[!is.na(process_data_signal$process_koff) &
                                  (process_data_signal$process_koff > koff_boundary), ]$process_KD = NA
   
 }
 if (length(process_data_signal[!is.na(process_data_signal$process_koff) &
                                         (process_data_signal$process_koff > koff_boundary), ]$process_koff) != 0) {
   process_data_signal[!is.na(process_data_signal$process_koff) &
                                  (process_data_signal$process_koff > koff_boundary), ]$process_Rmax = NA
   
 }
 if (length(process_data_signal[!is.na(process_data_signal$process_koff) &
                                         (process_data_signal$process_koff > koff_boundary), ]$process_koff) != 0) {
   process_data_signal[!is.na(process_data_signal$process_koff) &
                                  (process_data_signal$process_koff > koff_boundary), ]$process_Res = NA
   
 }
 if (length(process_data_signal[!is.na(process_data_signal$process_koff) &
                                         (process_data_signal$process_koff > koff_boundary), ]$process_koff) != 0) {
   process_data_signal[!is.na(process_data_signal$process_koff) &
                                  (process_data_signal$process_koff > koff_boundary), ]$process_koff = NA
   
 }

#make table with top concentration used for process data
signal.three.process = signal.one.process[order(signal.one.process$Conc_M_process, decreasing = TRUE), ]

signal.topconc.process = signal.three.process[!duplicated(sprintf(
  "%s-%s",
  signal.three.process$ROI,
  signal.three.process$Analyte
)), ]

names(signal.topconc.process) = c(
  "ROI",
  "Ligand",
  "Analyte",
  "Top_Conc_M_process",
  "Top_conc_Signal_process"
)

#get equilibrium data
Equilibrium_data = read.xlsx(xlsxFile = batch_file, sheet = 4, startRow = 2, colNames = FALSE, cols = 1:9)

names(Equilibrium_data) = c(
  "ROI",
  "Bay",
  "Position",
  "Ligand",
  "Analyte",
  "Equilibrium_KD",
  "Equilibrium_KD_error",
  "Equilibrium_Rmax",
  "Equilibrium_Rmax_error"
)

Equilibrium_data <- Equilibrium_data[-c(2, 3, 7, 9)]

Equilibrium_data$ROI = as.numeric(Equilibrium_data$ROI)

Equilibrium_data$Equilibrium_KD = as.numeric(Equilibrium_data$Equilibrium_KD)

Equilibrium_data$Equilibrium_Rmax = as.numeric(Equilibrium_data$Equilibrium_Rmax)

#get abs data
abs      = read.xlsx(xlsxFile = batch_file, sheet = 12, rows = 4:AbsRowMax, colNames = FALSE, cols = 1:8)


names(abs) = c("ROI", "Ligand", "Track", "C4", "C5", "Level", "SD", "Slope")

abs <- abs[-c(3, 4, 5, 7, 8)]

abs$ROI = as.numeric(abs$ROI)

abs$Level = as.numeric(abs$Level)


#make table with max level of ligand capture for each ROI
abs = abs[order(abs$Level, decreasing = TRUE), ]

abs = abs[!duplicated(sprintf(
  "%s-%s",
  abs$ROI,
  abs$Ligand
)), ]

abs$Ligand_MW <- LigandMW

AnalytesMW_data      = read.csv(
  AnalytesMW_file,
  header = T,
  sep = "\t",
  fileEncoding = "UTF-8"
)

names(AnalytesMW_data) = c("Analyte", "Analyte_MW")

AnalytesMW_data$Analyte_MW = as.numeric(AnalytesMW_data$Analyte_MW)

#combine all tables into one
one.table = merge(Equilibrium_data, Batch_data_signal, by = c("ROI", "Ligand", "Analyte"))

two.table = merge(one.table, process_data_signal, by = c("ROI", "Ligand", "Analyte"))

three.table = merge(two.table, signal.topconc.batch, by = c("ROI", "Ligand", "Analyte"))

four.table = merge(three.table,
                            signal.topconc.process,
                            by = c("ROI", "Ligand", "Analyte"))

five.table = merge(four.table, abs, by = c("ROI", "Ligand"))

seven.table = merge(five.table, AnalytesMW_data, by = "Analyte")

#make a vector with analyte max signals by combining values from batch and process columns
Signal = as.numeric(as.vector(
  ifelse(
    is.na(seven.table$batch_KD),
    seven.table$Signal_process,
    seven.table$Signal_batch
  )
))


# Compute expected Rmax:
LigandLevel = as.numeric(as.vector(seven.table$Level))

AnalyteMW   = as.numeric(as.vector(seven.table$Analyte_MW))

LigandMW = as.numeric(as.vector(seven.table$Ligand_MW))

Expected.Rmax = LigandLevel / LigandMW * LigandValency * AnalyteMW

# Put all data into one object (data frame)
df = data.frame(
  ROI = seven.table$ROI,
  Ligand = seven.table$Ligand,
  LigandMW = LigandMW,
  Analyte = seven.table$Analyte,
  AnalyteMW = AnalyteMW,
  AnalyteConc = seven.table$Top_Conc_M_batch,
  Kon = as.numeric(as.vector(
    ifelse(
      is.na(seven.table$batch_kon),
      seven.table$process_kon,
      seven.table$batch_kon
    )
  )),
  Koff = as.numeric(as.vector(
    ifelse(
      is.na(seven.table$batch_koff),
      seven.table$process_koff,
      seven.table$batch_koff
    )
  )),
  KD = as.numeric(as.vector(
    ifelse(
      is.na(seven.table$batch_KD),
      seven.table$process_KD,
      seven.table$batch_KD
    )
  )),
  RmaxFit = as.numeric(as.vector(
    ifelse(
      is.na(seven.table$batch_Rmax),
      seven.table$process_Rmax,
      seven.table$batch_Rmax
    )
  )),
  RmaxExp = Expected.Rmax,
  RmaxRatio = Expected.Rmax / as.numeric(as.vector(
    ifelse(
      is.na(seven.table$batch_Rmax),
      seven.table$process_Rmax,
      seven.table$batch_Rmax
    )
  )),
  Signal = Signal,
  Level = LigandLevel,
  SignalRatio = as.numeric(as.vector(Signal)) / as.numeric(as.vector(Expected.Rmax)),
  Chi2 = as.numeric(as.vector(
    ifelse(
      is.na(seven.table$batch_Res),
      seven.table$process_Res,
      seven.table$batch_Res
    )
  ))
)

# Compute a normalized Chi2 with Average Chi2 for the group.
df$NormChi2 = df$Chi2 / df$Signal


# Add Equilibrium fields
df$Rmax_equil = seven.table$Equilibrium_Rmax

df$KD_equil   = as.numeric(as.vector(seven.table$Equilibrium_KD))

# Add some question columns
df$RmaxYesNo   = ifelse(df$RmaxRatio >= rmax_mat_low &
                                   df$RmaxRatio < rmax_mat_high,
                                 "YES",
                                 "NO")


# Ranges taken from bc4000 technical documents. For low molecular weight compounds, on rate must be < 1*10^6 (probably put some flags in here based on
df$KonInRange  = ifelse(df$Kon  > 1 * 10 ^ 3 &
                                   df$Kon < 1 * 10 ^ 9, "YES", "NO")

df$KoffInRange = ifelse(df$Koff  > 5 * 10 ^ -5 &
                                   df$Koff < 0.5, "YES", "NO")

df$Capture     = ifelse(df$Level  < Level_low |
                                   df$RmaxExp < RmaxExp_low , "LOW", "GOOD")
# might need to update this value, I've seen GT8-GL abs go LOW here, but curves look real.
df$CurveFit    = ifelse(
  df$NormChi2 > 0 &
    df$NormChi2 <= 0.2 ,
  "GOOD",
  ifelse(df$NormChi2 > 0.2 & df$NormChi2 < 0.5, "FAIR", "POOR")
)

df$EquilYesNo  = ifelse(df$RmaxYesNo == "YES" &
                                   df$Capture == "GOOD" & df$Koff > 0.5,
                                 "YES",
                                 "NO")

df$KD_fix  = ifelse(is.na(df$KD) |
                               (df$KD > (df$AnalyteConc * 5)), NA , df$KD)

df$KD_fix  = ifelse(df$Koff >= 0.5 & 
                      ((df$KD/df$KD_equil > 3) | (df$KD/df$KD_equil < 0.33)), NA , df$KD_fix)

df$estimated = is.na(df$KD_fix) & (df$Capture == "GOOD")
df$drop = (df$Capture == "LOW")
df$KD_fix = ifelse(df$estimated, df$AnalyteConc, df$KD_fix)

df = df[order(df$ROI),]

# Formatting

df.final = format(df, scientific = 1, digits = 4)

csvName = sprintf("%s%s%s%s.csv", projectName, "_", Sys.Date(), outName)

print("df.final has this number of rows:")
print(nrow(df.final))

print("csvName is:")
print(csvName)

write.csv(x = df.final, file = csvName, row.names = FALSE)

print(negat_contr_max.batch)
print(negat_contr_max.process)

# csvName.sum = sprintf("%s%s%s%s%s.csv", projectName, "_", Sys.Date(), outName, "_summarised")
# 
# df.final.sum = df.final %>%
#   group_by(Ligand, Analyte, estimated) %>%
#   summarise(meanKD = mean(as.numeric(KD_fix)), minKD = min(KD_fix), maxKD = max(KD_fix), n =n())
#write.csv(x = df.final.sum, file = csvName.sum, row.names = FALSE)