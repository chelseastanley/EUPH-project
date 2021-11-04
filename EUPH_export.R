### Created by Chelsea Stanley
### Last edited March 15, 2018

### Uses COM objects to run Echoview
### Integrates "EUPH mask" variable to export EUPH NASC

## Updated - Sept 24, 2021 - added bottom check code - looks between
## "Final Bottom" and "Bottom Check" lines and exports the NASC values to allow
## for quick bottom intrusion checking

#########################################################

###############################################
#----------------  INPUT  --------------------#
###############################################

# acoustic variables(s) to integrate and their frequency
variables <- c("EUPH export 120kHz - Day Only")
frequency <- c("120")

# required packages
require(RDCOMClient)
require(dplyr)
require(stringr)

now <- Sys.time()
TimeStamp <- paste0(format(now, "%Y%m%d_%H%M"))

# location of EV files

EUPH_EV <- "Acoustics/Echoview/EUPH/EUPH_EVfiles"

# Create a EUPH export folder
EUPH_exp <- paste0("Acoustics/Echoview/EUPH/EUPH_exports/", TimeStamp)
dir.create(file.path(getwd(), EUPH_exp))

# Create EUPH export folder
#EUPH_exp <- paste0("Acoustics/Echoview/EUPH/EUPH_exports/", variables)
#dir.create(file.path(getwd(), EUPH_exp))


#list the EV files to integrate

EVfile.list <- list.files(file.path(getwd(), EUPH_EV), pattern = ".EV")

# bind variable and frequency together
vars <- data.frame(variables,frequency, stringsAsFactors = FALSE)

# create folder in Exports for each variable
# for(f in variables){
#   suppressWarnings(dir.create(file.path(getwd(), EUPH_exp, f)))
# }

# Loop through EV files 

for (i in EVfile.list){
  # create COM connection between R and Echoview
  EVApp <- COMCreate("EchoviewCom.EvApplication")
  
  # EV filenames to open
  EVfileNames <- file.path(getwd(), EUPH_EV, i)
  EvName <- strsplit(i, split = '*.EV')[[1]]
  
  
  # open EV file
  EVfile <- EVApp$OpenFile(EVfileNames)
 
  # Variables object
  Obj <- EVfile[["Variables"]]
  
  # loop through variables for integration
  for(v in 1:nrow(vars)){
    var <- vars$variables[v]
    freq <- vars$frequency[v]
    varac <- Obj$FindByName(var)$AsVariableAcoustic()
    
    # Set analysis lines
    Obj_propA<-varac[['Properties']][['Analysis']]
    Obj_propA[['ExcludeAboveLine']]<-"15m surface blank"
    Obj_propA[['ExcludeBelowLine']]<-"Final bottom" 
    
    # Set analysis grid and exclude lines on Sv data
    Obj_propGrid <- varac[['Properties']][['Grid']]
    Obj_propGrid$SetDepthRangeGrid(1, 10)
    Obj_propGrid$SetTimeDistanceGrid(3, 0.5)
 
    
    # export by cells
    exportcells <- file.path(getwd(), EUPH_exp, paste(EvName, var, "cells.csv", sep="_"))
    varac$ExportIntegrationByCellsAll(exportcells)
    
    # Set analysis grid and exclude lines on Sv data back to original values
    Obj_propGrid<-varac[['Properties']][['Grid']]
    Obj_propGrid$SetDepthRangeGrid(1, 50)
    Obj_propGrid$SetTimeDistanceGrid(3, 0.5)
    
    
    # Export between Final bottom and Bottom check
    
    # Set analysis lines
    Obj_propA <- tday[['Properties']][['Analysis']]
    Obj_propA[['ExcludeAboveLine']]<-"Final bottom"
    Obj_propA[['ExcludeBelowLine']]<-"Bottom check" 
    
    # export between these two lines
    exportBC <- file.path(getwd(), BOT_EXP, paste(EvName, var, "_BottomCheck.csv", sep="_"))
    EVExportIntegrationByCells(EVfile, var, exportBC)
    }

  
  
  # save EV file
  EVfile$Save()

  #close EV file
  EVApp$CloseFile(EVfile)
  
  
  #quit echoview
  EVApp$Quit()


## ------------- end loop

}

#####################################################
# Combine all Export .csv
#####################################################

cruise <- strsplit(getwd(), "/")[[1]][4]
vessel <- strsplit(getwd(), "/")[[1]][5]

now <- Sys.time()
TimeStamp <- paste0(format(now, "%Y%m%d_%H%M%S"))

# export name
expname <- paste0(cruise,"_",variables,"_", vessel,"_",TimeStamp,".csv")

# delete old .csv
unlink(file.path(getwd(), EUPH_exp, var, expname))

# list all integration files
nasc.list <- list.files(file.path(getwd(), EUPH_exp), 
                        pattern=paste0("*",var,"_cells.csv"))

# add all integration files together
df <- NULL
for (n in nasc.list){
  d <- read.csv(file.path(getwd(), EUPH_exp, n), header = T)
  # regid <- grep("Region_ID",names(d))
  # procid <- grep("Process_ID",names(d))
  # colnames(d)[regid] <- "Region_ID"
  # colnames(d)[procid] <- "Process_ID"
  df <- rbind(df,d)
}

# export to .csv
write.csv(df, file = file.path(getwd(), EUPH_exp, expname),row.names = FALSE)


# Combine all bottom check files

# export name
BCname <- paste0(cruise,"_",var, "_", vessel, "_", TimeStamp, "botcheck.csv")

# delete old .csv
#unlink(file.path(getwd(), EUPH_exp, var, expname))

# list all integration files
bot.list <- list.files(file.path(getwd(), BOT_EXP), 
                       pattern=paste0("*",var,"_BottomCheck.csv"))

# add all integration files together
df <- NULL
for (n in bot.list){
  d <- read.csv(file.path(getwd(), BOT_EXP, n), header = T)
  # regid <- grep("Region_ID",names(d))
  # procid <- grep("Process_ID",names(d))
  # colnames(d)[regid] <- "Region_ID"
  # colnames(d)[procid] <- "Process_ID"
  df <- rbind(df,d)
}

# export to .csv
write.csv(df, file = file.path(getwd(), BOT_EXP, BCname),row.names = FALSE)


