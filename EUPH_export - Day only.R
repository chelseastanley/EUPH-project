### Created by Chelsea Stanley
### Last edited March 15, 2018

### Uses COM objects to run Echoview
### Integrates "EUPH mask" variable to export EUPH NASC

## Updated - Sept 24, 2021 - added bottom check code - looks between
## "Final Bottom" and "Bottom Check" lines and exports the NASC values to allow
## for quick bottom intrusion checking

#########################################################


# required packages
require(RDCOMClient)
require(dplyr)
require(stringr)
require(StreamMetabolism)
require(lubridate)
require(EchoviewR)

###############################################
#             Load in cruisetrack             #
###############################################

#location of cruise track
Trackdir <- "Acoustics/Echoview/Exports/GPSTrack"
dir.create(file.path(getwd(), Trackdir))

# load cruise track 
track <- read.csv(file.path(Trackdir,"SlimCruiseTrack.csv"))

# remove bad fixes
track <- track[!track$GPS_status == 1,]

#set variable
var <- "EUPH export 120kHz - Day Only"

# Create timestamp
now <- Sys.time()
TimeStamp <- paste0(format(now, "%Y%m%d_%H%M"))

###############################################
#              Sunset - Sunrise               #
###############################################

daytrack <- track[!duplicated(track$GPS_date),]

riseset <- NULL
for (i in 1:nrow(daytrack)){
  r <- sunrise.set(daytrack$Latitude[i], daytrack$Longitude[i], daytrack$GPS_date[i], 
                   timezone = "UTC", num.days = 1)
  riseset <- rbind(riseset, r)
}

riseset <- cbind(daytrack$GPS_date, riseset)
names(riseset)[1] <- paste("Date")
riseset$Date <- format(as.Date(riseset$Date, format = "%Y-%m-%d"), "%Y%m%d" )

# add  buffer
riseset$sunrisebuffer <- riseset$sunrise + 3600
riseset$sunsetbuffer <- riseset$sunset - 1800

riseset$sunrisebuffer <-  format(as.POSIXct(riseset$sunrisebuffer, tz="UTC"),"%Y-%m-%d %H:%M")
riseset$sunsetbuffer <-  format(as.POSIXct(riseset$sunsetbuffer, tz="UTC"),"%Y-%m-%d %H:%M")
riseset$sunrange <- paste0(as.character(riseset$sunrisebuffer), " = ", as.character(riseset$sunsetbuffer))


####################################
# EV files
####################################

# location of EV files

EUPH_EV <- "Acoustics/Echoview/EUPH/EUPH_EVfiles"

# Create EUPH export folder
EUPH_exp <- paste0("Acoustics/Echoview/EUPH/EUPH_exports", TimeStamp)
dir.create(file.path(getwd(), EUPH_exp))

# Create a bottom check export folder
BOT_EXP <- paste0("Acoustics/Echoview/EUPH/EUPH_exports/Bottom check", TimeStamp)
dir.create(file.path(getwd(), BOT_EXP))

# Create EUPH export folder
# EUPHVAR_exp <- paste0("Acoustics/Echoview/EUPH/EUPH_exports_15-50m/",var)
# dir.create(file.path(getwd(), EUPHVAR_exp))

# delete any previous files in exports folder
EUPH_oldexports <- list.files(file.path(getwd(), EUPH_exp, var), pattern = ".csv")
file.remove(file.path(getwd(), EUPH_exp, EUPH_oldexports))



#list the EV files to integrate

EVfile.list <- list.files(file.path(getwd(), EUPH_EV), pattern = ".EV")



# Loop through EV files 

for (i in EVfile.list){
  
  # create COM connection between R and Echoview
  EVApp <- COMCreate("EchoviewCom.EvApplication")
  
  # EV filenames to open
  EVfileNames <- file.path(getwd(), EUPH_EV, i)
  EvName <- strsplit(i, split = '*.EV')[[1]]
  EvDate <- trimws(strsplit(i, split = '*.EV')[[1]])
  
  # open EV file
  EVfile <- EVApp$OpenFile(EVfileNames)
  
  # set all variables analysis area
  n.var <- EVfile[["Variables"]]$Count()
  for (j in 1:n.var){
    variable = EVfile[["Variables"]]$Item(n.var-j)
    # Set analysis lines
    Obj_propA <- variable[['Properties']][['Analysis']]
    Obj_propA[['ExcludeAboveLine']]<-"15m surface blank"
    Obj_propA[['ExcludeBelowLine']]<-"Final bottom" 
    
  }
  

  # set variable
  tday <- EVfile[["Variables"]]$FindByName(var)$AsVariableAcoustic()
 
    
  # Set analysis lines
    # Obj_propA <- tday[['Properties']][['Analysis']]
    # Obj_propA[['ExcludeAboveLine']]<-"15 m surface blank"
    # Obj_propA[['ExcludeBelowLine']]<-"50m EUPH exclusion" 
    
    # Set analysis grid and exclude lines on Sv data
    Obj_propGrid <- tday[['Properties']][['Grid']]
    Obj_propGrid$SetDepthRangeGrid(1, 10)
    Obj_propGrid$SetTimeDistanceGrid(3, 0.5)
 
    # set analysis variables
    #source("../../../../Pacific Region Acoustics Field Directory/Rscripts/SetBiomassExpParams.R")
    
    # specify ping ranges
    # tday <- tday[["Properties"]][["PingSubset"]]
    # tday[["Ranges"]] = riseset$sunrange[which(EvDate == riseset$Date)]
    
    # export by cells
    exportcells <- file.path(getwd(), EUPH_exp, paste(EvName, var, "cells.csv", sep="_"))
    EVExportIntegrationByCells(EVfile, var, exportcells)
    
    # Set analysis grid and exclude lines on Sv data back to original values
    # Obj_propGrid<-tday[['Properties']][['Grid']]
    # Obj_propGrid$SetDepthRangeGrid(1, 50)
    # Obj_propGrid$SetTimeDistanceGrid(3, 0.5)
    
  # Export between Final bottom and Bottom check
    
    # Set analysis lines
     Obj_propA <- tday[['Properties']][['Analysis']]
     Obj_propA[['ExcludeAboveLine']]<-"Final bottom"
     Obj_propA[['ExcludeBelowLine']]<-"Bottom check" 
     
     # export between these two lines
     exportBC <- file.path(getwd(), BOT_EXP, paste(EvName, var, "_BottomCheck.csv", sep="_"))
     EVExportIntegrationByCells(EVfile, var, exportBC)
  
  # save EV file
  EVfile$Save()

  #close EV file
  EVApp$CloseFile(EVfile)
  
}
  
  #quit echoview
  EVApp$Quit()


## ------------- end loop




#####################################################
# Combine all Export .csv
#####################################################
  
  cruise <- strsplit(getwd(), "/")[[1]][4]
  vessel <- strsplit(getwd(), "/")[[1]][5]    
    
  # export name
    expname <- paste0(cruise,"_",var, "_", vessel, "_", TimeStamp, ".csv")
    
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
