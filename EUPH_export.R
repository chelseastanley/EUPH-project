### Created by Chelsea Stanley
### Last edited March 15, 2018

### Uses COM objects to run Echoview
### Integrates "EUPH mask" variable to export EUPH NASC

#########################################################

# required packages
require(RDCOMClient)
require(dplyr)
require(stringr)

# set the working directory
setwd('..'); setwd('..')

# location of EV files

EUPH_EV <- "Acoustics/Echoview/Day-files/EUPH/SchoolDetection"

# Where to put exports
EUPH_export <- "Acoustics/Echoview/Day-files/EUPH/EUPH_Exports"
dir.create(file.path(getwd(), EUPH_export))

#list the EV files to integrate

EVfile.list <- list.files(file.path(getwd(), EUPH_EV), pattern="*school_detection.EV")

# Loop through EV files 

for (i in EVfile.list){
  # create COM connection between R and Echoview
  EVApp <- COMCreate("EchoviewCom.EvApplication")
  
  # EV filenames to open
  EVfileName <- file.path(getwd(), EUPH_EV, i) 
  
  # Export filename
  name <- paste(str_extract(EVfileName,"2017[0-9]{4}_(DAY|NIGHT)"),
                "EUPH_export.csv",sep="_")
  ExportFilename <- file.path(getwd(), EUPH_export, name)
  
  # open EV file
  EVfile <- EVApp$OpenFile(EVfileName)
  
  # Define mask variable object
  Obj <- EVfile[["Variables"]]$FindByName("EUPH mask")$AsVariableAcoustic()

  
  # Define EUPH region class
  EUPHObj <- EVfile[["RegionClasses"]]$FindByName("Euphausiid")
  
  
  
  # export by region by cell
  Obj$ExportIntegrationByRegionsByCells(ExportFilename, EUPHObj)
  
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


