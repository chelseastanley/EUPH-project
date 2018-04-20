### Created by Chelsea Stanley
### Last edited March 13, 2018

### Uses COM objects to run Echoview
### Detects EUPH schools from the 120 kHz Euphau 120 kHz_120-38 variable  


### Script runs from location ~Survey/Vessel/Rscripts/EVfiles
#-----------------------------------------------------------------------------------------



## EDIT ##
  
# Where to find EV files

 setwd('..');setwd('..')
 EV_source <- "Acoustics/Echoview/Day-files/EUPH template/Updated template"

# Where to put files that have had schools detected
 SD_location <- "Acoustics/Echoview/Day-files/EUPH template/School detection"
 dir.create(file.path(getwd(), SD_location))

 ## DO NOT EDIT###

library(RDCOMClient)


# List all EV files
EVfiles.list <- list.files(EV_source,pattern = ".*EV$") 



for(i in EVfiles.list){
  # Establish communication with Echoview
  
  EVApp <- COMCreate("EchoviewCom.EvApplication")
  
  # Define EV filename
  EVfilename <- file.path(getwd(), EV_source, i)
  
  # Open EV file
  EVfile <- EVApp$OpenFile(EVfilename)

  # Exclude data above 15m and below bottom
  varName <- "Euphau 120 kHz_120-38"
  Obj <- EVfile[["Variables"]]$FindByName(varName)$AsVariableAcoustic()
  Obj1 <- Obj[["Properties"]][["Analysis"]]
  Obj1[["ExcludeAboveLine"]] <- "15 m surface blank"
  Obj1[["ExcludeBelowLine"]] <- "Bottom offset + max depth line"
  
  # Run school detection 
  schoolDet <- Obj$DetectSchools("Euphausiid")
  if(schoolDet == -1){
    paste(EVfilename," : School detection failed.",sep="")
  }
  
  # Save new EV file as school_detection.EV
  filename <- sub("*.EV", "", i)
  nEVfilename <- paste(filename, "_school_detection.EV", sep="")
  
  # Save file
  EVfile$SaveAs(file.path(getwd(), SD_location, nEVfilename))
  
  # Quit Echoview
  EVApp$Quit()
}
  

