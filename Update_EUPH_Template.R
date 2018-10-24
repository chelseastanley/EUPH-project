### Created by Jessica Nephin
### Modified by Chelsea Stanley
### Last edited July 3, 2018

### Uses COM objects to run Echoview
### Exports regions and lines from EV files
### Imports regions and lines into new EV files with new template

### Script runs from location ~Survey/Vessel/Rscripts/EVfiles
#-----------------------------------------------------------------------------------------


# required packages 
require(RDCOMClient)

# set relative working directory
setwd('..');setwd('..')


###############################################
                # INPUT #
###############################################


# Location of original EV files
EVdir <- "Acoustics/Echoview/Day-files"

# Create EUPH folder
EUPH_folder <- "Acoustics/Echoview/EUPH"
dir.create(file.path(getwd(), EUPH_folder))

# Location for files updated to EUPH template
EUPH_template <- "Acoustics/Echoview/EUPH/EUPH_EVfiles"
dir.create(file.path(getwd(), EUPH_template))

#tempate name and location
template <- "TULLY_EUPH_template_5FREQ.EV"
Tempdir <- "Rscripts/EUPH/Templates/Tully"

# Name of the bottom line in original EV files
EVbottom <- F
EVbottomname <- "Smoothed bottom to edit"

# Does the new template include a bottom line? What is it's name?
bottomline <- TRUE
bottomname <- "EV bottom pick to edit"


###############################################




###################################################
                 # Locate #

#location of calibration file (.ecs)
CALdir <- "Acoustics/Echoview"

#location of .raw files
RAWdir <- "Acoustics/RAW"

#location for Exports
Exports <- "Acoustics/Echoview/Exports"
dir.create(file.path(getwd(), Exports))

#location for region exports
Reg <- "Acoustics/Echoview/Exports/Regions"
dir.create(file.path(getwd(), Reg))

#location for line exports
Line <- "Acoustics/Echoview/Exports/Lines"
dir.create(file.path(getwd(), Line))



#########################
# list the EV files to run
EVfile.list <- list.files(file.path(getwd(),EVdir), pattern=".EV")

### move old ev files to old template directory
# file.copy(file.path(getwd(), EVdir, EVfile.list), file.path(getwd(), EVOlddir))
# file.remove(file.path(getwd(), EVdir, EVfile.list))


###############################################
#           Open EV file to update            #
###############################################

for (i in EVfile.list){
  
  # EV filename
  name <- sub(".EV","",i)
  EVfileName <- file.path(getwd(),EVdir, i)
  
  # create COM connection between R and Echoview
  EVApp <- COMCreate("EchoviewCom.EvApplication")
  
  # open EV file
  EVfile <- EVApp$OpenFile(EVfileName)
  
  # Set fileset object
  filesetObj <- EVfile[["Filesets"]]$Item(0)
  
  # list raw files
  num <- filesetObj[["DataFiles"]]$Count()
  raws <- NULL
  for (l in 0:(num-1)){
    dataObj <- filesetObj[["DataFiles"]]$Item(l)
    dataPath <- dataObj$FileName()
    dataName <- sub(".*\\\\|.*/","",dataPath)
    raws <- c(raws,dataName) 
  }
  
  # get .ecs filename
  calPath <- filesetObj$GetCalibrationFileName()
  calName <- sub(".*\\\\|.*/","",calPath)
  
  # export .evr file
  # filename
  regionfilename <-  file.path(getwd(),Reg, paste(name, "evr", sep="."))
  # export
  EVfile[["Regions"]]$ExportDefinitionsAll(regionfilename)
  
  # export bottom line
  if(EVbottom == TRUE){
  linesObj <- EVfile[["Lines"]]
  bottom <- linesObj$FindbyName(EVbottomname)
  bottomfilename <- file.path(getwd(),Line, paste(name, "bottom", "evl", sep="."))
  bottom$Export(bottomfilename)
  }

  

  #quit echoview
  
  EVApp$Quit()
  

  
  
  #####################################
  #          Make EV file             #
  #####################################
  
  # create COM connection between R and Echoview
  EVApp <- COMCreate("EchoviewCom.EvApplication")
  
  # Open template EV file
  EVfile <- EVApp$OpenFile(file.path(getwd(), Tempdir, template))
  
  # Set fileset object
  filesetObj <- EVfile[["Filesets"]]$Item(0)
  
  # Set calibration file
  if(!calPath == ""){
  add.calibration <- filesetObj$SetCalibrationFile(file.path(getwd(),CALdir, calName))
  }
  
  # Add raw files
  for (r in raws){
    filesetObj[["DataFiles"]]$Add(file.path(getwd(),RAWdir,r))
  }
  
  # Add regions
  EVfile$Import(regionfilename)
  
  # number of editable lines in template
  ls <- NULL
  linesObj <- EVfile[["Lines"]]
  for(k in 0:(linesObj$Count()-1)){
    tmp <- linesObj$Item(k)
    linedit <- tmp$AsLineEditable()
    ls <- c(ls,linedit)
  }
  linenum <- length(ls)
  
  # Add bottom line and overwrite template bottom line if it exists
  if(EVbottom==TRUE){
  EVfile$Import(bottomfilename)
  bottom <- linesObj$FindbyName(name)
  linenum <- linenum + 1
  if(bottomline == TRUE){
    oldbottom <- linesObj$FindbyName(bottomname)
    oldbottom$OverwriteWith(bottom)
    linesObj$Delete(bottom)
  } else if(bottomline == FALSE){
    bottom[["Name"]] <- "Bottom"
  }
  }
  if(EVbottom==FALSE){
    EVVar<-EVfile[["Variables"]]$FindByName("Fileset1: Sv raw pings T2")
    bottom<- EVfile[["Lines"]]$CreateLinePick(EVVar, F)
    if(bottomline == TRUE){
      oldbottom <- linesObj$FindbyName(bottomname)
      oldbottom$OverwriteWith(bottom)
      linesObj$Delete(bottom)
    } else if(bottomline == FALSE){
      bottom[["Name"]] <- "Bottom"
    }
  }
  
  
  
  
  # Save EV file
  EVfile$SaveAS(file.path(getwd(),EUPH_template,i))
  
  # Close EV file
  EVApp$CloseFile(EVfile)
  
  # Quit echoview
  EVApp$Quit()
  
}
