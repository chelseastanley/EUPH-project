### Created by Jessica Nephin
### Last edited Sep 09, 2016

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
EVOlddir <- "Acoustics/Echoview/Day-files/Original files"
EVdir <- "Acoustics/Echoview/Day-files"
dir.create(file.path(getwd(), EVOlddir))


#location for new EV files
EV_updated <- "Acoustics/Echoview/Day-files/EUPH template/Updated template"
dir.create(file.path(getwd(), EVnew))

#tempate name and location
template <- "EUPH template - CS.EV"
Tempdir <- "Acoustics/Echoview/Other/Templates"

# Name of the bottom line in EV files
EVbottom <- "1.0 m bottom offset"

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
file.copy(file.path(getwd(), EVdir, EVfile.list), file.path(getwd(), EVOlddir))
file.remove(file.path(getwd(), EVdir, EVfile.list))


###############################################
#           Open EV file to update            #
###############################################

for (i in EVfile.list){
  
  # EV filename
  name <- sub(".EV","",i)
  EVfileName <- file.path(getwd(),EVOlddir, i)
  
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
  linesObj <- EVfile[["Lines"]]
  bottom <- linesObj$FindbyName(EVbottom)
  bottomfilename <- file.path(getwd(),Line, paste(name, "bottom", "evl", sep="."))
  bottom$Export(bottomfilename)

  # export other editable lines
  ls <- NULL
  for(k in 0:(linesObj$Count()-1)){
    tmp <- linesObj$Item(k)
    linedit <- tmp$AsLineEditable()
    if(!is.null(linedit)){
      ls <- c(ls,linedit$Name())
    }
  }
  
  linesExp <- ls[!(ls %in% c(EVbottom,"120 kHz range limit"))]
  if(length(linesExp) > 0){
    for(l in linesExp){
      lobj <- linesObj$FindbyName(l)
      linefilename <- file.path(getwd(),Line, paste(name, l, "evl", sep="."))
      lobj$Export(linefilename)
    }
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
  EVfile$Import(bottomfilename)
  bottom <- linesObj$FindbyName(paste0("Line",linenum+1))
  linenum <- linenum + 1
  if(bottomline == TRUE){
    oldbottom <- linesObj$FindbyName(bottomname)
    oldbottom$OverwriteWith(bottom)
    linesObj$Delete(bottom)
  } else if(bottomline == FALSE){
      bottom[["Name"]] <- "Bottom"
  }
  
  
  # Add other lines
  if(length(linesExp) > 0){
    for(l in 1:length(linesExp)){
      num <- linenum + l
      linefilename <- file.path(getwd(),Line, paste(name, linesExp[l], "evl", sep="."))
      EVfile$Import(linefilename)
      trawl <- linesObj$FindbyName(paste0("Line",num))
      trawl[["Name"]] <- linesExp[l]
    }
  }
  
  # Save EV file
  EVfile$SaveAS(file.path(getwd(),EV_updated,i))
  
  # Close EV file
  EVApp$CloseFile(EVfile)
  
  # Quit echoview
  EVApp$Quit()
  
}
