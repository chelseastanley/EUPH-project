### Created by Jessica Nephin
### Modified by Chelsea Stanley
### Last edited May 6, 2020

### Uses COM objects to run Echoview
### Exports regions and lines from EV files
### Imports regions and lines into new EV files with new template

### Script runs from location ~Survey/Vessel/Rscripts/EVfiles
#-----------------------------------------------------------------------------------------


# required packages 
require(RDCOMClient)
require(EchoviewR)
require(tidyr)
require(gtools)
require(dplyr)

# set relative working directory
#setwd('..'); setwd('..')


###############################################
                # INPUT #
###############################################


# Location of original EV files
EVdir <- "Acoustics/Echoview/EUPH/EUPH_EVfiles/EK80"

# Create EUPH folder
 EUPH_folder <- "Acoustics/Echoview/EUPH"
 dir.create(file.path(getwd(), EUPH_folder))

# Location for files updated to EUPH template
EUPH_template <- "Acoustics/Echoview/EUPH/EUPH_EVfiles"
dir.create(file.path(getwd(), EUPH_template))

#tempate name and location
template <- "TULLY_EUPH_template_5FREQ_EK80.EV"
Tempdir <- "../../../../Pacific Region Acoustics Field Directory/Rscripts/EUPH/Templates/Tully"

# Name of the bottom line in original EV files
EVbottom <- FALSE
EVbottomname <- "EV bottom pick to edit"
#EVbottomnameNIGHT <- "0.5m bottom offset"

# Does the new template include a bottom line? What is it's name? 
bottomline <- TRUE
bottomname <- "EV bottom pick to edit"

# Variable used to see if files are EK60 or EK80 
# VarEK60 <- "Fileset 1: Sv raw pings T1"
# VarEK80 <- "Fileset 1: Sv pings T1"
###############################################




###################################################
                 # Locate #

#location of calibration file (.ecs)
CALdir <- "Acoustics/Echoview"
calName<- EVfile.list <- list.files(file.path(getwd(),CALdir), pattern=".ecs")


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

# location for day files 
  Day <- "Acoustics/Echoview/EUPH/EUPH_EVfiles/EK80/Day"
  dir.create(file.path(getwd(), Day))
# 
# # location for day files
  Night <- "Acoustics/Echoview/EUPH/EUPH_EVfiles/EK80/Night"
  dir.create(file.path(getwd(), Night))




#########################
# list the EV files to run
EVfile.list <- list.files(file.path(getwd(),EVdir), pattern="*.EV")

### move old ev files to old template directory
# file.copy(file.path(getwd(), EVdir, EVfile.list), file.path(getwd(), EVOlddir))
#file.remove(file.path(getwd(), EVdir, EVfile.list))


###############################################
#           Open EV file to update            #
###############################################


# create COM connection between R and Echoview

EVApp <- COMCreate("EchoviewCom.EvApplication")


for (i in EVfile.list){
  
  # EV filename
  name <- sub(".EV","",i)
  EVfileName <- file.path(getwd(),EVdir, i)
  
  
  
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
  # calPath <- filesetObj$GetCalibrationFileName()
  # calName <- sub(".*\\\\|.*/","",calPath)
  
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
 
  # set line status to good
  bottomline <- read.table(bottomfilename, skip=2,colClasses = "character", fileEncoding="UTF-8-BOM")
  bottomline$V4<-3 #replace with all 3's
  range(bottomline$V4)
  bottomline <- unite(bottomline, "V1", sep=' ', remove=T)
  EVLheader <- read.delim(bottomfilename, header=F, nrows=2, fileEncoding="UTF-8-BOM")
  finalEVL <- bind_rows(EVLheader, bottomline)
  write.table(finalEVL, bottomfilename, quote=F, row.names=F, col.names=F)
}

  
  # Close EV file
  EVApp$CloseFile(EVfile)
  

  
  
  #####################################
  #          Make EV file             #
  #####################################
  
  
  # Open template EV file
  EVfile <- EVApp$OpenFile(file.path(getwd(), Tempdir, template))
  
  # Set fileset object
  filesetObj <- EVfile[["Filesets"]]$Item(0)
  
  # Set calibration file
  add.calibration <- filesetObj$SetCalibrationFile(file.path(getwd(),CALdir, calName))
  
  
  # Add raw files
  for (r in raws){
    filesetObj[["DataFiles"]]$Add(file.path(getwd(),RAWdir,r))
  }
  
  # Add regions
  EVfile$Import(regionfilename)
  
 
  
  # number of editable lines in template
  # ls <- NULL
   linesObj <- EVfile[["Lines"]]
  # for(k in 0:(linesObj$Count()-1)){
  #   tmp <- linesObj$Item(k)
  #   linedit <- tmp$AsLineEditable()
  #   ls <- c(ls,linedit)
  # }
  # linenum <- length(ls)
  
  # Repick 120kHz line
  limit120 <- EVNewFixedDepthLine(EVfile, depth = 300, "New 120 limit")
  oldlimit120 <- linesObj$FindByName("120kHz limit")
  oldlimit120$OverwriteWith(limit120)
  linesObj$Delete(limit120)
  
  # Add bottom line and overwrite template bottom line if it exists
  if(EVbottom==TRUE){
    EVImportLine(EVfile, bottomfilename, "EV bottom pick to edit", T)

  }
  
    if(EVbottom==FALSE){
     # Add bottom line and overwrite template bottom line if it exists
     EVVar<-EVfile[["Variables"]]$FindByName("Fileset1: Sv raw pings T1")
     #bottom<- EVfile[["Lines"]]$CreateLinePick(EVVar, T)
     #linesObj <- EVfile[["Lines"]]
     EVLine <- linesObj$FindbyName('EV bottom pick to edit')
     bottom <- EVfile[["Lines"]]$CreateOffsetLinear(EVLine, 1, -1, TRUE)
     bottom[["Name"]]<-'Bottom'
     linesObj <- EVfile[["Lines"]]
     oldbottom <- linesObj$FindbyName('EV bottom pick to edit')
     oldbottom$OverwriteWith(bottom)
     linesObj$Delete(bottom)
  
  
     # Export picked line and set line status to GOOD and then re-import
     linesObj <- EVfile[["Lines"]]
     bottom <- linesObj$FindbyName("EV bottom pick to edit")
     bottomfilename <- file.path(getwd(),Line, paste(name, "bottom", "evl", sep="."))
     bottom$Export(bottomfilename)
     
      # set line status to good
     bottomline <- read.table(bottomfilename, sep="\t", fill=T, colClasses = "character", fileEncoding="UTF-8-BOM")
     if((as.numeric(bottomline[2,1]))>0 ){
     bottomline <- read.table(bottomfilename, skip=2, colClasses = "character", fileEncoding="UTF-8-BOM")
     bottomline$V4<-3 #replace with all 3's
     range(bottomline$V4)
     bottomline <- unite(bottomline, "V1", sep=' ', remove=T)
     EVLheader <- read.delim(bottomfilename, header=F, nrows=2, fileEncoding="UTF-8-BOM")
     finalEVL <- bind_rows(EVLheader, bottomline)
     write.table(finalEVL, bottomfilename, quote=F, row.names=F, col.names=F)
  
  
    # reimport the bottom line
      EVImportLine(EVfile, bottomfilename, "EV bottom pick to edit", T)
   }
  

}
  
  # Save EV file
  if (grepl("DAY", i) == T){
    EVfile$SaveAS(file.path(getwd(),Day,i))
    }
    if (grepl("NIGHT", i) == T){
      EVfile$SaveAS(file.path(getwd(),Night,i))
  }
 
 #  # Save EV file
 #  EVfile$SaveAS(file.path(getwd(),EUPH_template,i))
  
  # Close EV file
  EVApp$CloseFile(EVfile)
}

  # Quit echoview
  EVApp$Quit()
  

  


