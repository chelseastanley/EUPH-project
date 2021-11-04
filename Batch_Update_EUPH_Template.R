### Created by Jessica Nephin
### Modified by Chelsea Stanley
### Last edited June 6, 2019

### Runs through all survey folders and updates templates and exports integrations
### Uses COM objects to run Echoview
### Exports regions and lines from EV files
### Imports regions and lines into new EV files with new template

### Script runs from location Pacific folder in Acoustic Survey Data
#-----------------------------------------------------------------------------------------


# required packages 
require(RDCOMClient)
require(EchoviewR)
require(tidyr)
require(gtools)
require(dplyr)
require(stringr)
library(RDCOMClient)


# set working directory
DataFolder <- setwd(paste0("Acoustic Survey Data/Pacific/"))

# acoustic variables(s) to integrate and their frequency
variables <- c("EUPH export 38kHz","EUPH export 120kHz")
frequency <- c("38","120")

# list survey folders

surveyfolderlist <- list.files()

###########################################################
# Setup parameters for running scripts
###########################################################

for (i in surveyfolderlist){
  
  #assign ship name
  ship <- (list.files(i))
  
  # Location of original EV files
  EVdir <- paste0(i, "/",ship,"/Acoustics/Echoview/EUPH/EUPH_EVfiles")
  
  # Create EUPH folder
  EUPH_folder <- paste0(i,"/", ship,"/Acoustics/Echoview/EUPH")
  dir.create(file.path(getwd(), EUPH_folder))
  
  # Location for files updated to EUPH template
  EUPH_template <- paste0(i, "/", ship,"/Acoustics/Echoview/EUPH/EUPH_EVfiles")
  dir.create(file.path(getwd(), EUPH_template))
  
  
  # Name of the bottom line in original EV files
  EVbottom <- F
  EVbottomname <- "EV bottom pick to edit"
  
  # Does the new template include a bottom line? What is it's name?
  bottomline <- TRUE
  bottomname <- "EV bottom pick to edit"
  


# set template based on ship

if (ship == "Tully"){
  template <- "TULLY_EUPH_template_5FREQ.EV"
  Tempdir <- "F:/Pacific Region Acoustics Field Directory/Rscripts/EUPH/Templates/Tully"
}
if(ship =="Ricker"){
  # assign year
  year <- strsplit(i, "-")[[1]][1]
  if (year < 2016){
  template <- "RICKER_EUPH_template_2FREQ.EV"
  Tempdir <- "F:/Pacific Region Acoustics Field Directory/Rscripts/EUPH/Templates/Ricker"
  }
  if (year >= 2016){
    template <- "RICKER_EUPH_template_3FREQ.EV"
    Tempdir <- "F:/Pacific Region Acoustics Field Directory/Rscripts/EUPH/Templates/Ricker"
  }
}
if(ship=="Nordic Pearl"){
  template <-"NP_EUPH_template_3FREQ.EV"
  Tempdir <- "F:/Pacific Region Acoustics Field Directory/Rscripts/EUPH/Templates/Nordic Pearl"
}

###################################################
# Locate #

#location of calibration file (.ecs)
CALdir <- paste0(i, "/", ship,"/Acoustics/Echoview")

#location of .raw files
RAWdir <- paste0(i, "/", ship,"/Acoustics/RAW")

#location for Exports
Exports <- paste0(i, "/", ship,"/Acoustics/Echoview/EUPH/EUPH_exports")
dir.create(file.path(getwd(), Exports))

#location for region exports
Reg <- paste0(i, "/", ship,"/Acoustics/Echoview/Exports/Regions")
dir.create(file.path(getwd(), Reg))

#location for line exports
Line <- paste0(i, "/", ship,"/Acoustics/Echoview/Exports/Lines")
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

for (j in EVfile.list){
  
  # EV filename
  name <- sub(".EV","",j)
  EVfileName <- file.path(getwd(),EVdir, j)
  
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
    
    # set line status to good
    bottomline <- read.table(bottomfilename, skip=2,colClasses = "character", fileEncoding="UTF-8-BOM")
    bottomline$V4<-3 #replace with all 3's
    range(bottomline$V4)
    bottomline <- unite(bottomline, "V1", sep=' ', remove=T)
    EVLheader <- read.delim(bottomfilename, header=F, nrows=2, fileEncoding="UTF-8-BOM")
    finalEVL <- bind_rows(EVLheader, bottomline)
    write.table(finalEVL, bottomfilename, quote=F, row.names=F, col.names=F)
  }
  
  #quit echoview
  
  EVApp$Quit()
  
  
  
  
  #####################################
  #          Make EV file             #
  #####################################
  
  # create COM connection between R and Echoview
  EVApp <- COMCreate("EchoviewCom.EvApplication")
  
  # Open template EV file
  EVfile<- EVApp$OpenFile(file.path(Tempdir,template))
  
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
    # Add bottom line and overwrite template bottom line if it exists
    EVVar<-EVfile[["Variables"]]$FindByName("Fileset1: Sv pings T1")
    bottom<- EVfile[["Lines"]]$CreateLinePick(EVVar, F)
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
    bottomline <- read.table(bottomfilename, skip=2,colClasses = "character", fileEncoding="UTF-8-BOM")
    bottomline$V4<-3 #replace with all 3's
    range(bottomline$V4)
    bottomline <- unite(bottomline, "V1", sep=' ', remove=T)
    EVLheader <- read.delim(bottomfilename, header=F, nrows=2, fileEncoding="UTF-8-BOM")
    finalEVL <- bind_rows(EVLheader, bottomline)
    write.table(finalEVL, bottomfilename, quote=F, row.names=F, col.names=F)
    
    
    # reimport the bottom line 
    EVImportLine(EVfile, bottomfilename, "EV bottom pick to edit", T)
    }
  
  
  # Repick 120kHz line
  limit120 <- EVNewFixedDepthLine(EVfile, depth = 300, "New 120 limit")
  oldlimit120 <- linesObj$FindByName("120kHz limit")
  oldlimit120$OverwriteWith(limit120)
  linesObj$Delete(limit120)
  
  
  # Save EV file
  EVfile$SaveAS(file.path(getwd(),EUPH_template,j))
  
 #################################################################
 # Run intergration
  ###############################################################
  
  
  
  # bind variable and frequency together
  vars <- data.frame(variables,frequency, stringsAsFactors = FALSE)
  
  # create folder in Exports for each variable
  for(f in variables){
    suppressWarnings(dir.create(file.path(getwd(), Exports, f)))
  }
  
    
    # Variables object
    Obj <- EVfile[["Variables"]]
    
    # loop through variables for integration
    for(v in 1:nrow(vars)){
      var <- vars$variables[v]
      freq <- vars$frequency[v]
      varac <- Obj$FindByName(var)$AsVariableAcoustic()
      
      # Set analysis lines
      Obj_propA<-varac[['Properties']][['Analysis']]
      Obj_propA[['ExcludeAboveLine']]<-"50m EUPH exclusion"
      Obj_propA[['ExcludeBelowLine']]<-"Final bottom" 
      
      # Set analysis grid and exclude lines on Sv data
      Obj_propGrid <- varac[['Properties']][['Grid']]
      Obj_propGrid$SetDepthRangeGrid(1, 10)
      Obj_propGrid$SetTimeDistanceGrid(3, 0.5)
      
      
      # export by cells
      exportcells <- file.path(getwd(), Exports, var, paste(name, freq, "cells.csv", sep="_"))
      varac$ExportIntegrationByCellsAll(exportcells)
      
      # Set analysis grid and exclude lines on Sv data back to original values
      Obj_propGrid<-varac[['Properties']][['Grid']]
      Obj_propGrid$SetDepthRangeGrid(1, 50)
      Obj_propGrid$SetTimeDistanceGrid(3, 0.5)
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
  
  for(v in 1:nrow(vars)){
    var <- vars$variables[v]
    freq <- vars$frequency[v]
    cruise <- strsplit(getwd(), "/")[[1]][4]
    
    for (k in c("cells")){
      
      # export name
      expname <- paste0(cruise,"_",freq,k,".csv")
      
      # delete old .csv
      unlink(file.path(getwd(), EUPH_exp, var, expname))
      
      # list all integration files
      nasc.list <- list.files(file.path(getwd(), EUPH_exp, var), 
                              pattern=paste0("*",freq,"_",k,".csv"))
      
      # add all integration files together
      df <- NULL
      for (n in nasc.list){
        d <- read.csv(file.path(getwd(), EUPH_exp, var, n), header = T)
        regid <- grep("Region_ID",names(d))
        procid <- grep("Process_ID",names(d))
        colnames(d)[regid] <- "Region_ID"
        colnames(d)[procid] <- "Process_ID"
        df <- rbind(df,d)
      }
      
      # export to .csv
      write.csv(df, file = file.path(getwd(), EUPH_exp, var, expname),row.names = FALSE)
      
    }
  }
}
  

  