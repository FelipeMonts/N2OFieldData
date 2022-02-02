##############################################################################################################
# 
# 
# Program get the field data from the excel files, compile it ans sumarize it 
# 
# 
#  Felipe Montes 2022/02/01
# 
# 
# 
# 
############################################################################################################### 



###############################################################################################################
#                             Tell the program where the package libraries are stored                        
###############################################################################################################


#  Tell the program where the package libraries are  #####################

.libPaths("C:/Felipe/SotwareANDCoding/R_Library/library")  ;


###############################################################################################################
#                             Setting up working directory  Loading Packages and Setting up working directory                        
###############################################################################################################


#      set the working directory

# readClipboard() Willow Rock Spring\\SkyCap_SelectionTrial\\DataCollection") ;   # 


setwd( "C:\\Felipe\\CCC Based Experiments\\StrategicTillage_NitrogenLosses_OrganicCoverCrops\\Data\\HandWritenData")

###############################################################################################################
#                            Install the packages that are needed                       
###############################################################################################################


#install.packages("",  dependencies = T, lib="C:/Felipe/SotwareANDCoding/R_Library/library")



###############################################################################################################
#                           load the libraries that are needed   
###############################################################################################################

library(openxlsx)

library(lattice)


###############################################################################################################
#                           Explore the files and directory and files with the data
###############################################################################################################
### Read the Files from the local directory that is synchronized with the Share point directory where the Field data is stored

#Directory path for the local directory that is synchronized with the Share point directory

Directory.path<-"C:\\Users\\frm10\\The Pennsylvania State University\\StrategicTillageAndN2O - Documents\\Data\\N2OFieldSampling" 

Files.List<-list.files(Directory.path); 

### get only the excel files

#The first file is not in the same format as the rest. I will skip it for now.

Field.Data.Excell.Files<-Files.List[grep(".xlsx", Files.List)[2:20]] ;

##############################################

#Initiallize a data frame where all the field data will be collected

Field.Data<-data.frame(Block = double() , Cover.Crop = character() ,	Treatment = character(), Leter = character(), 	Plot = character() ,	Location = character() ,	Label = character() ,	TimeStart = character(),	Soil.T..F = double(),	WC.1 = double() ,	WC.2 = double() ,	WC.3 = double(), comments = character(), File = character()) ;
str(Field.Data) ; head(Field.Data);

###Read the data

for (i in  seq(1, length(Field.Data.Excell.Files) ) ){
#i=1
  Field.Data.p1<-read.xlsx(paste0(Directory.path,"\\" , Field.Data.Excell.Files[i]), sheet= "Page 1", colNames=T , cols = c(1:13) ) ;
  str(Field.Data.p1) ; head(Field.Data.p1);
  
  if (dim(Field.Data.p1)[2]<=12) {
   
    Field.Data.p1$comments<-"No Comments"
    
  }
  
  names(Field.Data.p1)<-c("Block" ,	"Cover.Crop" ,	"Treatment", 	"Leter", 	"Plot",	"Location",	"Label" ,	"TimeStart",	"Soil.T..F",	"WC.1" ,	"WC.2" ,	"WC.3", "comments") ;
  
  Field.Data.p1$File<-Field.Data.Excell.Files[i] ;
  
  
  Field.Data.p2<-read.xlsx(paste0(Directory.path,"\\" , Field.Data.Excell.Files[i]), sheet= "Page 2", colNames=T, cols = c(1:13)) ;
  str(Field.Data.p2) ; head(Field.Data.p2);
  
  if (dim(Field.Data.p2)[2]<=12) {
    
    Field.Data.p2$comments<-"No Comments"
    
  }
  
  
  names(Field.Data.p2)<-c("Block" ,	"Cover.Crop" ,	"Treatment", 	"Leter", 	"Plot",	"Location",	"Label" ,	"TimeStart",	"Soil.T..F",	"WC.1" ,	"WC.2" ,	"WC.3" , "comments") ;
  
  Field.Data.p2$File<-Field.Data.Excell.Files[i] ;
  
  
  Field.Data.1<-rbind(Field.Data.p1,Field.Data.p2)
  str(Field.Data.1) ; head(Field.Data.1);
  
  Field.Data.2<-rbind(Field.Data, Field.Data.1) ;
  
  Field.Data<-Field.Data.2 ;
  str(Field.Data) ; head(Field.Data);

  
}


###############################################################################################################
#                           
# calculate average wc  and Analysis date
#
###############################################################################################################

# convert WC.1, WC.2 , WC.3 from character to numeric and calculate the mean

Field.Data$WC.1<-as.numeric(Field.Data$WC.1) ;

Field.Data$WC.2<-as.numeric(Field.Data$WC.2) ;

Field.Data$WC.3<-as.numeric(Field.Data$WC.3) ;

Field.Data$Mean.WC<-rowMeans(Field.Data[,c("WC.1" , "WC.2" , "WC.3")] ,  na.rm=T) ;
str(Field.Data) ; head(Field.Data);

### There Are some NA Start Time, Check Which ones

Field.Data[is.na(Field.Data$TimeStart),]

# Because TimeStart data is just written as HH:MM without AM or PM, that needs to be fixed

# Transform it into Date-time format, this will add the date of today and AM to all

TimeStart.1<-as.POSIXct(strptime(Field.Data$TimeStart, format = "%R")) ;
TimeStart.1

##### There are some NA, check which

Field.Data[which(is.na(TimeStart.1)), c("TimeStart" , "File", "Label")]

# select times that are earlier than 8:00 am and add 12 hours (12 * 60 * 60 seconds) to those times. Because there are some NA, we need to not select those (!is.na())


TimeStart.1[TimeStart.1 <= as.POSIXct(strptime(c("08:00 AM"), format = "%I:%M %p")) & !is.na(TimeStart.1)]<-TimeStart.1[TimeStart.1 <= as.POSIXct(strptime(c("08:00 AM"), format = "%I:%M %p")) & !is.na(TimeStart.1)] + (12* 60 *60) ;

TimeStart.1

## Add it back to the dataframe as TimeStar.Fix, as character with the 24 hours format HH:MM 

Field.Data$TimeStar.Fix<-strftime(TimeStart.1, format = "%H:%M" ) ;
head(Field.Data)


# get the date of analysis from the File  column 

# first get the part of the File that represents the Date ("20210601" from N2OSampling20210601)

Date.1<-substring(Field.Data$File, first = 12, last= 19) ;
head(Date.1)

# get it in  date time format (POSIXct)

Date.2<-as.Date(Date.1, format ="%Y%m%d") ;
head(Date.2)

## Add it back to the dataframe as Date

Field.Data$Date<-Date.2 ;
head(Field.Data)

## Order the data frame according to date and time to start
names(Field.Data)
Ordered.Field.Data<-Field.Data[order(Field.Data$Date, Field.Data$TimeStar.Fix, na.last=F), c("Date" , "TimeStar.Fix" , "Block" ,   "Cover.Crop" ,  "Treatment"  ,  "Leter" , "Plot" ,  "Location" , "Label", "TimeStart" , "Soil.T..F" ,   "WC.1"  ,  "WC.2" , "WC.3" ,  "comments" , "File"  )] ;


### Write it into an excell file

write.xlsx(Ordered.Field.Data, file="FieldDataSummary.xlsx", asTable=F)

