##############################################################################################################
# 
# 
# Program to put together the handwritten field data imported from pdf files into digital from using AWS Textract
# 
#                                              https://aws.amazon.com/textract/
#  
# 
# 
#  Felipe Montes 2022/01/26
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

# readClipboard()   Willow Rock Spring\\SkyCap_SelectionTrial\\DataCollection") ;   # 


setwd("C:\\Felipe\\CCC Based Experiments\\StrategicTillage_NitrogenLosses_OrganicCoverCrops\\Data\\HandWritenData");

###############################################################################################################
#                            Install the packages that are needed                       
###############################################################################################################




###############################################################################################################
#                           load the libraries that are needed   
###############################################################################################################

library(openxlsx)

library(lattice)





###############################################################################################################
#                           Read the excel files into R data frames
###############################################################################################################

###############################################################################################################
#                           Read the Notepad csv file created with AWS Textract 
#                             https://aws.amazon.com/textract/
###############################################################################################################


#Create the MS excel workbook

FieldDataWkb<-createWorkbook() ;

options(openxlsx.datetimeFormat = "hh:mm")

addWorksheet(wb=FieldDataWkb, sheetName = "Page 1")

saveWorkbook(FieldDataWkb,file= "Field.Data.xlsx")



#### Read the notepad file with the height data extracted with AWS Textract

#First Page



Field.Data.1<-read.csv('C:\\Users\\frm10\\Downloads\\output1.csv', header=T, skip=2);
str(Field.Data.1)
dim(Field.Data.1)
fix(Field.Data.1)

writeData(FieldDataWkb, sheet= "Page 1",Field.Data.1 ) ;

saveWorkbook(FieldDataWkb,file= "Field.Data.xlsx", overwrite = T) ;

#Second Page
addWorksheet(wb=FieldDataWkb, sheetName = "Page 2")

Field.Data.2<-read.csv('C:\\Users\\frm10\\Downloads\\output2.csv', header=T, skip=2);
str(Field.Data.2)
dim(Field.Data.2)
fix(Field.Data.2)

writeData(FieldDataWkb, sheet= "Page 2",Field.Data.2) ;

saveWorkbook(FieldDataWkb,file= "Field.Data.xlsx", overwrite = T) ;

#Third Page
addWorksheet(wb=FieldDataWkb, sheetName = "Page 3")

Field.Data.3<-read.csv('C:\\Users\\frm10\\Downloads\\output3.csv', header=T, skip=2);
str(Field.Data.3)
dim(Field.Data.3)
fix(Field.Data.3)

writeData(FieldDataWkb, sheet= "Page 3",Field.Data.3) ;


#Fourth Page
addWorksheet(wb=FieldDataWkb, sheetName = "Page 4")

Field.Data.4<-read.csv('C:\\Users\\frm10\\Downloads\\output4.csv', header=T, skip=2);
str(Field.Data.4)
dim(Field.Data.4)
fix(Field.Data.4)

writeData(FieldDataWkb, sheet= "Page 4",Field.Data.4) ;

#Fifth Page
addWorksheet(wb=FieldDataWkb, sheetName = "Page 5")

Field.Data.5<-read.csv('C:\\Users\\frm10\\Downloads\\output5.csv', header=T, skip=2);
str(Field.Data.5)
dim(Field.Data.5)
fix(Field.Data.5)

writeData(FieldDataWkb, sheet= "Page 5",Field.Data.5) ;

#Sixth Page
addWorksheet(wb=FieldDataWkb, sheetName = "Page 6")

Field.Data.6<-read.csv('C:\\Users\\frm10\\Downloads\\output6.csv', header=T, skip=2);
str(Field.Data.6)
dim(Field.Data.6)
fix(Field.Data.6)

writeData(FieldDataWkb, sheet= "Page 6",Field.Data.6) ;


saveWorkbook(FieldDataWkb,file= "Field.Data.xlsx", overwrite = T) ;




