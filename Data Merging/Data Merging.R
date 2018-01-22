#packages
install.packages('readxl', repos='http://cran.us.r-project.org')
install.packages('XLConnect', repos='http://cran.us.r-project.org')
library('readxl')
library('XLConnect')


#selection of file
filein<-file.choose()


#access to the data row
year1<-read_xlsx(filein,sheet='2014')
year2<-a<-read_xlsx(filein,sheet='2015')
year3<-read_xlsx(filein,sheet='2016')



#merge of 2014 and 2015
tab<-year1
tab<-cbind(tab,0)
year1<-cbind(c(1:nrow(year1)),year1)


for(i in 1:nrow(year2)){
if(length(year1[year1[2]==as.character(year2[i,1])])==0)
tab<-rbind(tab,c(as.character(year2[i,1]),as.character(year2[i,2]),0,as.numeric(year2[i,3])))
else
tab[year1[year1[2]==as.character(year2[i,1]),1],4]<-year2[i,3]
}


#merge of 2016
year1<-tab
tab<-cbind(tab,0)
year1<-cbind(c(1:nrow(year1)),year1)


for(i in 1:nrow(year3)){
if(length(year1[year1[2]==as.character(year3[i,1])])==0)
tab<-rbind(tab,c(as.character(year3[i,1]),as.character(year3[i,2]),0,0,as.numeric(year3[i,3])))
else
tab[year1[year1[2]==as.character(year3[i,1]),1],5]<-year3[i,3]
}


#rename of col
colnames(tab)<-c('NIT','Razon Social','2014','2015','2016')



#generation of output file
fileout<-loadWorkbook('Data Merging.xlsx',create=T)
createSheet(fileout,name='FOB')
createName(fileout,name='FOB',formula='FOB!$A$1')
writeNamedRegion(fileout,tab,name='FOB')
saveWorkbook(fileout)

