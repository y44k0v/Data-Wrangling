#packages
install.packages('readxl', repos='http://cran.us.r-project.org')
install.packages('XLConnect', repos='http://cran.us.r-project.org')
library('readxl')
library('XLConnect')



#selection of file
filein<-file.choose()


#access to the data row
dataraw<-read_xlsx(filein,sheet='Censo',col_names=T)


#construction of name and code
tempcode<-c()
for(i in 1:nrow(dataraw)){
tempcode<-c(tempcode,as.numeric(dataraw[i,1]))
}
codecom<-rep(tempcode,each=22)

tempname<-c()
for(i in 1:nrow(dataraw)){
tempname<-c(tempname,as.character(dataraw[i,2]))
}
namecom<-rep(tempname,each=22)


#construccion dataframe
tab<-cbind(codecom,namecom)
ldat<-c()
for(i in 1:nrow(dataraw)){
for(j in 1:22){
ldat<-c(ldat,as.numeric(dataraw[i,j+2]))
}}
tab<-cbind(tab,ldat)
colnames(tab)<-c('Codigo Comuna','Nombre Coomuna','Total Censo')


#generation of output file
fileout<-loadWorkbook('Tidy Data.xlsx',create=T)
createSheet(fileout,name='Censo')
createName(fileout,name='Censo',formula='Censo!$A$1')
writeNamedRegion(fileout,tab,name='Censo')
saveWorkbook(fileout)






