#packages
install.packages('readxl', repos='http://cran.us.r-project.org')
install.packages('XLConnect', repos='http://cran.us.r-project.org')
library('readxl')
library('XLConnect')


#selection of file
filein<-file.choose()


#access to the data row
dataraw<-read_xlsx(filein,sheet='COMUNAS',skip=24,col_names=F)


#construction of variables
temp1<-duplicated(dataraw[1:nrow(dataraw),7])
dataraw<-cbind(dataraw,temp1)
codecom<-dataraw[dataraw[12]==F,7]
namecom<-dataraw[dataraw[12]==F,6]
codecom<-codecom[-length(codecom)]
namecom<-namecom[-length(namecom)]
tab<-data.frame()
line<-c()
ages<-dataraw[1:22,8]


#generation of dataframe
for(i in 1:length(codecom)){
line<-c()
for(j in 1:22){
line<-c(line,dataraw[dataraw[7]==codecom[i]&dataraw[8]==ages[j],11][1])
}
tab<-rbind(tab,line)
}
tab<-cbind(codecom,namecom,tab)
colnames(tab)<-c('Nombre Comuna','Codigo Comuna',ages)



#generation of output file
fileout<-loadWorkbook('Tidy Data.xlsx',create=T)
createSheet(fileout,name='Censo')
createName(fileout,name='Censo',formula='Censo!$A$1')
writeNamedRegion(fileout,tab,name='Censo')
saveWorkbook(fileout)

