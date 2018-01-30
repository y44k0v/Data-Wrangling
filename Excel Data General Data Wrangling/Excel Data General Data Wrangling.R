#packages
install.packages('readxl', repos='http://cran.us.r-project.org')
install.packages('XLConnect', repos='http://cran.us.r-project.org')
library('readxl')
library('XLConnect')


#selection of file
filein<-file.choose()


#access to the data row
card_balance<-read_xlsx(filein,sheet='Card Balance')
credit_card_bank_a<-read_xlsx(filein,sheet='Credit Card Bank A')
credit_card_bank_b<-read_xlsx(filein,sheet='Credit Card Bank B')
mortgage_credit<-read_xlsx(filein,sheet='Mortgage Credit')


#generation dataframe of wrong ID 
#Card Balance
card_balance_id_wrong<-c()
for(i in 1:nrow(card_balance)){
	if(as.numeric(nchar(card_balance[i,1]))!=6)
	card_balance_id_wrong<-c(card_balance_id_wrong,as.numeric(card_balance[i,1]))
}
#Credit Card Bank A
credit_card_bank_a_id_wrong<-c()
for(i in 1:nrow(credit_card_bank_a)){
	if(as.numeric(nchar(credit_card_bank_a[i,1]))!=6)
	credit_card_bank_a_id_wrong<-c(credit_card_bank_a_id_wrong,as.numeric(credit_card_bank_a[i,1]))
}
#Credit Card Bank B
credit_card_bank_b_id_wrong<-c()
for(i in 1:nrow(credit_card_bank_b)){
	if(as.numeric(nchar(credit_card_bank_b[i,1]))!=6)
	credit_card_bank_b_id_wrong<-c(credit_card_bank_b_id_wrong,as.numeric(credit_card_bank_b[i,1]))
}
#Mortgage Credit
mortgage_credit_id_wrong<-c()
for(i in 1:nrow(mortgage_credit)){
	if(as.numeric(nchar(mortgage_credit[i,1]))!=6)
	mortgage_credit_id_wrong<-c(mortgage_credit_id_wrong,as.numeric(mortgage_credit[i,1]))
}
#generation table
temp1<-max(length(card_balance_id_wrong),length(credit_card_bank_a_id_wrong),length(credit_card_bank_b_id_wrong),length(mortgage_credit_id_wrong))
length(card_balance_id_wrong)<-temp1
length(credit_card_bank_a_id_wrong)<-temp1
length(credit_card_bank_b_id_wrong)<-temp1
length(mortgage_credit_id_wrong)<-temp1
id_wrong<-cbind(card_balance_id_wrong,credit_card_bank_a_id_wrong,credit_card_bank_b_id_wrong,mortgage_credit_id_wrong)
for(i in 1:nrow(id_wrong)){
	for(j in 1:ncol(id_wrong)){
	if(is.na(id_wrong[i,j])==T)
		id_wrong[i,j]<-''
}}
colnames(id_wrong)<-c('Card Balance','Credit Card Bank A','Credit Card Bank B','Mortgage Credit')


#elimination of registers with id wrong
#card balance
card_balance_temp<-cbind(c(1:nrow(card_balance)),card_balance)
temp2<-0
for(i in 1:length(card_balance_id_wrong)){
	if(is.na(card_balance_id_wrong[i])==F)
	temp2<-as.numeric(card_balance_temp[card_balance_temp[2]==card_balance_id_wrong[i]][1])
	card_balance_temp<-card_balance_temp[-temp2,]
}
#credit card bank A
credit_card_bank_a_temp<-cbind(c(1:nrow(credit_card_bank_a)),credit_card_bank_a)
temp2<-0
for(i in 1:length(credit_card_bank_a_id_wrong)){
	if(is.na(credit_card_bank_a_id_wrong[i])==F)
	temp2<-as.numeric(credit_card_bank_a_temp[credit_card_bank_a_temp[2]==credit_card_bank_a_id_wrong[i]][1])
	credit_card_bank_a_temp<-credit_card_bank_a_temp[-temp2,]
}
#credit card bank B
credit_card_bank_b_temp<-cbind(c(1:nrow(credit_card_bank_b)),credit_card_bank_b)
temp2<-0
for(i in 1:length(credit_card_bank_b_id_wrong)){
	if(is.na(credit_card_bank_b_id_wrong[i])==F)
	temp2<-as.numeric(credit_card_bank_b_temp[credit_card_bank_b_temp[2]==credit_card_bank_b_id_wrong[i]][1])
	credit_card_bank_b_temp<-credit_card_bank_b_temp[-temp2,]
}
#mortgage credit
mortgage_credit_temp<-cbind(c(1:nrow(mortgage_credit)),mortgage_credit)
temp2<-0
for(i in 1:length(mortgage_credit_id_wrong)){
	if(is.na(mortgage_credit_id_wrong[i])==F)
	temp2<-as.numeric(mortgage_credit_temp[mortgage_credit_temp[2]==mortgage_credit_id_wrong[i]][1])
	mortgage_credit_temp<-mortgage_credit_temp[-temp2,]
}


#inconsistencies with the genre
genre_wrong<-c()
for(i in 1:nrow(card_balance)){
	if(is.na(as.character(mortgage_credit[mortgage_credit[1]==as.numeric(card_balance[i,1]),2][1,1]))==F & as.character(mortgage_credit[mortgage_credit[1]==as.numeric(card_balance[i,1]),2][1,1])!=card_balance[i,2])
	genre_wrong<-c(genre_wrong,card_balance[i,1])
}
genre_wrong<-as.data.frame(as.numeric(genre_wrong))
colnames(genre_wrong)<-c('ID')


#generation files
fileout1<-loadWorkbook('Problems.xlsx',create=T)
fileout2<-loadWorkbook('Data Wrangled.xlsx',create=T)
#card balance
createSheet(fileout2,name='CardBalance')
createName(fileout2,name='CardBalance',formula='CardBalance!$A$1')
writeNamedRegion(fileout2,card_balance_temp,name='CardBalance')
#ID wrong
createSheet(fileout1,name='ID')
createName(fileout1,name='ID',formula='ID!$A$1')
writeNamedRegion(fileout1,id_wrong,name='ID')
#Genre Wrong
createSheet(fileout1,name='Genre')
createName(fileout1,name='Genre',formula='Genre!$A$1')
writeNamedRegion(fileout1,genre_wrong,name='Genre')
#Save Files
saveWorkbook(fileout1)
saveWorkbook(fileout2)

