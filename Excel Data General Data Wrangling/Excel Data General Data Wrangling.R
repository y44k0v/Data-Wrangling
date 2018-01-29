
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



#generation file
fileout<-loadWorkbook('Problems.xlsx',create=T)
createSheet(fileout,name='ID')
createName(fileout,name='ID',formula='ID!$A$1')
writeNamedRegion(fileout,id_wrong,name='ID')
saveWorkbook(fileout)

