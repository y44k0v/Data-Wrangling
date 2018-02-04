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
	if(as.numeric(nchar(card_balance[i,1]))!=6){
		card_balance_id_wrong<-c(card_balance_id_wrong,as.numeric(card_balance[i,1]))
	}
}
#Credit Card Bank A
credit_card_bank_a_id_wrong<-c()
for(i in 1:nrow(credit_card_bank_a)){
	if(as.numeric(nchar(credit_card_bank_a[i,1]))!=6){
		credit_card_bank_a_id_wrong<-c(credit_card_bank_a_id_wrong,as.numeric(credit_card_bank_a[i,1]))
	}
}
#Credit Card Bank B
credit_card_bank_b_id_wrong<-c()
for(i in 1:nrow(credit_card_bank_b)){
	if(as.numeric(nchar(credit_card_bank_b[i,1]))!=6){
		credit_card_bank_b_id_wrong<-c(credit_card_bank_b_id_wrong,as.numeric(credit_card_bank_b[i,1]))
	}
}
#Mortgage Credit
mortgage_credit_id_wrong<-c()
for(i in 1:nrow(mortgage_credit)){
	if(as.numeric(nchar(mortgage_credit[i,1]))!=6){
		mortgage_credit_id_wrong<-c(mortgage_credit_id_wrong,as.numeric(mortgage_credit[i,1]))
	}
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
		if(is.na(id_wrong[i,j])==T){
			id_wrong[i,j]<-''
		}
	}
}
colnames(id_wrong)<-c('Card Balance','Credit Card Bank A','Credit Card Bank B','Mortgage Credit')


#elimination of registers with id wrong
#card balance
card_balance_temp<-cbind(c(1:nrow(card_balance)),card_balance)
temp2<-0
for(i in 1:length(card_balance_id_wrong)){
	if(is.na(card_balance_id_wrong[i])==F){
		temp2<-as.numeric(card_balance_temp[card_balance_temp[2]==card_balance_id_wrong[i]][1])
		card_balance_temp<-card_balance_temp[-temp2,]
		card_balance_temp<-card_balance_temp[-1]
		card_balance_temp<-cbind(c(1:nrow(card_balance_temp)),card_balance_temp)
	}
}
card_balance_temp<-card_balance_temp[-1]
#credit card bank A
credit_card_bank_a_temp<-cbind(c(1:nrow(credit_card_bank_a)),credit_card_bank_a)
temp2<-0
for(i in 1:length(credit_card_bank_a_id_wrong)){
	if(is.na(credit_card_bank_a_id_wrong[i])==F){
		temp2<-as.numeric(credit_card_bank_a_temp[credit_card_bank_a_temp[2]==credit_card_bank_a_id_wrong[i]][1])
		credit_card_bank_a_temp<-credit_card_bank_a_temp[-temp2,]
		credit_card_bank_a_temp<-credit_card_bank_a_temp[-1]
		credit_card_bank_a_temp<-cbind(c(1:nrow(credit_card_bank_a_temp)),credit_card_bank_a_temp)
	}	
}
credit_card_bank_a_temp<-credit_card_bank_a_temp[-1]
#credit card bank B
credit_card_bank_b_temp<-cbind(c(1:nrow(credit_card_bank_b)),credit_card_bank_b)
temp2<-0
for(i in 1:length(credit_card_bank_b_id_wrong)){
	if(is.na(credit_card_bank_b_id_wrong[i])==F){
		temp2<-as.numeric(credit_card_bank_b_temp[credit_card_bank_b_temp[2]==credit_card_bank_b_id_wrong[i]][1])
		credit_card_bank_b_temp<-credit_card_bank_b_temp[-temp2,]
		credit_card_bank_b_temp<-credit_card_bank_b_temp[-1]
		credit_card_bank_b_temp<-cbind(c(1:nrow(credit_card_bank_b_temp)),credit_card_bank_b_temp)
	}
}
credit_card_bank_b_temp<-credit_card_bank_b_temp[-1]
#mortgage credit
mortgage_credit_temp<-cbind(c(1:nrow(mortgage_credit)),mortgage_credit)
temp2<-0
for(i in 1:length(mortgage_credit_id_wrong)){
	if(is.na(mortgage_credit_id_wrong[i])==F){
		temp2<-as.numeric(mortgage_credit_temp[mortgage_credit_temp[2]==mortgage_credit_id_wrong[i]][1])
		mortgage_credit_temp<-mortgage_credit_temp[-temp2,]
		mortgage_credit_temp<-mortgage_credit_temp[-1]
		mortgage_credit_temp<-cbind(c(1:nrow(mortgage_credit_temp)),mortgage_credit_temp)
	}
}
mortgage_credit_temp<-mortgage_credit_temp[-1]


#inconsistencies with the genre
genre_wrong<-c()
for(i in 1:nrow(card_balance)){
	if(is.na(as.character(mortgage_credit[mortgage_credit[1]==as.numeric(card_balance[i,1]),2][1,1]))==F & as.character(mortgage_credit[mortgage_credit[1]==as.numeric(card_balance[i,1]),2][1,1])!=card_balance[i,2]){
		genre_wrong<-c(genre_wrong,card_balance[i,1])
	}
}
genre_wrong<-as.data.frame(as.numeric(genre_wrong))
colnames(genre_wrong)<-c('ID')

#empty registers and replacement for mean	
#card balance 
card_balance_empty_reg<-c(0,0)
temp1<-0
temp2<-c()
temp3<-c()
temp4<-c()	
for(i in 1:nrow(card_balance_temp)){
	for(j in 3:ncol(card_balance_temp)){
		if(is.na(card_balance_temp[i,j])==T){
			temp1<-temp1+1
		}
		if(is.na(card_balance_temp)[i,j]==F){
			temp3<-c(temp3,as.numeric(card_balance_temp[i,j]))
		}
	}
	if(temp1>5){
		temp4<-c(temp4,as.numeric(card_balance_temp[i,1]))
	}
	if(temp1==(ncol(card_balance_temp)-2)){
		temp2<-c(card_balance_temp[i,1],'No exist')
	}
	if(temp1<(ncol(card_balance_temp)-2) & temp1>0){
		card_balance_empty_reg<-rbind(card_balance_empty_reg,c(card_balance_temp[i,1],temp1))
		for(j in 3:(ncol(card_balance_temp))){
			if(is.na(card_balance_temp[i,j])==T){
				card_balance_temp[i,j]<-as.numeric(mean(temp3))
			}
		}
	}
	temp1<-0
	temp2<-c()
	temp3<-c()
}
card_balance_empty_reg<-card_balance_empty_reg[-1,]
card_balance_temp<-card_balance_temp[-temp4,]
colnames(card_balance_empty_reg)<-c('ID','Number of Missing Values')


#suspect card balance
suspect_card_balance<-c()
index_suspect<-0
temp3<-F
temp4<-c()
temp5<-0
temp6<-0
temp7<-c()
for(i in 1:nrow(card_balance_temp)){
	for(j in 3:(ncol(card_balance_temp)-1)){
			index_suspect<-card_balance_temp[i,j+1]/card_balance_temp[i,j]
		if(index_suspect>3){
			if(temp3==F){
				temp4<-c(card_balance_temp[i,1],paste('Month',as.character(j-1)))
				temp3<-T
			}
				else{
				temp4<-c(temp4,paste('Month',as.character(j-1)))
				}
		}		
	}
	if(length(temp4)!=0){
		temp6<-temp6+1
		if(length(temp4)<temp5){
			for(i in 1:(temp5-length(temp4))){
				temp4<-c(temp4,'')
				suspect_card_balance<-rbind(suspect_card_balance,temp4)
			}
		}
		if(length(temp4)>temp5){
			for(i in 1:(length(temp4)-temp5)){
				for(i in (temp6)){
					temp7<-c(temp7,'')
				}
				suspect_card_balance<-cbind(suspect_card_balance,temp7)
			}
			temp7<-c()
			suspect_card_balance<-rbind(suspect_card_balance,temp4)
		}
		if(length(temp4)==temp5){
					suspect_card_balance<-rbind(suspect_card_balance,temp4)
		}
	}
	temp5<-max(length(temp4),temp5)
	temp3<-F
	temp4<-c()
}
suspect_card_balance<-suspect_card_balance[-1,]
temp8<-duplicated(suspect_card_balance)
temp9<-cbind(c(1:nrow(suspect_card_balance)),suspect_card_balance)
temp9<-as.data.frame(temp9,row.names=T)
temp10<-0
for(i in 1:length(temp8)){
	if (temp8[i]==T){
		temp10<-as.numeric(temp9[temp9[1]==i,1])
		temp9<-temp9[-temp10,]
	}
}
suspect_card_balance<-temp9[-1]
colnames(suspect_card_balance)<-c('ID',rep('Suspect Month',times=ncol(suspect_card_balance)-1))


#merge credit balance
credit_merge<-credit_card_bank_a_temp
temp1<-cbind(c(1:nrow(credit_merge)),credit_merge[1])
temp2<-0
for(i in 1:nrow(credit_card_bank_b_temp)){
	if(length(credit_card_bank_a_temp[credit_card_bank_a_temp[1]==credit_card_bank_b_temp[i,1]])==0){
		credit_merge<-rbind(credit_merge,as.numeric(credit_card_bank_b[i,]))
	}
	if(length(credit_card_bank_a_temp[credit_card_bank_a_temp[1]==credit_card_bank_b_temp[i,1]])>0){
		for(j in 2:ncol(credit_card_bank_a_temp)){
			temp2<-temp1[temp1[2]==credit_card_bank_b_temp[i,1]][1]
			credit_merge[temp2,j]<-sum(credit_card_bank_a[temp2,j],credit_card_bank_b_temp[i,j])
		}
	}
}


#generation files
fileout1<-loadWorkbook('Problems.xlsx',create=T)
fileout2<-loadWorkbook('Data Wrangled.xlsx',create=T)
#card balance
createSheet(fileout2,name='CardBalance')
createName(fileout2,name='CardBalance',formula='CardBalance!$A$1')
writeNamedRegion(fileout2,card_balance_temp,name='CardBalance')
#total credit
createSheet(fileout2,name='TotalCredit')
createName(fileout2,name='TotalCredit',formula='TotalCredit!$A$1')
writeNamedRegion(fileout2,credit_merge,name='TotalCredit')
#mortgage credit
createSheet(fileout2,name='MortgageCredit')
createName(fileout2,name='MortgageCredit',formula='MortgageCredit!$A$1')
writeNamedRegion(fileout2,mortgage_credit_temp,name='MortgageCredit')
#ID wrong
createSheet(fileout1,name='ID')
createName(fileout1,name='ID',formula='ID!$A$1')
writeNamedRegion(fileout1,id_wrong,name='ID')
#Genre Wrong
createSheet(fileout1,name='Genre')
createName(fileout1,name='Genre',formula='Genre!$A$1')
writeNamedRegion(fileout1,genre_wrong,name='Genre')
#empty registers
createSheet(fileout1,name='EmptyRegisters')
createName(fileout1,name='EmptyRegisters',formula='EmptyRegisters!$A$1')
writeNamedRegion(fileout1,card_balance_empty_reg,name='EmptyRegisters')
#suspect card balance
createSheet(fileout1,name='SuspectCardBalance')
createName(fileout1,name='SuspectCardBalance',formula='SuspectCardBalance!$A$1')
writeNamedRegion(fileout1,suspect_card_balance,name='SuspectCardBalance')
#Save Files
saveWorkbook(fileout1)
saveWorkbook(fileout2)

