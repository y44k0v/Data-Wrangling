#packages
install.packages('xml2', repos='http://cran.us.r-project.org')
install.packages('rvest', repos='http://cran.us.r-project.org')
install.packages('XLConnect', repos='http://cran.us.r-project.org')
library('xml2')
library('rvest')
library('XLConnect')
library('stringr')

#extraction of the data from the web
url<-'https://www.bvc.com.co/pps/tibco/portalbvc/Home/Empresas/Listado+de+Emisores?action=dummy'
web<-read_html(url)
nod<-html_nodes(web,xpath='//*[@id="texto_28"]')
tab<-html_table(nod[1])
tab<-as.data.frame(tab)
tab<-tab[2:length(tab)]



#orthography corrections
for(i in 1:nrow(tab)){
tab[i,4]<-str_replace_all(tab[i,4],'BogotÃ¡','Bogota')
tab[i,4]<-str_replace_all(tab[i,4],'MedellÃ­n','Medellin')
tab[i,4]<-str_replace_all(tab[i,4],'ItaguÃ­','Itagui')
tab[i,4]<-str_replace_all(tab[i,4],'TuluÃ¡','Tulua')
}


#generation of file
file<-loadWorkbook('BVC.xlsx',create=T)
createSheet(file,name='Indices')
createName(file,name='Indices',formula='Indices!$A$1')
writeNamedRegion(file,tab,name='Indices')
saveWorkbook(file)

