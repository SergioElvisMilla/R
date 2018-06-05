#_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_XML_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-

# instalando paquetes
install.packages("XML")
#usando la libreria
library(XML)
#cargando un XML
url<- "PU-SA/data/cd_catalog.xml"

#APUNTANDO EL XML
xmlDoc<- xmlParse(url) #NOS DEVUELVE UN XMLInternalDocument
rootnode <- xmlRoot(xmlDoc)
rootnode[1]

# visualizar el xml como tabla
cds_data <- xmlSApply(rootnode,function(x) xmlSApply(x,xmlValue) )
#transponer filas por columnas
cds.catalog <- data.frame(t(cds_data),row.names = NULL)
head(cds.catalog,2)
cds.catalog[1:5,]

#xpathSApply()
#getNodeSet()


population_url <- "PU-SA/data/WorldPopulation-wiki.htm"
tables<- readHTMLTable(population_url)


