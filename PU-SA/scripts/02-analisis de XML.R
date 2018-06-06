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

#APUNTANDO EL XML
population_url <- "PU-SA/data/WorldPopulation-wiki.htm"

 # devuelve una lista de todas las tablas en una pagina
tables<- readHTMLTable(population_url)

mas_populares <- tables[[1]]
head(mas_populares,2)


# especificando una tabla, para no cargar todo
custom_table <- readHTMLTable(population_url,which = 6)

custom_table[3:5,]

