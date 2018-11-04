install.packages("jsonlite")
library(jsonlite)




dat.1 <-fromJSON("PU-SA/data/students.json")
dat.2 <-fromJSON("PU-SA/data/student-courses.json")

# toJson -> escribe un dato en un json
toJSON()

# link para entrar al json de tipo de cambio de yahoo en tiempo real   ---->  https://finance.yahoo.com/webservice/v1/symbols/allcurrencies/quote?format=json


url<- "https://finance.yahoo.com/webservice/v1/symbols/allcurrencies/quote?format=json"
currencies<- fromJSON(url)

#el simbolo dolar para ingresar a las capas y extraer datos

currency.data<- currencies$list$resources$resource$fields

head(dat.1, 3)
dat.1$Email

currency.data[1:5,1:2]
dat.1[c(2,5,8),]
dat.1[,c(2,5)]
head(dat.2,3)
