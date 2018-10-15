data <- read.csv("PU-SA/data/missing-data.csv",na.strings = "")


###########################################################################################################################
                               # REEMPLAZANDO COM MEDIA #



# reemplazando por el promedio  (.mean) de la poblacion , con data$Income.mean estoy crando una nueva columna al data frame "data" y se le asignan los valores 
# resultantes de nuestro ifelse
                      #### si sucede esto###   ########## hacer esto #######   ##sino, esto## 
data$Income.mean <- ifelse(is.na(data$Income), mean(data$Income, na.rm = TRUE), data$Income)



###########################################################################################################################
                             # REEMPLAZANDO ALEATORIAMENTE #
# CARGANDO DATA LIMPIA
data <- read.csv("PU-SA/data/missing-data.csv",na.strings = "")
data$Income[data$Income==0]<-NA

# PARA VARIABLE NUMERICA y categorica

# x es un vector de datos que puede contener NA
rand.impute <- function(x){
# missing contiene un vector de valor true o false dependiendo de NA de X
 missing<- is.na(x)
 # la variable n.missing contiene cuantos valores son NA dentro de X
 n.missing <- sum(missing)
 #x.obs son los valores conocidos que tinen datos diferentes de NA en X
 x.obs<- x[!missing]
 # por defecto devolvere lo que habia entrado por parametro
 imputed <- x
 # en los valore que faltaban los reemplazamos por una muestra que si conocemos
 imputed[missing] <- sample(x.obs, n.missing, replace = TRUE)
 return (imputed)
   
}


random.impute.data.frame<- function(dataframe, cols){
  names<- names(dataframe)
  for (col in cols) {
    # añadir una nueva columna con el nombre de la columna original + imputed deparado por un punto "."
    name<- paste(names[col],"imputed", sep = ".")
    # asignado el valor a la nueva columna, el valor sera el obtenido en la funcion "rand.impute"
    dataframe[name]= rand.impute(dataframe[,col])
  }
  dataframe
  
}
 
# llamada a la funcion, Y mis parametros son la columna 1 y 2
data<-random.impute.data.frame(data,c(1,2))



names(data)







