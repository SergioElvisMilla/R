
#  na.string nos llena los datos completando espacios en blanco con NA (para variables categoricas)
data <- read.csv("PU-SA/data/missing-data.csv", na.strings = "")


#
##################################################################################################################
#na.omit   omite los campos vacios, obviando esos registros
data.cleaned<- na.omit(data)

# para revisar si el dato es vacio
is.na(data[4,2])
is.na(data[4,1])
#por columna
is.na(data$Income)

#########################limpieza selectiva##################################################################################

#LIMPIAR NA de solamente la variable Income
data.income.cleaned <- data[!is.na(data$Income),]

#filas comletas para un dataframe

complete.cases(data)


data.cleaned.2<- data[complete.cases(data),]
# convertir los ceros de ingresos en NA


data$Income[data$Income==0] <- NA




#medidas de centralizacion y dispercion

#promediando la data (columna income) "si hay valores NA no promediara nada"
mean(data$Income)
# promendiando y obviando los datos vacios  (na.rm los omite)
mean(data$Income, na.rm = TRUE)



#desviiacion tipica

sd(data$Income)
sd(data$Income, na.rm = TRUE)




