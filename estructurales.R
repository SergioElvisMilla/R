#cargar archivo
bajop=read.table(file.choose(),header=T,sep="\t")
bajop
colnames(bajop)

#tabla de frecuencia
bajop$fuma
table(bajop$fuma)
prop.table(table(bajop$fuma))

#promedio
colnames(bajop)
mean(bajop$edad)
mean(bajop$pesomama)

#mediana
median(bajop$edad)

hist(bajop$edad)

#mediana peso mama
mean(bajop$pesomama)
median(bajop$pesomama)
hist(bajop$pesomama)

#medidas de posicion 
#dividir los datos en 4 partes iguales -----> cuartiles
#dividir los datos en 10 partes iguales -----> deciles
#dividir los datos en 2 partes iguales -----> mediana
#dividir los datos en 100 partes iguales -----> percentiles

#percentil 10 del peso de la medre
quantile(bajop$pesomama,0.1)

#todos los deciles
quantile(bajop$pesomama,c(0.1,0.2,0.3,0.4,0.5,0.6,0.7,0.8,0.9))

#todos los cuartiles
quantile(bajop$pesomama,c(0.25,0.5,0.75))

#medidas de dispersion

#rango
range(bajop$pesomama)

#rango intercuartilico: Q3-Q1
quantile(bajop$pesomama,0.75)-quantile(bajop$pesomama,0.25)

#varianza
var(bajop$pesomama)

#desviacion estandar
sd(bajop$pesomama)
sqrt(var(bajop$pesomama))

#coeficiente de asimetria
hist(bajop$pesomama)
mean(bajop$pesomama)
median(bajop$pesomama)

install.packages("moments")
library(moments)
skewness(bajop$pesomama)

#AS=1.39 ----->cola es a la derecha, la media es mayor a la mediana
#AS<0 -----> cola esa la izquierda, la media es menor a la mediana
#AS= 0 ----->distribucion simetrica, media es igual a la mediana
