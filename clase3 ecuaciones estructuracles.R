data(iris)
datos<- iris
View(datos)
#  seet.seed, obtiene datos aleatorios desde un punto especifico (101)
set.seed(101)
tamano.total <- nrow(datos)
tamano.entreno <- round(tamano.total*0.7)
datos.indices <- sample(1:tamano.total , size=tamano.entreno)
datos.entreno <- datos[datos.indices,]
View(datos.entreno)
datos.test <- datos[-datos.indices,]
View(datos.test) 


# Ejecución del modelo de clasificación C5.0  [antes de ejecutar, se debe marcar con check el paquete c50]
modelo <- C5.0(Species ~ .,data = datos.entreno)
summary(modelo) # Información sobre el modelo

plot(modelo) # Gráfico


# Para detallar un nodo en particular se usaria la siguiente función
plot(modelo, subtree=3)  #Muestra un nodo en particular

# predicción
prediccion <- predict(modelo,newdata=datos.test)
View(prediccion)
# Matriz de confusión
tabla <- table(prediccion, datos.test$Species)
tabla


# % correctamente clasificados
100*sum(diag(tabla))/sum(tabla)



# Sepal no interviene, por lo tanto no necesita un valor
nuevo <- data.frame(Sepal.Length=NA,Sepal.Width=NA,Petal.Length=5,Petal.Width=1)
View(nuevo)
prediccion<-predict(modelo,nuevo)
View(prediccion)


prediccion.prob<-predict(modelo,nuevo, type = "prob")
View(prediccion.prob)
predict(modelo,nuevo)

###################################################################################
datos<- read.delim("clipboard")
View(datos)

# modelo C5
modelo<-C5.0(PlayTennis~.,data=datos)
summary(modelo)
# grafico del arbol
plot(modelo)

datos.test<-read.delim('clipboard')
View(datos.test)


prediccion <- predict(modelo,datos.test)
View(prediccion)


# Matriz de confusión
tabla <- table(prediccion, datos.test$PlayTennis)
tabla
# % correctamente clasificados
100 * sum(diag(tabla)) / sum(tabla)





