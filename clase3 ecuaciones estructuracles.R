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


# Ejecuci�n del modelo de clasificaci�n C5.0  [antes de ejecutar, se debe marcar con check el paquete c50]
modelo <- C5.0(Species ~ .,data = datos.entreno)
summary(modelo) # Informaci�n sobre el modelo

plot(modelo) # Gr�fico


# Para detallar un nodo en particular se usaria la siguiente funci�n
plot(modelo, subtree=3)  #Muestra un nodo en particular

# predicci�n
prediccion <- predict(modelo,newdata=datos.test)
View(prediccion)
# Matriz de confusi�n
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


# Matriz de confusi�n
tabla <- table(prediccion, datos.test$PlayTennis)
tabla
# % correctamente clasificados
100 * sum(diag(tabla)) / sum(tabla)





