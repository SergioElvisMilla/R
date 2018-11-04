#trabajando con datos rdata -- creando un data frame

#as.Date() para convertir mi cadena de string a fecha

clientes <- c("Juan Gabriel", "Ricardo","Pedro")
fecha<- as.Date(c("2017-12-27","2017-11-1","2017-1-1"))
pago <- c(315,192.55,40.15)
pedidos <- data.frame(clientes, fecha, pago)
################################################################################
# guardar el data frame

clientes_vip <-c("Juan Gabril")

save(pedidos, file = "PU-SA/data/pedidos.Rdata")
saveRDS(pedidos, file = "PU-SA/data/pedidos2.rds")

#para quitarlo de Global Environment
remove(pedidos)


######################################################################################
# cargar un fichero RDATA

load("PU-SA/data/pedidos.Rdata")

orders <-readRDS("PU-SA/data/pedidos.rds")
 

# cargando data sets
#iris= mide la anchura  de los petalos , sepalos de tre especies de flores stosa versicolor y virginica
data(iris)
names(iris)

data(cars)


# guarda todo en un rdata

save.image(file ="PU-SA/data/alldata.Rdata")

#

primes<- c(2,3,5,7,11,13)
pow2 <-c(2,4,8,16,32,64,128)

# list especifica un vector que queremos guardar en formato string (guara en este ejemplo dos objetos = primes y pow2)
save(list = c("primes","pow2"), file = c("PU-SA/data/primes_and_pow2.Rdata"))


#advertencia


attach("PU-SA/data/primes_and_pow2.Rdata")


data()


