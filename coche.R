#carga archivo
bajop=read.table(file.choose(),header = T,sep = "\t")


#tabla de frecuencia



bajop$fuma
table(bajop$fuma)
prop.table(table(bajop$fuma))
