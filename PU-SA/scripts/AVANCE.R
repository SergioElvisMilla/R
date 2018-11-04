# 01-analisis-csv

#cargarndo un archivo a R
AUTO = read.csv("PU-SA/data/xzzxzxzxzx.csv", header = TRUE, sep = "," )   
names(AUTO)  #PARA REVISAR LAS CABECERAS

#para cargar archivos con separadores ";"
read.csv2== read.csv("PU-SA/data/xzzxzxzxzx.csv", sep =";", dec =",") # tambien puede ser sep="\t"    -----> para cargar archivos con separadores ";"



# cuando no hay cabeceras

auto_no_header<- read.csv("PU-SA/data/auto-mpg-noheader.csv", header = FALSE)  
head(auto_no_header,4)

# este ejemplo no se debe de hacer
auto_no_sense<- read.csv("PU-SA/data/auto-mpg-noheader.csv")

# colocar columnas (customizar)
auto_custom_header <- read.csv("PU-SA/data/auto-mpg-noheader.csv",
                               header = FALSE, 
                               col.names = c("numero","millas_por_galon","cilindrada",
                                             "desplazamineto","caballos_de_potencia",
                                             "peso","aceleracion","año","modelo"
                                             )
                               )
head(auto_custom_header,4)


# trabajar con valores especiales

#NA: Not Available
#na.strings=""
#as.character()
AUTO = read.csv("PU-SA/data/auto-mpg.csv", header = TRUE, sep = "," , strings.na = "", stringsAsFactors = FALSE )  



