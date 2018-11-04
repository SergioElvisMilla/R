
# subir ficheros con ancho fijo
students_data <- read.fwf("PU-SA/data/student-fwf.txt",
                          widths =c(4,15,20,15,4),
                          col.names = c("id","nombre","email","carrera","año")
                          )

# skip me permite saltame a la siguiente linea 
students_data_header<- read.fwf("PU-SA/data/student-fwf-header.txt",
                                widths =c(4,15,20,15,4),
                                header = TRUE,
                                sep = "\t",
                                skip = 2
                                )
# -20, al poner el ancho de la columna en negativo, estamos cargando el fichero sin esa columna

students_data_no_email <- read.fwf("PU-SA/data/student-fwf.txt",
                          widths =c(4,15,-20,15,4),
                          col.names = c("id","nombre","carrera","año")
)