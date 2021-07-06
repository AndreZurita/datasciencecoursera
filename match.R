rm(list = ls()) #limpia los objetos 
graphics.off() #limpiar el dispositivo de graficas
gc() #limpiar la memoria virtual 
cat("\014") #limpiar la consola

                #/// MATCH entre SF y CRM DYNAMICS \\\#

#Paso 1. Cruce de CI de postulantes con su informacion SF vs informacion en SF

library(tidyverse)

library(readxl)
ci_post <- read_excel("C:/Users/ritaz/Desktop/CONSULTOR CIERRE/Analítica/AnalisisVarios/MATCH CRM_SF/ci_post.xlsx")
ci_post <- as.data.frame(ci_post)
ci_post <- ci_post[,-(3:9)] 


av_carr <- read_excel("C:/Users/ritaz/Desktop/CONSULTOR CIERRE/Analítica/AnalisisVarios/MATCH CRM_SF/av_carr.xlsx")
av_carr <- as.data.frame(av_carr)


base_sf <- merge(ci_post,av_carr,by="nom")

table(base_sf$`Programa Pregrado: Nombre del plan del programa`)

library("writexl")
write_xlsx(base_sf,"C:/Users/ritaz/Desktop/CONSULTOR CIERRE/Analítica/AnalisisVarios/MATCH CRM_SF//base_sf.xlsx")


#Renombre de variables en SF para append futuro

base_sf$`Propietario de oportunidad` <- gsub("Paola Janeth Rodríguez Caicedo", "Paola Rodríguez Caicedo", base_sf$`Propietario de oportunidad` )
base_sf$`Propietario de oportunidad` <- gsub("Bryan Nicolas Andrango Quiroz", "Bryan Nicolas Andrango", base_sf$`Propietario de oportunidad` )
base_sf$`Propietario de oportunidad` <- gsub("Nicolas Arturo  Pazmiño Mesias", "Nicolas Arturo Pazmiño", base_sf$`Propietario de oportunidad` )
base_sf$`Propietario de oportunidad` <- gsub("María Belén Paredes", "Maria Belen Paredes", base_sf$`Propietario de oportunidad` )
base_sf$`Propietario de oportunidad` <- gsub("Yessenia Isabel Herrera Herrera", "Yessenia Herrera", base_sf$`Propietario de oportunidad` )
base_sf$`Propietario de oportunidad` <- gsub("Estefania Nicolle Narvaez Chediak", "Estefania Narváez", base_sf$`Propietario de oportunidad` )
base_sf$`Propietario de oportunidad` <- gsub("Gloria Ximena Ramos", "Gloria Ximena Ramos", base_sf$`Propietario de oportunidad` )
base_sf$`Propietario de oportunidad` <- gsub("Jorge Alejandro Tapia Acosta", "Jorge Alejandro Tapia", base_sf$`Propietario de oportunidad` )
base_sf$`Propietario de oportunidad` <- gsub("Juan Domingo Banderas Espinosa", "Juan Domingo Banderas", base_sf$`Propietario de oportunidad` )
base_sf$`Propietario de oportunidad` <- gsub("Brenda Nicole Jaramillo Miranda", "Brenda Nicole Jaramillo", base_sf$`Propietario de oportunidad` )
base_sf$`Propietario de oportunidad` <- gsub("Jessica Dayana Almeida Morillo", "Jessica Dayana Almeida", base_sf$`Propietario de oportunidad` )
base_sf$`Propietario de oportunidad` <- gsub("Patricia Camila Chiriboga", "Patricia Camila Chiriboga", base_sf$`Propietario de oportunidad` )
base_sf$`Propietario de oportunidad` <- gsub("Paola Alejandra Rueda Puga", "Paola Alejandra Rueda", base_sf$`Propietario de oportunidad` )
base_sf$`Propietario de oportunidad` <- gsub("Giuliana Andrea Varela Zambrano", "Giuliana Varela", base_sf$`Propietario de oportunidad` )
base_sf$`Propietario de oportunidad` <- gsub("André Vinicio Zurita Serrano", "André Zurita", base_sf$`Propietario de oportunidad` )


base_sf$`Programa Pregrado: Nombre del plan del programa` <- gsub("UDLA1P714-ENFERMERIA FLEX-202210", "UDLA1P714-ENFERMERIA-202210", base_sf$`Programa Pregrado: Nombre del plan del programa` )


#Separado de columnas de la carrera y su codigo correspon

library(stringr)
base_sf$`Programa Pregrado: Nombre del plan del programa`<-  str_split_fixed(base_sf$`Programa Pregrado: Nombre del plan del programa`, "-", 3)

colnames(base_sf)[7] <- "cod"

base_sf$cod <-  sub(".*UDLA1", "", base_sf$cod) #SOLO COD de carrera OK!
base_sf$cod <-  sub(".*UDLA1", "", base_sf$cod)
base_sf$cod <-  sub(".*UDLA1", "", base_sf$cod)
base_sf$cod <-  sub(".*UDLA1", "", base_sf$cod)
base_sf$cod <-  sub(".*UDLA2", "", base_sf$cod) 
base_sf$cod <-  sub(".*UDLA2", "", base_sf$cod)
base_sf$cod <-  sub(".*UDLA2", "", base_sf$cod)
base_sf$cod <-  sub(".*UDLA2", "", base_sf$cod)

base_sf$cod <- gsub(" ", "", base_sf$cod, fixed = TRUE)

base_sf$cod <- gsub("O", "000", base_sf$cod)

base_sf$cod <-  sub(".*000", "", base_sf$cod)

base_sf <- base_sf[complete.cases(base_sf$cod), ]



#Unir el codigo correspondiente con el nombre de la carrera siempre usado 

cod_carr <- read_excel("C:/Users/ritaz/Desktop/CONSULTOR CIERRE/Analítica/AnalisisVarios/MATCH CRM_SF/cod_carr.xlsx")

cod_carr[-c(37),] #ELIMINO ENFERMERIA ET PARA EVITAR DUPLICADOS

base_sf <- merge(base_sf,cod_carr,by="cod")


table(base_sf$carrera)


#Paso 2. Cruce de registros SF con CRM Dynamics
#Aqui se muestran solo las que tienen un correspondiente, es decir los migrados
#y su actualizacion de estado

pre_SF <- read_excel("C:/Users/ritaz/Desktop/CONSULTOR CIERRE/Analítica/AnalisisVarios/MATCH CRM_SF/pre_SF.xlsx")
pre_SF <- as.data.frame(pre_SF)
colnames(pre_SF)[19] <- "id"

pre_SF$new_carrera_acuerdoname <- gsub("UDLA1P714 - ENFERMERIA - MODALIDAD ESTUDIO TRABAJO", "UDLA1P714 - ENFERMERIA", pre_SF$new_carrera_acuerdoname)


library(stringr)

base_mig <- merge(base_sf,pre_SF,by="id") #Me interesa el que se llama etapa, dado que este es el cambio de estado

table(base_sf$carrera)
table(base_mig$carrera)

#Paso 3. Cruce de registros SF con CRM Dynamics, aquellos NO MIGRADOs
#Aqui se muestran solo las que no tienen un registro migrado, es decir 
#se quedaron por alguna razon en CRM Dymanics. 
#NO SE MIGRAN GANADO PERDIDOS Y ALGUNOS ???ABIERTOS??


base_nomig <- anti_join(pre_SF,base_sf,by="id")

colnames(base_nomig)[1] <- "nom"
colnames(base_nomig)[11] <- "Propietario de oportunidad"
colnames(base_nomig)[16] <- "cod"
colnames(base_nomig)[20] <- "Etapa"

base_nomig$cod <- gsub("UDLA1P714 - ENFERMERIA - MODALIDAD ESTUDIO TRABAJO", "UDLA1P714 - ENFERMERIA", base_nomig$cod)


base_nomig$cod <-  sub(".*UDLA1", "", base_nomig$cod) #SOLO COD de carrera OK!
base_nomig$cod <-  sub(".*UDLA1", "", base_nomig$cod)
base_nomig$cod <-  sub(".*UDLA1", "", base_nomig$cod)
base_nomig$cod <-  sub(".*UDLA1", "", base_nomig$cod)
base_nomig$cod <-  sub(".*UDLA2", "", base_nomig$cod) 
base_nomig$cod <-  sub(".*UDLA2", "", base_nomig$cod)
base_nomig$cod <-  sub(".*UDLA2", "", base_nomig$cod)
base_nomig$cod <-  sub(".*UDLA2", "", base_nomig$cod)

base_nomig$cod <-  sub("*-.", "", base_nomig$cod)

base_nomig$cod <- substr(base_nomig$cod, 0, 4)

base_nomig$cod <- gsub(" ", "", base_nomig$cod, fixed = TRUE)

base_nomig$cod <- gsub("O", "000", base_nomig$cod)

base_nomig$cod <-  sub(".*000", "", base_nomig$cod)

base_nomig <- base_nomig[complete.cases(base_nomig$cod), ]



#Unir el codigo correspondiente con el nombre de la carrera siempre usado 

base_nomig <- merge(base_nomig,cod_carr,by="cod")


#Cambiar en Etapa a los nuevos nombres de SF

table(base_nomig$Etapa, base_nomig$statecodename)

aux1 <- ifelse(base_nomig$statecodename == "Perdido",1,0)

base_nomig$indicador <- aux1

base_nomig$Etapa[base_nomig$Etapa == "Afluente" & base_nomig$indicador == 1] <- "Cerrada Perdida"
base_nomig$Etapa[base_nomig$Etapa == "Afluente" & base_nomig$indicador == 0] <- "Afluente"
base_nomig$Etapa[base_nomig$Etapa == "Inscrito" & base_nomig$indicador == 1] <- "Cerrada Perdida"
base_nomig$Etapa[base_nomig$Etapa == "Inscrito" & base_nomig$indicador == 0] <- "Inscrito"
base_nomig$Etapa[base_nomig$Etapa == "Matriculado" & base_nomig$indicador == 1] <- "Cerrada Perdida"
base_nomig$Etapa[base_nomig$Etapa == "Matriculado" & base_nomig$indicador == 0] <- "Matrícula"


base_nomig$Etapa <- gsub("Contactado", "Cerrada Perdida", base_nomig$Etapa)
base_nomig$Etapa <- gsub("Documentado", "Cerrada Ganada", base_nomig$Etapa)
base_nomig$Etapa <- gsub("No contactado", "Cerrada Perdida", base_nomig$Etapa)
base_nomig$Etapa <- gsub("Test Aprobado", "Cerrada Perdida", base_nomig$Etapa)
base_nomig$Etapa <- gsub("Test Reprobado", "Cerrada Perdida", base_nomig$Etapa)



#Paso 4. Cruce de registros SF con CRM Dynamics, exclusivos nuevos procesos en SF

base_exSF <- anti_join(base_sf, pre_SF, by="id")

base_exSF2 <- anti_join(base_sf, base_mig, by="id")


#Paso 5. Append de la base completa:
# 1. Registros no migrados (base_nomig)
# 2. Regsitros migrados (base_mig)
# 3. Registros exclusivos SF (base_exSF)


#Seleccion de columnas de interes
#Base de no migrados 

base_nomig <- base_nomig %>%
  select(id, nom, Etapa, carrera, `Propietario de oportunidad`)

#Base de migrados 
base_mig <- base_mig %>%
  select(id, nom, Etapa, carrera, `Propietario de oportunidad`)

#Registros Exclusivos SF
base_exSF <- base_exSF %>%
  select(id, nom, Etapa, carrera, `Propietario de oportunidad`)

base_final <- rbind(base_mig,base_nomig,base_exSF)

base_final$dummy <- as.numeric(base_final$carrera == "ENFERMERIA MODALIDAD ESTUDIO TRABAJO")

base_final <- filter(base_final, dummy == 0)


table(base_final$carrera)
table(base_final$Etapa)


base_final <- base_final[!duplicated(base_final), ]

write_xlsx(base_final,"C:/Users/ritaz/Desktop/CONSULTOR CIERRE/Analítica/AnalisisVarios/MATCH CRM_SF//rep_unidos.xlsx")


#Paso 6.
##Ahora necesito limpiar para poder realizar la accion de Perdidos


#Dataframe del historial de estados OJO subirlo limpio de puntos y con nombres propios

historial <- read_excel("C:/Users/ritaz/Desktop/CONSULTOR CIERRE/Analítica/AnalisisVarios/MATCH CRM_SF/historial.xlsx") 

base_perdidos <- merge(historial,ci_post,by="nom")


base_perdidos$Perdido <- 0

#Doy un 1 a los perdidos dentro del historial
base_perdidos$Perdido[base_perdidos$`Valor nuevo` == "Cerrada Perdida" | base_perdidos$`Valor anterior` == "Cerrada Perdida"] <- 1 
base_perdidos <- filter(base_perdidos, Perdido == 1) #Me quedo solo con los perdidos

#Base solo perdidos, puede ser util
write_xlsx(base_perdidos,"C:/Users/ritaz/Desktop/CONSULTOR CIERRE/Analítica/AnalisisVarios/MATCH CRM_SF//base_perdidos.xlsx")


## Base exclusiva de los perdidos respecto a la base de reporte importante REPUNIDOS

perdidos <- read_excel("C:/Users/ritaz/Desktop/CONSULTOR CIERRE/Analítica/AnalisisVarios/MATCH CRM_SF/base_perdidos.xlsx")
base_final <- read_excel("C:/Users/ritaz/Desktop/CONSULTOR CIERRE/Analítica/AnalisisVarios/MATCH CRM_SF/rep_unidos.xlsx")

base_perdidos <- merge(base_final,perdidos,by="id")

drop <- c("dummy", "nom.y", "Propietario de oportunidad.y", "Valor nuevo", "Programa Pregrado", "Programa Posgrado")
base_perdidos <- base_perdidos[,!(names(base_perdidos) %in% drop)]

base_perdidos <- base_perdidos[!duplicated(base_perdidos), ]


## MERGE final para determinar la base tabulable y presentable

base_final$Perdidos <- 0

library(data.table)
library(plyr)
library(sqldf)

# Left Join: keep all rows in x even if there's no match in y

base_definitiva <- merge(base_final, base_perdidos, all.x = TRUE)
drop2 <- c("dummy", "Perdidos", "nom.x", "Propietario de oportunidad.x", "Programa Pregrado", "Programa Posgrado")

base_definitiva <- base_definitiva[,!(names(base_definitiva) %in% drop2)]

#Aquellos que son cerrados perdidos y no se identifican con 1 los voy a llenar con codigo

base_definitiva$Perdido[base_definitiva$Etapa == "Cerrada Perdida"] <- 1 
base_definitiva$`Valor anterior`[base_definitiva$Perdido == 1] <- "Afluente" 

#Lleno los NA como 0 para que no sesgen la base 

base_definitiva$Perdido[is.na(base_definitiva$Perdido)] <- 0
base_definitiva$`Valor anterior`[is.na(base_definitiva$`Valor anterior`)] <- 0

base_definitiva$Perdido[is.na(base_definitiva$Perdido)] <- 0


#Ahora para ETAPA voy a quitar el Cerrado perdido dado que ya tengo mi enttidad de perdido
#Procedo a llenar si y solo en la entindad de Valor anterior hay algo distinto a 0


base_definitiva <- base_definitiva %>% mutate(Etapa = ifelse(Etapa %in% "Cerrada Perdida", `Valor anterior`, Etapa))
base_definitiva <- base_definitiva[!duplicated(base_definitiva), ]


#Base utilizable
write_xlsx(base_definitiva,"C:/Users/ritaz/Desktop/CONSULTOR CIERRE/Analítica/AnalisisVarios/MATCH CRM_SF//base_definitiva.xlsx")




