#=============VERSION 3 PESOS============

#==============INSTRUCCIONES=============
#1.Cambiar ruta del fichero de respuestas en linea 22
#2.Cambiar ruta del fichero del marco de informacion auxiliar de la linea 23
#3.Cambiar ruta del Output en linea 274
#4.Ejecutar
#========================================

mis_paquetes <- c("readxl", "Hmisc", "plyr", "dplyr","srvyr","openxlsx")                  #Lista con los paquetes a instalar
no_instalados <- mis_paquetes[!(mis_paquetes %in% installed.packages()[ , "Package"])]    #Comprueba qué paquetes no están instalados y crea una lista con sus nombres
if(length(no_instalados)) install.packages(no_instalados)                                 #Instala los paquetes no instalados
rm(mis_paquetes,no_instalados)

library(readxl) #Libreria lectura excel
library(Hmisc)    
library(plyr)   
library(dplyr)
library(srvyr)    
library(openxlsx) 

datos <- read_excel("C:/Users/Cristina/Desktop/Practicas/EIL1718RESPUESTAS.xlsx")                   #Lee los resultados de las encuestas
marco <- read_excel("C:/Users/Cristina/Desktop/Practicas/EIL1718_marco_informacion_auxiliar.xlsx")  #Lee el archivo de información auxiliar para extraer la variable sexo

marco$TIPO_ESTUDIO<-as.character(marco$TIPO_ESTUDIO)               #Convierte la variable TIPO_ESTUDIO a caracter
marco["TIPO_ESTUDIO"][marco["TIPO_ESTUDIO"] == 5] <- "GRADO"       #Recodifica los 5 como grado en el marco
marco["TIPO_ESTUDIO"][marco["TIPO_ESTUDIO"] == 6] <- "MASTER"      #Recodifica los 6 como master en el marco
marco$TIPO_RAMA <-paste(marco$TIPO_ESTUDIO,marco$RAMA,sep="_")     #Crea la variable TIPO_RAMA dentro de el dataframe "marco"
datos$TOTAL <- "TOTAL"

#PesoV1
total_v1<-summarise(group_by(marco,as.factor(marco$TIPO_RAMA),marco$CENTRO_SIGMA),n_total=n())         #Calcula el N del dominio
n_muestra_v1 <-summarise(group_by(datos,as.factor(datos$TIPO_RAMA),datos$Codigo_Centro),n_muestra=n()) #Calcula el n del dominio
PesoV1<-data.frame(c(total_v1,n_muestra_v1[,3],total_v1[,3]/n_muestra_v1[,3]))                         #Crea un dataframe con el peso V1
colnames(PesoV1)<-c("TIPO_RAMA","Codigo_Centro","n_total ","n_muestra ","Pesos_V1")                    #Cambia los nombres de las variables
datos<-merge(x = datos, y = PesoV1[ , c("TIPO_RAMA","Codigo_Centro", "Pesos_V1")], by = c("TIPO_RAMA","Codigo_Centro"), all.x=TRUE) #Hace un leftjoin entre el dataframe de peso V1 y 
rm(total_v1,PesoV1,n_muestra_v1)#Elimina las variables que no vamos a utilizar mas                                                                                                                               #el dataframe de datos por las variables TiPO_RAMA y CODIGO_CENTRO


#PesoV2
total_v2<-summarise(group_by(marco,(as.factor(marco$TIPO_RAMA))),n_total=n()) #Calcula el N del dominio
n_muestra_v2 <-summarise(group_by(datos,(as.factor(datos$TIPO_RAMA))),n_muestra=n()) #Calcula el n del dominio
PesoV2<-data.frame(c(total_v2,n_muestra_v2[,2],total_v2[,2]/n_muestra_v2[,2])) #Crea un dataframe con el peso V1
colnames(PesoV2)<-c("TIPO_RAMA","n_total ","n_muestra ","Pesos_V2") #Cambia los nombres de las variables
datos<-merge(x = datos, y = PesoV2[ , c("TIPO_RAMA","Pesos_V2")], by = c("TIPO_RAMA"), all.x=TRUE) #Hace un leftjoin entre el dataframe de peso V1 y 
rm(total_v2,PesoV2,n_muestra_v2)#Elimina las variables que no vamos a utilizar mas                                                                                                  #el dataframe de datos por la variable TiPO_RAMA

#Añade la variable sexo al dataset
datos<-merge(x = datos, y = marco[ , c("NIP_ALUMNO", "SEXO","CODIGO_ESTUDIO")], by = c("NIP_ALUMNO","CODIGO_ESTUDIO"), all.x=TRUE)

#Añade las variables TIPO_SEXO y TIPO_CENTRO
datos$TIPO_SEXO   <- paste(datos$D_Tipo_Estudio,datos$SEXO,sep="_")
datos$TIPO_CENTRO <- paste(datos$D_Tipo_Estudio,datos$Codigo_Centro,sep="_")


#Especifica los dominios a estudiar, estos deben ser variables del dataset
dominio<-c("Codigo_Centro","TIPO_ESTUDIO","RAMA","TIPO_RAMA","COD_EST_CAM","TIPO_CENTRO","SEXO","TIPO_SEXO","TOTAL") #D_Tipo_Estudio?

#Agrupo los nombres de las variables para que sea más cómodo y claro trabajar con ellas
#Variables categoricas
preg_c<-c("P1G","P25","P3","P25_2","P21","P2","P2B","P2C","P4","P4A","P4B","P5",
          "P6","P7","P9A","P9B","P10","P11","P12","p13rama","p14funcion","P15","P15_2",
          "P16","P17","P21A","P21B","P23_1","P23_2","P23_3","P23_4","P23_5","P23_6","P23_7",
          "P23_8","P23_9","P23_10","P23_11","P23_98","P23_99","P23A","P24","P24A","P26",
          "P28_1","P28_2","P28_3","P28_4","P28_5","P28_6","P28_7","P28_8","P28_9","P28_10",
          "P28_11","P29COD","P1M")


#Variables no categoricas
preg_nc<-c("P2A_1","P18A_1","P18A_2","P18A_3","P18A_4","P18A_5","P18A_6","P18B",
            "P19_1","P19_2","P19_3","P21A1_1","P21B1_1","P21C1_1","P21D1_1")

#Reemplazo los "99" por NA en las variables categoricas
df_nonreplace <- select(datos, all_of(preg_c))
df_replace <- datos[ ,!names(datos) %in% names(df_nonreplace)]
df_replace <- replace(df_replace, df_replace == 99,NA)
datos <- cbind(df_nonreplace, df_replace)
rm(df_nonreplace,df_replace) #Elimina las variables que no vamos a utilizar mas

attach(datos) #Con attach nos evitamos mencionar el dataframe al que pertenecen las variables continuamente cuando estas van a ser siempre del mismo, podremos usar P3 en vez de datos$P3
T1<-T2<-T3<-T4<-T5<-c() #Crea las variables de tasas
for(i in 1:(dim(datos)[1])){
  #T1
  if((P3[i]==1 || P3[i]==2 || P3[i]==4 || P3[i]==5 || P3[i]==7 || P3[i]==13) || ((P3[i]==3 || P3[i]==6 || P3[i]==8 || P3[i]==12) && 
                                                                                 (P25[i]==1))){T1[i]<-1
  }else{T1[i]<-0}
  
  #T2
  if((P3[i]==3 || P3[i]==6 || P3[i]==8 || P3[i]==11 || P3[i]==12) && 
     (!is.na(P25_2[i]) && (P25_2[i]==1 || P25_2[i]==2 || P25_2[i]==3 || P25_2[i]==5 || P25_2[i]==6 || P25_2[i]==7 || P25_2[i]==10 || P25_2[i]==11))){T2[i]<-1
  }else{T2[i]<-0}
  
  #T3
  if(P3[i]==1 || P3[i]==2 || P3[i]==4 || P3[i]==5 || P3[i]==7 || P3[i]==13){T3[i]<-1
  }else{T3[i]<-0}
  
  #T4
  if((P3[i]==3 || P3[i]==6 || P3[i]==8 || P3[i]==12)&&(P25[i]==1)){T4[i]<-1
  }else{T4[i]<-0}
  
  #T5
  if((P3[i]==6 || P3[i]==8 || P3[i]==11) && (!is.na(P25_2[i]) && (P25_2[i] == 12 || P25_2[i] == 8))){T5[i]<-1
  }else{T5[i]<-0}
}
datos<-cbind(datos,T1,T2,T3,T4,T5)             #Introduce las variables de tasas en el dataframe
detach(datos)                                  #Deja de usar el attach
preg_nc<-c(preg_nc,"T1","T2","T3","T4","T5")   #Incluye las tasas en la lista de preguntas no categoricas

#A continuación crea y recodifica cada una de las variables que deseamos añadir
#========================P10cod==============================
datos$P10cod <- dplyr::recode(datos$P10, '1' = "1", 
                                         '2' = "3",
                                         '3' = "1",
                                         '4' = "2",
                                         '5' = "2",
                                         '6' = "5",
                                         '7' = "3",
                                         '8' = "4",
                                         '9' = "3",
                                         '10' = "3",
                                         '11' = "3",
                                         '12' = "3",
                                         '99' = "99")
#table(datos$P10cod)
#===========================================================

#=========================T3*p15==========================
datos$T3_P15 <- paste(datos$T3,datos$P15,sep="_") #Junta la combinación de variables deseada separando sus valores con "_"
datos$T3_P15 <- dplyr::recode(datos$T3_P15, '1_NA' = "99", 
                                            '1_1' = "1",
                                            '1_2' = "2", 
                                            .default= NA_character_)
#table(datos$T3_P15)
#===========================================================


#=========================T3*p10cod==========================
datos$T3_P10cod <- paste(datos$T3,datos$P10cod,sep="_")
datos$T3_P10cod <- dplyr::recode(datos$T3_P10cod, '1_1' = "1", 
                                                  '1_2' = "2",
                                                  '1_3' = "3", 
                                                  '1_4' = "4",
                                                  '1_5' = "5",
                                                  '1_99' = "99",
                                                  .default= NA_character_)
#table(datos$T3_P10cod)
#===========================================================

#=================t3*p15*p10cod============================
datos$T3_P15_P10cod <- paste(datos$T3,datos$P15,datos$P10cod,sep="_")
datos$T3_P15_P10cod <- dplyr::recode(datos$T3_P15_P10cod, '1_NA_NA' = "1_99_99", 
                                                          '1_1_1' = "1_1_1",
                                                          '1_1_2' = "1_1_2", 
                                                          '1_1_3' = "1_1_3",
                                                          '1_1_4' = "1_1_4",
                                                          '1_1_5' = "1_1_5",
                                                          '1_1_99' = "1_1_99",
                                                          '1_2_1' = "1_2_1",
                                                          '1_2_2' = "1_2_2",
                                                          '1_2_3' = "1_2_3",
                                                          '1_2_4' = "1_2_4",
                                                          '1_2_5' = "1_2_5",
                                                          '1_2_99' = "1_2_99",
                                                          .default= NA_character_)
#table(datos$T3_P15_P10cod)
#===========================================================


#======================t3*p11=====================================
datos$T3_P11 <- paste(datos$T3,datos$P11,sep="_")
datos$T3_P11 <- dplyr::recode(datos$T3_P11, '1_NA' = "99", 
                                            '1_1' = "1_1",
                                            '1_2' = "1_2", 
                                            '1_3' = "1_3",
                                            '1_4' = "1_4",
                                            .default= NA_character_)
#table(datos$T3_P11)
#===========================================================

#=========================t3*p9*p9a===============================
datos$T3_P9_P9A <- paste(datos$T3,datos$P9,datos$P9A,sep="_")
datos$T3_P9_P9A <- dplyr::recode(datos$T3_P9_P9A, '1_NA_NA' = "99", 
                                                  '1_1_NA' = "99",
                                                  '1_1_1' = "1", 
                                                  '1_1_2' = "2",
                                                  '1_2_1' = "1",
                                                  '1_2_2' = "2",
                                                  '1_5_NA' = "99",
                                                  '1_5_1' = "1",
                                                  '1_5_2' = "2",
                                                  .default= NA_character_)
#table(datos$T3_P9_P9A)
#============================================================

preg_c<-c(preg_c,"P10cod","T3_P15","T3_P10cod","T3_P15_P10cod","T3_P11","T3_P9_P9A") #Inluye las nuevas preguntas en la lista de variables categoricas 


wb <- createWorkbook()        #Crea el workbook (conjunto de datos que finalmente será el excel)

#Cada iteración de este bucle son calculos para cada dominio, dentro de cada dominio hace calculos 
#para variables categoricas y no categoricas y las graba en una hoja excel
for(k in 1:(length(dominio))){ #Bucle que opera en cada dominio
  
  #Crea un dataframe para cada uno de los tipos de variables 
  data_nc<-subset(datos[!(is.na(datos[[dominio[k]]])),],select=c(dominio[[k]],preg_nc,"Pesos","Pesos_V1","Pesos_V2")) #Dataframe con las variables no categóricas y sus pesos
  data_c<-subset(datos[!(is.na(datos[[dominio[k]]])),],select=c(dominio[[k]],preg_c,"Pesos","Pesos_V1","Pesos_V2"))   #Dataframe con las variables categóricas y sus pesos
  
  #===Media Categóricas====
  DatosTotal <- data.frame()                                                                         #Crea el dataframe que va a ir rellenando con las medias que calcule
  for(j in 1:(length(preg_c))){                                                                      #Bucle que hará los calculos correspondientes a cada pregunta de tipo categorico
    final2<-summarise(group_by(data_c[!(is.na(data_c[[preg_c[j]]])),], data_c[!(is.na(data_c[[preg_c[j]]])),][[dominio[[k]]]],data_c[!(is.na(data_c[[preg_c[j]]])),][[preg_c[j]]]),n = n()) #Calcula "n" para cada valor diferente de esa pregunta
    sumas_totales_SinPesos<-aggregate(final2$n, list(as.matrix(final2[,1])), sum)                    #Calcula la suma total de las n de cada pregunta
    final2 <- merge(final2, sumas_totales_SinPesos, by.x = 1, by.y = 1, all.x = TRUE, all.y = TRUE)  #Añade n y las sumas totales al dataframe de variables categoricas
    final2$media_sinPesos<-final2$n/final2$x                                                         #Calcula la media sin pesos
    for(i in 1:length(unique(data_c[[dominio[[k]]]]))){
      datosH<-data_c[data_c[[dominio[[k]]]]==unique(sort(data_c[[dominio[[k]]]]))[i],] #Extrae los datos correspondientes a cada pregunta de ese dominio
      
      sumas_V0<- aggregate(x = datosH$Pesos,by = list(datosH[[preg_c[j]]]),FUN = sum)#Calcula la suma de los pesos V0 para cada elemento del dominio
      suma_total_V0<-sum(datosH[!is.na(datosH[[preg_c[j]]]),]$Pesos) #Calcula el la suma total de los pesos V0 del dominio
      
      sumas_V1<- aggregate(x = datosH$Pesos_V1,by = list(datosH[[preg_c[j]]]),FUN = sum)#Calcula la suma de los pesos V1 para cada elemento del dominio
      suma_total_V1<-sum(datosH[!is.na(datosH[[preg_c[j]]]),]$Pesos_V1) #Calcula el la suma total de los pesos V1 del dominio
      
      sumas_V2<- aggregate(x = datosH$Pesos_V2,by = list(datosH[[preg_c[j]]]),FUN = sum)#Calcula la suma de los pesos V2 para cada elemento del dominio
      suma_total_V2<-sum(datosH[!is.na(datosH[[preg_c[j]]]),]$Pesos_V2) #Calcula el la suma total de los pesos V2 del dominio
      
      if(suma_total_V0==0||dim(sumas_V0)[1]==0){next} #Si la media es 0 (no hay respuestas para hacer la media), salta directamente al siguiente elemento del bucle
      if(exists("datosH2")){ #Si datos H2 existe
        datosH3<-unique(sort(datosH[[preg_c[j]]])) #Crea H3
        datosH_Media_V0<-cbind(sumas_V0$x/suma_total_V0) #Calcula la media V0 de cada elemento del dominio
        datosH_Media_V1<-cbind(sumas_V1$x/suma_total_V1) #Calcula la media V1 de cada elemento del dominio
        datosH_Media_V2<-cbind(sumas_V2$x/suma_total_V2) #Calcula la media V2 de cada elemento del dominio
        
        datosH3<-cbind(unique(sort(data_c[[dominio[[k]]]]))[i],datosH3,datosH_Media_V0,datosH_Media_V1,datosH_Media_V2) #Incluye en H3 las diferentes columnas que luego veremos en el excel
        datosH2<-rbind(datosH2,datosH3) #Junta H2 y H3
      }else{ #Si datosH2 no existe
        datosH2<-unique(sort(datosH[[preg_c[j]]])) #Crea Datos H2
        datosH_Media_V0<-cbind(sumas_V0$x/suma_total_V0) #Calcula la media V0 de cada elemento del dominio
        datosH_Media_V1<-cbind(sumas_V1$x/suma_total_V1) #Calcula la media V1 de cada elemento del dominio
        datosH_Media_V2<-cbind(sumas_V2$x/suma_total_V2) #Calcula la media V2 de cada elemento del dominio
        datosH2<-cbind(unique(sort(data_c[[dominio[[k]]]]))[i],datosH2,datosH_Media_V0,datosH_Media_V1,datosH_Media_V2)} #Incluye en H2 las diferentes columnas que luego veremos en el excel
    }
    if(exists("datosH2")){
    categorica<-as.data.frame(cbind(datosH2[,1],preg_c[j],datosH2[,2],final2$n,datosH2[,3],final2$media_sinPesos,datosH2[,4],datosH2[,5],"CAT")) #Guarda en el dataframe lo que incluiremos en el excel
    colnames(categorica)<-c(dominio[[k]],"Var","VarLevel","n","Mean_V0","Mean_SinPesos","Mean_V1","Mean_V2","TIPO_VAR") #Cambia los nombres de las variables
    DatosTotal<-rbind(DatosTotal,categorica) #Incluye en datosTotal la información de la nueva pregunta que estaba contenida en categorica
    rm(final2,datosH,sumas_V0,sumas_V1,sumas_V2,suma_total_V0,suma_total_V1,suma_total_V2,datosH3,datosH_Media_V0,datosH_Media_V1,datosH_Media_V2,categorica,datosH2,sumas_totales_SinPesos)} #Borra las variables que no necesitaremos fuera del bucle
  }
  
  #===Media Numericas===
  for(i in 1:(length(preg_nc))){ #Bucle que opera en cada una de las variables no categoricas
    data_sin_pesos<-ddply(data_nc,~data_nc[[dominio[[k]]]],function(data_nc)mean(data_nc[[preg_nc[i]]],na.rm=TRUE)) #Calcula la media sin pesos
    data_V0<-ddply(data_nc,~data_nc[[dominio[[k]]]],function(data_nc)wtd.mean(data_nc[[preg_nc[i]]],data_nc$Pesos,na.rm=TRUE)) #Calcula la media con los pesos V0
    data_V1<-ddply(data_nc,~data_nc[[dominio[[k]]]],function(data_nc)wtd.mean(data_nc[[preg_nc[i]]],data_nc$Pesos_V1,na.rm=TRUE)) #Calcula la media con los pesos V1
    data_V2<-ddply(data_nc,~data_nc[[dominio[[k]]]],function(data_nc)wtd.mean(data_nc[[preg_nc[i]]],data_nc$Pesos_V2,na.rm=TRUE)) #Calcula la media con los pesos V2
    
    var<- rep(preg_nc[i],dim(data_V0[!(is.na(data_V0[,"V1"])),])[1]) #Vector que será la columna con el nombre de la variable
    varLevel<- rep(NA,dim(data_V0[!(is.na(data_V0[,"V1"])),])[1]) #Vector que será la columna con el nivel de la variable
    final2<-summarise(group_by(data_nc[!(is.na(data_nc[[preg_nc[i]]])),], data_nc[!(is.na(data_nc[[preg_nc[i]]])),][[dominio[[k]]]]),n = n()) #Calculan para cada nivel de la variable
    data2<-data.frame(data_V0[!(is.na(data_V0[,"V1"])),]$`data_nc[[dominio[[k]]]]`,var,varLevel,final2$n,na.exclude(data_V0$V1),na.exclude(data_sin_pesos$V1),na.exclude(data_V1$V1),na.exclude(data_V2$V1),"NUM") #Junta todas las columnas que añadira al excel
    
    colnames(data2)<-c(dominio[[k]],"Var","VarLevel","n","Mean_V0","Mean_SinPesos","Mean_V1","Mean_V2","TIPO_VAR") #Cambia el nombre de las columnas
    DatosTotal<-rbind(DatosTotal,data2) #Añade los calculos de la pregunta actual al dataframe que teniamos con los resultados de las categoricas y no categoricas en ese dominio
  }

  addWorksheet(wb, dominio[[k]]) #Añade una hoja al excel con el nombre de el dominio con el que estamos trabajando
  writeDataTable(wb, dominio[[k]], DatosTotal) #Graba en el excel el dataframe DatosTotal en el que teniamos los calculos de las variables categoricas y no categoricas de ese dominio
  freezePane(wb,dominio[[k]], firstRow = TRUE) #Da formato a la primera fila del excel (formato tabla)
  setColWidths(wb,dominio[[k]],cols = 1:ncol(DatosTotal),widths = "auto") #Ajusta el ancho de las columnas del excel
  rm(data_V0,var,varLevel,final2,data2,DatosTotal,data_sin_pesos) #Elimina las variables que no son necesarias al acabar el bucle
  
}
#Graba el workbook creado en un archivo excel
saveWorkbook(wb, file = "C:/Users/Cristina/Desktop/Practicas/Output_20-09 - 3Pesos.xlsx",overwrite = TRUE) 




