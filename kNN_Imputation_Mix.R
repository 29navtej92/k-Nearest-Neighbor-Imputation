#Packages needed to be installed

install.packages("xlsx")
install.packages("rJava")
install.packages("VIM")
install.packages("colorspace")
install.packages("data.table")
install.packages("grid")
install.packages("laeken")
install.packages("ECLRMC")
install.packages("softImpute")
install.packages("Matrix")
install.packages("Metrics")
install.packages("dplyr")
install.packages("caret")
install.packages("pROC")
install.packages("mlbench")
install.packages("hydroGOF")
install.packages("missForest")


#libraries used
library(readxl)
library(rJava)
library(VIM)
library(colorspace)
library(grid)
library(data.table)
library(laeken)
library(ECLRMC)
library(softImpute)
library(Matrix)
library(dplyr)
library(Metrics)
library(laeken)
library(caret)
library(pROC)
library(mlbench)
library(hydroGOF)
library(missForest)


##Imploring and Preparing the data

#Imploring Complete Data Set

library(readxl)
Org_dataset <- read_excel("C:/Users/Jaspreet/Desktop/M.Engg/2nd_Sem/DM/Project/DATA/Original Datasets/Aheart.xlsx", col_names = FALSE)
Org_data <- Aheart
names(Org_data) <- c("A","B","C","D","E","F","G","H","I")  #,"J","K","L","M","N","O","P","Q","R","S","T","U","V",
                     #"W","X","Y","Z","AA","AB","AC","AD","AE","AF","AG","AH","AI","AJ","AK","AL","AM","AN","AO",
                     #"AP","AQ","AR","AS","AT","AU","AV","AW","AX","AY","AZ","BA","BB","BC","BD","BE","BF","BG","BH")

Org_data=as.data.frame(Org_data)                            #Table to Array


Org_data=Org_data %>% mutate_if(is.character, as.factor)    # Mutate character variables into factors
View(Org_data)

dim(Org_data)
head(Org_data)
str(Org_data)
summary(Org_data)

#For mix dataset dividing data into categorical and numarical

Org_data_n <- subset(Org_data, select = c(A,B,C,D,F,G,H,I)) #numarical dataset
Org_data_c <- subset(Org_data, select = E) #categorical dataset

#Imploring Incomplete Data
libraby(readxl)
filename <- "Aheart/Aheart_AE_1"
Imp_data<- read_excel(paste("C:/Users/Jaspreet/Desktop/M.Engg/2nd_Sem/DM/Project/DATA/Incomplete Datasets/",filename,".xlsx", sep=""), col_names = FALSE)
names(Imp_data) <- c("A","B","C","D","E","F","G","H","I")  #,"J","K","L","M","N","O","P","Q","R","S","T","U","V",
                     #"W","X","Y","Z","AA","AB","AC","AD","AE","AF","AG","AH","AI","AJ","AK","AL","AM","AN","AO"
                     #,"AP","AQ","AR","AS","AT","AU","AV","AW","AX","AY","AZ","BA","BB","BC","BD","BE","BF","BG","BH")

Imp_data=as.data.frame(Imp_data)                            #Table to Array


Imp_data=Imp_data %>% mutate_if(is.character, as.factor)    # Mutate character variables into factors
View(Imp_data)

dim(Imp_data)
head(Imp_data)
str(Imp_data)
summary(Imp_data)



#Missing Pattern preprocessing
Imp_data=as.data.frame(Imp_data)
Imp_data[ Imp_data == "" ] = NA   # fill empty cells with NA

set.seed(958)
str(Imp_data)

# Extract the pattern of missing data
NAloc <- is.na(Imp_data)  #location of missing data
noNAvar <- apply(NAloc, 2, sum)  # how many are missing

# Shows missing data and pattern
miss_plot <- aggr(Imp_data, col=c('navyblue','yellow'),
                  numbers=TRUE, sortVars=TRUE,
                  labels=names(Imp_data), cex.axis=.7,
                  gap=4, ylab=c("Missing data","Pattern"))


#Impute using K Nearest Neighbor Imputation Algorithm
#now to predict and impute the missing values
n <- ceiling(sqrt(nrow(Org_data)))

# declaration to initiate for loop
Imptd_data.NRMS=1
Imptd_data.accuracy=0
Overall=1

a=1

#optimum k Factor
for (i in 1:(n)){
  
  knn <-  kNN(Imp_data, k=i, numFun = median, catFun = maxCat)
  #knn <-  kNN(Imp_data, k=i, numFun = weightedMean ,weightDist=TRUE , catFun = maxCat)
  Imptd_data_n <- subset(knn, select = c(A,B,C,D,F,G,H,I))
  Imptd_data_c <- subset(knn, select = E)
  
  Imptd_data.NRMS[i] <- NRMS(Imptd_data_n, Org_data_n)
  cat('Overall K.optm',i,'=','NRMS','=',Imptd_data.NRMS[i],'\n')
  
  Imptd_data.accuracy[i] <- accuracy(Org_data_c, Imptd_data_c)
  cat('Overall K.optm',i,'=','accuracy','=',Imptd_data.accuracy[i],'\n')    # #to print % accuracy
  
  Overall[i] <- ((Imptd_data.NRMS[i]+(1-Imptd_data.accuracy[i]))/2) 
  
  #cat('Overall K.optm',i,'=', Overall[i],'\n','\n')
  
  if(Overall[i]<Overall){
    Overall = Overall[i]
    a = i
    
  }
  else{
    Overall=Overall
    a=a
    
  }
  
}


a


#k equal to most accurate k factor
K.fctr <- a            #k optimum i

knn_model <-  kNN(Imp_data, k= K.fctr,numFun = median, catFun = maxCat)     #Knn model
#knn_model <-  kNN(Imp_data, k=K.fctr, numFun = weightedMean ,weightDist=TRUE , catFun = maxCat)
summary(knn_model)
Imputed_data <- subset(knn_model, select = A:I)
Imputed_data_n <- subset(knn_model, select = c(A,B,C,D,F,G,H,I))
mputed_data_c <- subset(knn_model, select = E)



#calculate NRMS through ECLRMC
Imputed_data.NRMS <- NRMS(Imputed_data_n, Org_data_n)
Imputed_data.NRMS
NRMS_Value <- as.data.frame(Imputed_data.NRMS)

#Absolute Error/Accuracy through Metrics
Imputed_data.accuracy <- accuracy(Org_data_c, Imputed_data_c)
Imputed_data.accuracy
accuracy_value <- as.data.frame(Imputed_data.accuracy)

Imputed_data.A <- as.data.frame(a)

#write imputed data in the file
filename<- "Aheart_AE_1"
write.xlsx(Imputed_data, paste("C:/Users/Jaspreet/Desktop/M.Engg/2nd_Sem/DM/Project/DATA/Imputed Datasets/Aheart/",filename,".xlsx", sep=""),col.names=FALSE)
write.xlsx(knn_model, paste("C:/Users/Jaspreet/Desktop/M.Engg/2nd_Sem/DM/Project/DATA/Imputed Datasets Logical/Aheart_logical/",filename,".xlsx", sep=""),col.names=FALSE)


#write nrms value and Absolute Error/Accuracy
write.xlsx(nrms_value, paste("C:/Users/Jaspreet/Desktop/M.Engg/2nd_Sem/DM/Project/DATA/Aheart/",filename,"_nrms.xlsx", sep=""))
write.xlsx(accuracy_value, paste("C:/Users/Jaspreet/Desktop/M.Engg/2nd_Sem/DM/Project/DATA/Aheart/",filename,"_accuracy.xlsx", sep=""))

#write K Factor
write.xlsx(Imputed_data.A, paste("C:/Users/Jaspreet/Desktop/M.Engg/2nd_Sem/DM/Project/DATA/Aheart/",filename,"_K_Factor.xlsx", sep=""))
