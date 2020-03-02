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
Iris <- read_excel("C:/Users/Jaspreet/Desktop/M.Engg/2nd_Sem/DM/Project/DATA/Original Datasets/",dataset_org,".xlsx", col_names = FALSE)
Org_data <- Iris
names(Org_data) <- c("A","B","C","D") #,"E","F","G","H","I","J","K","L","M","N","O","P","Q","R","S","T","U","V",
                     #"W","X","Y","Z","AA","AB","AC","AD","AE","AF","AG","AH","AI","AJ","AK","AL","AM","AN","AO",
                     #"AP","AQ","AR","AS","AT","AU","AV","AW","AX","AY","AZ","BA","BB","BC","BD","BE","BF","BG","BH")

Org_data=as.data.frame(Org_data)                            #Table to Array


View(Org_data)

dim(Org_data)
head(Org_data)
str(Org_data)
summary(Org_data)


#Imploring Incomplete Data
libraby(readxl)
filename <- "Iris/Iris_AE_1"
Imp_data<- read_excel(paste("C:/Users/Jaspreet/Desktop/M.Engg/2nd_Sem/DM/Project/DATA/Incomplete Datasets/",filename,".xlsx", sep=""), col_names = FALSE)
names(Imp_data) <- c("A","B","C","D") #,"E","F","G","H","I","J","K","L","M","N","O","P","Q","R","S","T","U","V",
                     #"W","X","Y","Z","AA","AB","AC","AD","AE","AF","AG","AH","AI","AJ","AK","AL","AM","AN","AO"
                     #,"AP","AQ","AR","AS","AT","AU","AV","AW","AX","AY","AZ","BA","BB","BC","BD","BE","BF","BG","BH")

Imp_data=as.data.frame(Imp_data)                            #Table to Array

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

a=1

#optimum k Factor
for (i in 1:(2*n)){
  
  knn <-  kNN(Imp_data, k=i, numFun = median, catFun = maxCat)
  #knn <-  kNN(Imp_data, k=i, numFun = weightedMean ,weightDist=TRUE , catFun = maxCat)
  Imptd_data <- subset(knn, select = A:D)
  
  Imptd_data.NRMS[i] <- NRMS(Imptd_data, Org_data)
  cat('Overall K.optm',i,'=','NRMS','=',Imptd_data.NRMS[i],'\n')
  if(Imptd_data.NRMS[i]<Imptd_data.NRMS){
    Imptd_data.NRMS= Imptd_data.NRMS[i]
    a = i
    
  }
  else{
    Imptd_data.NRMS=Imptd_data.NRMS
    a=a
    
  }
}


a


#k equal to most accurate k factor
K.fctr <- a            #k optimum i

knn_model <-  kNN(Imp_data, k= K.fctr,numFun = median, catFun = maxCat)     #Knn model
#knn_model <-  kNN(Imp_data, k=K.fctr, numFun = weightedMean ,weightDist=TRUE , catFun = maxCat)
summary(knn_model)
Imputed_data <- subset(knn_model, select = A:D)

#calculate NRMS ECLRMC
Imputed_data.NRMS <- NRMS(Imputed_data, Org_data)
Imputed_data.NRMS
NRMS_Value <- as.data.frame(Imputed_data.NRMS)

Imputed_data.A <- as.data.frame(a)

#Visualize imputed data and original data together for numerical data
par(mfrow=c(1,2))
set.seed(9999)
x <- as.matrix(Org_data_1)
cols <- c("blue", "red")[(x > 0) + 1]  
barplot(x, col = cols, main = "Original Data")

set.seed(9999)
y <- as.matrix(Imputed_data_1)
cols <- c("blue", "red")[(x > 0) + 1]  
barplot(y, col = cols, main = "Imputed Data")


#write imputed data in the file
write.xlsx(Imputed_data, paste("C:/Users/Jaspreet/Desktop/M.Engg/2nd_Sem/DM/Project/DATA/Imputed Datasets/",filename,".xlsx", sep=""),col.names=FALSE)


#write nrms value and Absolute Error/Accuracy
write.xlsx(nrms_value, paste("C:/Users/Jaspreet/Desktop/M.Engg/2nd_Sem/DM/Project/DATA/",filename,"_nrms.xlsx", sep=""))

#write K Factor
write.xlsx(Imputed_data.A, paste("C:/Users/Jaspreet/Desktop/M.Engg/2nd_Sem/DM/Project/DATA/",filename,"_K_Factor.xlsx", sep=""))

