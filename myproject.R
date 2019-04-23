#####################################################################################################################

#Project Description and Problem statement -

#XYZ is a courier company. #As we appreciate that human capital plays an important rolein collection, transportation 
 # and delivery. The company is passing through genuine issue of Absenteeism

#Human Resource plays a vital role. Lets see how we can help the company

#We will follow CRISP-DM process.

#######################################################################################################################


#######################################################################################################################
#Let"s Start Absentism Project

#Clear envt
rm(list = ls())


#Lets set our Working Directory
setwd("F:/Data Science/Project Employee Absentism")

#Cross Check
getwd()

#Lets Load our File.
library("rJava")
#Sys.setenv(JAVA_HOME="C:\\Program Files\\Java\\jre1.8.0_144") 
#library(rJava)
#install.packages("rJava")

library("xlsx")

d=read.xlsx("Absenteeism_at_work_Project.xlsx",sheetIndex = 1)

#Lets read our file
head(d) 

########################################################################################3
#Lets start with basic data pre-processing

#Check Variable Names
colnames(d)

#Check Shape and Structure of Data

str(d)
dim(d)
summary(d)
#Check the class
class(d)

#Replace . with _
#names(d)=gsub("\\.","_",names(d))

#DataExplorer::create_report(d,output_format = pdf())

#Data Explorer automatically creates report on Data Exploration which includes Missing Value anlysis, Basic Preprocessing 
#report and some advanced concept also


###################################################################################################################
#Data Cleaning

#Remove Variable named ID, as we are not focusing on each employees. We want to know reason of absenteeism

d= subset(d,select=-c(ID))

#Month has 12 values only. Inorder to avoid future problems replace 0 with NA
d$Month.of.absence[d$Month.of.absence %in% 0]= NA

#Dividing Work_load_Average/day_ by 1000 
d$Work.load.Average.day.= d$Work.load.Average.day./1000

#divide the data

num= c('Distance.from.Residence.to.Work', 'Service.time', 'Age','Work.load.Average.day.', 'Transportation.expense',
       'Hit.target', 'Weight', 'Height', 
       'Body.mass.index', 'Absenteeism.time.in.hours')

cat= c('Reason.for.absence','Month.of.absence','Day.of.the.week',
       'Seasons','Disciplinary.failure', 'Education', 'Social.drinker',
       'Social.smoker', 'Son', 'Pet')


###############----------------------------------------------------------------------------#######################
##############################----------------ADVANCED DATA_PREPROCESSING-----------------#########################

###############################----------------------Missing Value Analysis-------------#######################

#Lets check missing values
sum(is.na(d))

#total 138 missing values

df=d

 ###according to the inductrial standards we would impute missing values rather than ignoring them.

#Checking Missing values in Target Variable.

sum(is.na(df$Absenteeism.time.in.hours))

#Removing those observations
df= df[(!df$Absenteeism.time.in.hours %in% NA),]

#22 observations are removed

#Calculate Missing Vales column wise.
apply(df,2,function(x){sum(is.na(x))})

#Now Create a Data Frame
Missing_Values= data.frame(apply(df,2,function(x){sum(is.na(x))}))

#Represent Data Frame in a good Manner

Missing_Values$Variable=row.names(Missing_Values)
row.names(Missing_Values)=NULL  
names(Missing_Values)[1]='Missing Percentage'
Missing_Values=Missing_Values[,c(2,1)]

#Convert them into Percentage

Missing_Values$`Missing Percentage`=(Missing_Values$`Missing Percentage`/nrow(df))*100

#Sort the rows in accordance with their missing values
Missing_Values=Missing_Values[order(-Missing_Values$`Missing Percentage`),]


#Plot a graph for visualising top variables with missing values
#install.packages('ggplot2')

library('ggplot2')

ggplot(data = Missing_Values[1:5,], aes(x=reorder(Variable,`Missing Percentage`),y=`Missing Percentage`))+
  geom_bar(stat = 'identity',fill = 'grey') + xlab("Variables")+
  ggtitle("Missing Data Visualizaiton")+theme_grey()

###########################################################################################

#From the graph we can see that 
#BOdy Mass Index holds highest number of missing values.
#Transportation Expense have least missing values.

#Lets impute for Categorical Data first

MODE=function(x){
  unique_values=unique(x)
  unique_values[which.max(tabulate(match(x,unique_values)))]
}


for(i in cat){
  print(i)
  df[,i][is.na(df[,i])] = MODE(df[,i])
}

#Check Missing Values Now
sum(is.na(df))

#Out of 138, 72 are left.


#Now according to the SOP, we will impute missing values.
#Lets check which method would be suitable
#Original Value=30
#Mean Value=26.68
#Median Value=25
#KNN= '
#k=3 (30.00)
      

df$Body.mass.index[1]

#Now remove 4th Value and Replace with different method
#Mean Method

df$Body.mass.index[1]=NA

df$Body.mass.index[1]=mean(df$Body.mass.index,na.rm = T)

df$Body.mass.index[1]

#Median
df$Body.mass.index[1]=NA

df$Body.mass.index[1]=median(df$Body.mass.index,na.rm = T)

df$Body.mass.index[1]

#KNN Method
#Library
library('DMwR')

df$Body.mass.index[1]=NA

df=knnImputation(df,k=3)

df$Body.mass.index[1]


#Cross Check
sum(is.na(df))

#So there are 0 missing values left now
#From the all methods used, KNN gave the most close value. hence we will lock KNN Imputation
#########################################################################################################################

##########################---------------OUTLIERS ANALYSIS--------------------#########################################

#Now we will go for Outliers Analysis

#First plot the boxplot


for(i in 1:length(num)) {
assign(paste0("gn",i), ggplot(data = df, aes_string(y = num[i])) +
       stat_boxplot(geom = "errorbar", width = 0.5) +
        geom_boxplot(outlier.colour = "red", fill = "grey", outlier.size = 1) +
         labs(y = num[i],x='Absenteeism.time.in.hours') +
          ggtitle(paste("Boxplot: ",num[i])))
}


#Plot Box Plot step wise with the help of Grid Extra
library('gridExtra')

gridExtra::grid.arrange(gn1,gn2,gn3,gn4,ncol=2)
gridExtra::grid.arrange(gn5,gn6,gn7,ncol=2)
gridExtra::grid.arrange(gn7,gn8,gn9,gn10,ncol=2)

#Remove outliers from dataset-
#1st Step- Detect Outliers

for(i in num)
  {
  print(i)
  val= df[,i][df[,i] %in% boxplot.stats(df[,i])$out]
  print(length(val))
  print(val)
} 


#2nd Step- Replace Outliers with NA

for (i in num) {
  val = df[,i][df[,i] %in% boxplot.stats(df[,i])$out]
  df[,i][df[,i] %in% val] = NA
}

sum(is.na(df))

#3rd Step- Check NA values
for (i in num) {
  print(i)
  print(sum(is.na(df[,i])))
}

sum(is.na(df))


#As we have already done Missing Value analysis, KNN @K=3 got us the perfect results,
  #we will do the same here

df=knnImputation(df,k=3)

#Check if there are any missing values left and as crosschecked our data is free from Outliers.
sum(is.na(df))
dim(df)

#Check Outliers after KNN Imputation
for(i in 1:length(num)) {
  assign(paste0("gn",i), ggplot(data = df, aes_string(y = num[i])) +
           stat_boxplot(geom = "errorbar", width = 0.5) +
           geom_boxplot(outlier.colour = "red", fill = "grey", outlier.size = 1) +
           labs(y = num[i],x='Absenteeism.time.in.hours') +
           ggtitle(paste("Boxplot: ",num[i])))
}



gridExtra::grid.arrange(gn1,gn2,gn3,gn4,ncol=2)
gridExtra::grid.arrange(gn5,gn6,gn7,ncol=2)
gridExtra::grid.arrange(gn7,gn8,gn9,gn10,ncol=2)


####################################################################################################################


##########################---------------FEATURE SELECTION--------------------#########################################

#Feature Selection, we will select only those variables which explains TV.

#For that we will plot Corrgram Plot
#install.packages('corrgram')
library('corrgram')

corrgram(df,order = F,upper.panel = panel.pie,lower.panel = panel.shade,text.panel = panel.txt,
         main='Corrgram Plot')

#Dark Blue color indicates positive corrrelation and Dark Red color indicates Negative Correlation.

#Lets go for ANOVA
#install.packages('lsr')

library('lsr')

anova_test=aov(Absenteeism.time.in.hours~Reason.for.absence+Month.of.absence+Day.of.the.week+Seasons+
              Disciplinary.failure+Education+Son+Social.drinker+Social.smoker+Pet,data = df)

#Lets check out the summary
summary(anova_test)



#Now remove highly correlated values
df=subset(df,select=-c(Weight,Education,Seasons,Social.smoker,Pet))
df=subset(df,select=-c(Month.of.absence))

#Body Mass index itself is a measure of how physically fit you are so from the knowledge we can only use Body Mass Index
  #and ignore Height and Weight

colnames(df)
dim(df)
###############################---------------------------------##################################
#Now we will update categorical and continuous variables by ingnoring Target Variable
num_data= c('Work.load.Average.day.','Hit.target','Body.mass.index','Transportation.expense',
            'Distance.from.Residence.to.Work', 'Service.time', 
           'Age','Absenteeism.time.in.hours','Height')

cat_data=c("Reason.for.absence", "Day.of.the.week",'Son', 'Social.drinker','Disciplinary.failure')

colnames(df)



####################################################################################################################


##########################---------------FEATURE ENGINEERING--------------------#########################################

#Lets go for Feature Scaling/ Feature Engineering

#we need to plot distribution curves first inorder to select a suitable method
hist(df$Absenteeism.time.in.hours)
qqnorm(df$Absenteeism.time.in.hours)
hist(df$Body.mass.index)
hist(df$Work.load.Average.day.)

#It is clear from the graph that our data is not at all normally distributed, Hence we will go for Normalizaiton
#Logarithamic Transform

df$Absenteeism.time.in.hours = log1p(df$Absenteeism.time.in.hours)

#Normalization

for(i in num_data)
{
  if(i !='Absenteeism.time.in.hours'){
  print(i)
  df[,i] = (df[,i] - min(df[,i]))/(max(df[,i]-min(df[,i])))
}
}
  
hist(df$Absenteeism.time.in.hours)
summary(df)

#Create a Fresh/New Data set which contains cleaned data

write.csv(df,"Cleaned Data.csv",row.names = T)


###########################################################################################################


##########################---------------SAMPLING--------------------#########################################

#Now our data is cleaned, henceforth we will go for SAMPLING inorder to feed the data to model

#install.packages('dummies')

#Sampling
#devtools::install_github("r-lib/rlang", build_vignettes = TRUE)
#install.packages('caret')
library('caret')
set.seed(101)
Sampling=createDataPartition(df$Absenteeism.time.in.hours,p=0.60,list = FALSE)
Train_Data=df[Sampling,]
Test_Data=df[-Sampling,]

####################################################################################################################


#####################################-------MODELS WIHTOUT DUMMIES---------########################################


##########################---------------DESICION TREE--------------------#########################################

#Now lets start building the model.


#Deicison Tree
library('rpart')

Decision_Tree=rpart(Absenteeism.time.in.hours~.,data=Train_Data,method='anova')

summary(Decision_Tree)
plot(Decision_Tree)
#Now lets preidct Test Cases
Predict_DT=predict(Decision_Tree,Test_Data[,-14])

#Predict_DT
summary(Predict_DT)

plot(Test_Data$Absenteeism.time.in.hours,type="l",lty=1.8,col="Green",main="DT")
lines(Predict_DT,type="l",col="Blue")

##oR
library('DMwR')

Per_DT=print(postResample(obs = Test_Data[,14],pred = Predict_DT))

#Decsisoin Tree
#RMSE=0.48
#Rsq=0.45
#MAE=0.35

plot(Test_Data$Absenteeism.time.in.hours,Predict_DT)
######################################################################################################################


#############################---------------RANDOM FOREST--------------------#########################################

#Random Forest
library('randomForest')
Rf_Model=randomForest(Absenteeism.time.in.hours~.,Train_Data,ntree=300)

plot(Rf_Model)

summary(Rf_Model)

Rf_Predict=predict(Rf_Model,Test_Data[,-14],type = 'class')

summary(Rf_Predict)

Per_RF=print(postResample(pred = Rf_Predict,obs = Test_Data[,14]))

#Random Forest
#RMSE=0.44      
#Rsq=0.53
#MAE=0.33

####################################################################################################################


#############################---------------LINEAR REGRESSION--------------------#########################################

#Linear Regression

LR_Model=lm(Absenteeism.time.in.hours~.,Train_Data)

summary(LR_Model)

#Predict Test cases by ignoring TV

LR_Predict=predict(LR_Model,Test_Data[,1:13])

summary(LR_Predict)

Per_LR=print(postResample(pred = LR_Predict,obs = Test_Data[,14]))


#RMSE=0.49
#Rsq=0.43
#MAE=0.38

####################################################################################################################

#################################################------------MODEL AFTER USING DUMMIEs---------------###################

#After Using Dummies
library(dummies)
data=df
data = dummy.data.frame(data, cat_data)



S=createDataPartition(data$Absenteeism.time.in.hours,p=0.70,list = FALSE)
TD=data[S,]
TED=data[-S,]

#####################################################################################################################


#############################---------------DESISION TREE (with Dummies)--------------------#########################################


Decision_Tree1=rpart(Absenteeism.time.in.hours~.,data=TD,method='anova')

summary(Decision_Tree)

#Now lets preidct Test Cases
Predict_DT1=predict(Decision_Tree1,TED[,-51])

#Predict_DT
summary(Predict_DT1)

plot(TED$Absenteeism.time.in.hours,type="l",lty=1.8,col="Green",main="DT")
lines(Predict_DT1,type="l",col="Blue")

##oR
library('DMwR')

Per_DT1=print(postResample(obs = TED[,51],pred = Predict_DT1))
#RMSE=0.47
#rsq=0.48
#MAE=0.36

######################################################################################################################

#############################----------------RANDOM FOREST (with Dummies)--------------------#########################################

Rf_Model1=randomForest(Absenteeism.time.in.hours~.,TD,ntree=300)

plot(Rf_Model1)

summary(Rf_Model1)

Rf_Predict1=predict(Rf_Model1,TED[,-51])

summary(Rf_Predict1)


Per_RF1=print(postResample(pred = Rf_Predict1,obs = TED[,51]))

#RMSE=0.44
#Rsq=0.53
#MAE=0.34
#####################################################################################################################

#############################---------------LINEAR REGRESSION (with Dummies)--------------------#########################################


LR_Model1=lm(Absenteeism.time.in.hours~.,TD)

summary(LR_Model1)

#Predict Test cases by ignoring TV

LR_Predict1=predict(LR_Model1,TED[,-51])

summary(LR_Predict1)

Per_LR1=print(postResample(pred = LR_Predict1,obs = TED[,51]))

#RMSE=0.47
#Rsq=0.48
#MAE=0.37


###################################################################################################################

######################################----------Performance Comparison-------------#################################

#Per_DT1,RF1,LR1 is peformance when applied dummies.
#Per_DT,RF,LR is performance of Model without Dummies

CV=as.data.frame(Per_DT)
CV1=as.data.frame(Per_DT1)
Cv2=as.data.frame(Per_RF)
CV3=as.data.frame(Per_RF1)
Cv4=as.data.frame(Per_LR)
Cv5=as.data.frame(Per_LR1)
FP=cbind(CV,CV1,Cv2,CV3,Cv4,Cv5)

write.csv(FP,file = 'Error Metrics.csv')
#As we can see Random Forest performed well.

#Dummies helped to reduce error percentage.


###########################--------------------------------------#######################################################


#Interacitve Way to Plot Graph

#install.packages('esquisse')
library('esquisse')

#esquisse::esquisser()            
#library(ggplot2)

#ggplot(data = df) +
# aes(x = Absenteeism.time.in.hours, y = Reason.for.absence, color = Son) +
#geom_point() +
#scale_colour_viridis_c(option  = "viridis") +
#theme_economist_white()


#With this library you can easily plot graph just by a drag & dorp feature.
#####################################################################################################################