getwd() 
workL <- "/Users/sunxy/Desktop"
setwd(workL) 
library(readxl)
#data <- read.table("name.txt",header = T,sep = "")
sfn_total <- read_excel("肝穿小纤维数据-总0418.xlsx",col_names = T,col_types = NULL ,na="", skip=0)
colnames(sfn_total)
dim(sfn_total)
str(sfn_total2)
sfn_total2<-sfn_total
sfn_total2<-as.data.frame(sfn_total2)  #不然不能用asfactor函数

asfactor<-function(a,b){         #transfer v as factor
  num<-which(colnames(a)==b)
  #print(class(a))
  #print(num)
  a[,num]<-as.factor(a[,num])
  #print(class(a[,num]))
  return(a)
}
#sfn_total2$振动觉诊断结果二分类<-as.factor(sfn_total2$振动觉诊断结果二分类)  #ok

listf<-colnames(sfn_total2)[c(4,6,7,37:41)] # add need to be factor
for (i in c(1:length(listf))){
  sfn_total2<-asfactor(sfn_total2,listf[i])
}

sfn_withresult<-subset(sfn_total2,小纤维检查数据==1)
table(sfn_withresult$目前有无病理)
# 20人有病理，5人无病理
sfn_withresult[which(sfn_withresult[,6]==0),c(2,3)]
# 5人里 杨宝华和唐丽君 胆囊手术-有标本，未送检  #但杨宝华MT且结果异常
sfn_withbiopsy<-subset(sfn_withresult,目前有无病理==1)
sfn_withbiopsy2<-rbind(sfn_withbiopsy,sfn_withresult[which(sfn_withresult[,2]=="唐丽君"),])

database_s <- read_excel("/Users/sunxy/Desktop/20190410/肝穿数据库整理201909/383肝穿胆囊简易数据库200401.xlsx",col_names = T,col_types = NULL ,na="", skip=0)
colnames(database_s)
database_s2<-as.data.frame(database_s)
dim(database_s2)

total_sfn_withbiopsy <- merge(sfn_withbiopsy2,database_s2,by.database_s2 = "姓名",by.sfn_withbiopsy2 = "姓名",all.x=TRUE)  #前一个小名单，后一个大名单，all.x=true 保留第一个矩阵的结构
colnames(total_sfn_withbiopsy)
dim(total_sfn_withbiopsy)
#str(total_sfn_withbiopsy2)
total_sfn_withbiopsy[which(is.na(total_sfn_withbiopsy$`NAS（total）`)==TRUE),2] #看有NA的是谁!!!!!!!!
total_sfn_withbiopsy2<-total_sfn_withbiopsy
total_sfn_withbiopsy2[which(is.na(total_sfn_withbiopsy$`NAS（total）`)==TRUE),c(66,67,71:80)]<-0
total_sfn_withbiopsy2$nas4group <- ifelse(total_sfn_withbiopsy2$`NAS病理分组（0正常1NAFL2NASH）` ==0, 0, ifelse(total_sfn_withbiopsy2$`NAS病理分组（0正常1NAFL2NASH）` >0 & total_sfn_withbiopsy2$`NAS（total）` > 4, 2, 1))

total_sfn_withbiopsy2$冷敏度阈值大鱼际<-total_sfn_withbiopsy2$冷阈值均数大鱼际-total_sfn_withbiopsy2$冷痛阈值均数大鱼际
total_sfn_withbiopsy2$热敏度阈值大鱼际<-total_sfn_withbiopsy2$热痛阈值均数大鱼际-total_sfn_withbiopsy2$热阈值均数大鱼际
total_sfn_withbiopsy2$冷敏度阈值足背<-total_sfn_withbiopsy2$冷阈值均数足背-total_sfn_withbiopsy2$冷痛阈值均数足背
total_sfn_withbiopsy2$热敏度阈值足背<-total_sfn_withbiopsy2$热痛阈值均数足背-total_sfn_withbiopsy2$热阈值均数足背

colnames(total_sfn_withbiopsy2)
listf2<-colnames(total_sfn_withbiopsy2)[c(43,50,51,66,77,78,80,85)] # add need to be factor
for (i in c(1:length(listf2))){
  total_sfn_withbiopsy2<-asfactor(total_sfn_withbiopsy2,listf2[i])
}
str(total_sfn_withbiopsy2)
dim(total_sfn_withbiopsy2)

#total_sfn_withbiopsy2$"1NGT 2IGR 3T2DM"
#total_sfn_withbiopsy2[which(is.na(total_sfn_withbiopsy2$"1NGT 2IGR 3T2DM")==TRUE),2] #看有NA的是谁!!!!!!!!
#total_sfn_withbiopsy2$"1NGT 2IGR 3T2DM"[11]<-1   #唐丽君NGT

total_sfn_withbiopsy_nodm<-subset(total_sfn_withbiopsy2,total_sfn_withbiopsy2$"1NGT 2IGR 3T2DM"!=3)
#table(total_sfn_withbiopsy2$"1NGT 2IGR 3T2DM")
#table(subset(total_sfn_withbiopsy_nodm,nas4group!=0)$"fibrosis")
total_sfn_withbiopsy_nodm_nafld<-subset(total_sfn_withbiopsy_nodm,nas4group!=0)
colnames(total_sfn_withbiopsy2)[74]<-"nas_total"

match<-total_sfn_withbiopsy2[,c(2,43,44,45,50,52,53,74,75)]
match
write.xlsx(match,"match.xlsx")

total_sfn_withbiopsy2$matchgroup<-NA
total_sfn_withbiopsy2$matchgroup<-ifelse(total_sfn_withbiopsy2$姓名%in% c("徐海燕","马斌","潘尚特","唐丽君","林晨昱","黄春凤","陆磊")==TRUE, '1', ifelse(total_sfn_withbiopsy2$姓名%in% c("王梦蕾","陈惠斌","郑玉林","陈丽莺","顾金鹏","焦雪英","周启航")==TRUE, '2',NA))
total_sfn_withbiopsy2$matchpair<-NA
#total_sfn_withbiopsy2$matchpair[which(total_sfn_withbiopsy2$姓名=="周启航")]<-NA
total_sfn_withbiopsy2$matchpair[which(total_sfn_withbiopsy2$姓名=="徐海燕"|total_sfn_withbiopsy2$姓名=="王梦蕾")]<-"徐海燕-王梦蕾"
total_sfn_withbiopsy2$matchpair[which(total_sfn_withbiopsy2$姓名=="马斌"|total_sfn_withbiopsy2$姓名=="陈惠斌")]<-"马斌-陈惠斌"
total_sfn_withbiopsy2$matchpair[which(total_sfn_withbiopsy2$姓名=="潘尚特"|total_sfn_withbiopsy2$姓名=="郑玉林")]<-"潘尚特-郑玉林"
total_sfn_withbiopsy2$matchpair[which(total_sfn_withbiopsy2$姓名=="唐丽君"|total_sfn_withbiopsy2$姓名=="陈丽莺")]<-"唐丽君-陈丽莺"
total_sfn_withbiopsy2$matchpair[which(total_sfn_withbiopsy2$姓名=="林晨昱"|total_sfn_withbiopsy2$姓名=="顾金鹏")]<-"林晨昱-顾金鹏"
total_sfn_withbiopsy2$matchpair[which(total_sfn_withbiopsy2$姓名=="黄春凤"|total_sfn_withbiopsy2$姓名=="焦雪英")]<-"黄春凤-焦雪英"
#total_sfn_withbiopsy2$matchpair[which(total_sfn_withbiopsy2$姓名=="李俊毅"|total_sfn_withbiopsy2$姓名=="顾金鹏")]<-"李俊毅-顾金鹏"
total_sfn_withbiopsy2$matchpair[which(total_sfn_withbiopsy2$姓名=="陆磊"|total_sfn_withbiopsy2$姓名=="周启航")]<-"陆磊-周启航"


colnames(total_sfn_withbiopsy2)
#[71] "fat"                                                                                                              
#[72] "necrosisfoci"                                                                                                     
#[73] "ballooning"                                                                                                       
#[74] "NAS（total）"                                                                                                     
#[75] "fibrosis" 
#############
my_crosstab_plus<-function(dataname,vname,groupname){    #实现含sum的百分比显示
  crosstab<-table(dataname[,c(vname,groupname)]) 
  crosstab2<-addmargins(table(dataname[,c(vname,groupname)])) #分组变量写后面
  crosstab<-as.matrix(crosstab)
  chi2_p<-round(chisq.test(crosstab)$p.value,3)
  row=dim(crosstab2)[1]
  col=dim(crosstab2)[2]
  crosstab3<-crosstab2
  for (c in c(1:col)){
    for (r in c(1:row)){
      prop<-crosstab2[r,c]/crosstab2[row,c]
      x<-crosstab2[r,c]
      y<-paste(x,"(",round(prop*100,2), "%",")",sep='')
      crosstab3[r,c]<-y
      
    }
  }
  return(crosstab3)
}
my_crosstab_plus(total_sfn_withbiopsy2,"冷热阈值诊断结果二分类","nas4group")
my_crosstab_plus(total_sfn_withbiopsy2,"冷痛和热痛阈值诊断结果二分类","nas4group")
my_crosstab_plus(total_sfn_withbiopsy2,"振动觉诊断结果二分类","nas4group")
my_crosstab_plus(total_sfn_withbiopsy2,"SSR诊断结果二分类","nas4group")

#普通方差分析多重比较
boxplert <-  function(X,
                      Y,
                      main = NULL,
                      xlab = NULL,
                      ylab = NULL,
                      bcol = "bisque",
                      p.adj = "none",
                      cexy = 1,
                      varwidth = TRUE,
                      las = 1,
                      paired = FALSE)
{#Y=factor(Y,levels=unique(Y)) #如果运行次代码，表示组别是按照原来的顺序，而不是字母顺序排列
  aa <- levels(as.factor(Y))
  an <- as.character(c(1:length(aa)))
  tt1 <- matrix(nrow = length(aa), ncol = 6)    
  for (i in 1:length(aa))
  {
    temp <- X[Y == aa[i]]
    tt1[i, 1] <- mean(temp, na.rm = TRUE)
    tt1[i, 2] <- sd(temp, na.rm = TRUE) / sqrt(length(temp))
    tt1[i, 3] <- sd(temp, na.rm = TRUE)
    tt1[i, 4] <- min(temp, na.rm = TRUE)
    tt1[i, 5] <- max(temp, na.rm = TRUE)
    tt1[i, 6] <- length(temp)
  }
  
  tt1 <- as.data.frame(tt1)
  row.names(tt1) <- aa
  colnames(tt1) <- c("mean", "se", "sd", "min", "max", "n")
  
  boxplot(
    X ~ Y,
    main = main,
    xlab = xlab,
    ylab = ylab,
    las = las,
    col = bcol,
    cex.axis = cexy,
    cex.lab = cexy,
    varwidth = varwidth
  )    
  require(agricolae)
  Yn <- factor(Y, labels = an)
  sig <- "ns"
  model <- aov(X ~ Yn)    
  if (paired == TRUE & length(aa) == 2)
  {
    coms <- t.test(X ~ Yn, paired = TRUE)
    pp <- coms$p.value
  }    else
  {
    pp <- anova(model)$Pr[1]
  }    
  if (pp <= 0.1)
    sig <- "ns"
  if (pp <= 0.05)
    sig <- "*"
  if (pp <= 0.01)
    sig <- "**"
  if (pp <= 0.001)
    sig <- "***"
  
  mtext(
    sig,
    side = 3,
    line = 0.5,
    adj = 0,
    cex = 2,
    font = 1
  )    
  if (pp <= 0.05) {
    comp <- LSD.test(model,                       "Yn",
                     alpha = 0.05,
                     p.adj = p.adj,
                     group = TRUE)      # gror <- comp$groups[order(comp$groups$groups), ]
    # tt1$cld <- gror$M
    gror <- comp$groups[order(rownames(comp$groups)), ]
    tt1$group <- gror$groups
    mtext(
      tt1$group,
      side = 3,
      at = c(1:length(aa)),
      line = 0.5,
      cex = 1,
      font = 4
    )
  }
  list(comparison = tt1, p.value = pp)
}


colnames(total_sfn_withbiopsy2)[8:31]
#[1] "冷阈值均数大鱼际"         "冷阈值均数足背"           "冷痛阈值均数大鱼际"       "冷痛阈值均数足背"        
#[5] "热阈值均数大鱼际"         "热阈值均数足背"           "热痛阈值均数大鱼际"       "热痛阈值均数足背"        
#[9] "振动感觉-中指"            "振动感觉-踇趾"            "CHEP上肢Cz/N500"          "CHEP上肢Cz/N750"         
#[13] "CHEP上肢Pz/P1000"         "CHEP下肢CHEP上肢Cz/N500"  "CHEP下肢CHEP上肢Cz/N750"  "CHEP下肢CHEP上肢Pz/P1000"
#[17] "电刺激左足Lat"            "电刺激右足Lat"            "电刺激左手Lat"            "电刺激右手Lat"           
#[21] "电刺激左足Amp"            "电刺激右足Amp"            "电刺激左手Amp"            "电刺激右手Amp"   

boxplert(
  total_sfn_withbiopsy2$"振动感觉-踇趾",
  total_sfn_withbiopsy2$nas4group,
  #ylab = "冷阈值均数大鱼际",
  xlab = "0=control,1=nas<5,2=nas>=5",
  bcol = "grey",
  p.adj = "holm"
)

table(total_sfn_withbiopsy2[is.na(total_sfn_withbiopsy2$电刺激左足Lat)==FALSE,]$nas4group)  #得到每个指标每组有几个人做了
#boxplot(tapply(total_sfn_withbiopsy2$冷阈值均数大鱼际, total_sfn_withbiopsy2$nas4group, list), main = "a",varwidth = T,cex.main=2,cex.axis=0.6,cex.lab=1,
        #names=paste(c("NC,n=","NAFL,n=","NASH,n="),table(total_sfn_withbiopsy2[is.na(total_sfn_withbiopsy2$冷阈值均数大鱼际)==FALSE,]$nas4group)))

table(total_sfn_withbiopsy2$fibrosis)  #观察纤维化情况
table(total_sfn_withbiopsy2[which(total_sfn_withbiopsy2$nas4group!=0),]$fibrosis)
total_sfn_withbiopsy2$fibgroup <- ifelse(total_sfn_withbiopsy2$fibrosis<2, 0, 1)

total_sfn_withbiopsy2$fibgroup<-as.factor(total_sfn_withbiopsy2$fibgroup)
total_sfn_nafld<-subset(total_sfn_withbiopsy2,nas4group!=0)
boxplert(
  total_sfn_nafld$"热痛阈值均数大鱼际",
  total_sfn_nafld$fibgroup,
  #ylab = "冷阈值均数大鱼际",
  xlab = "0=NAFLD F0-1,1=NAFLD F2-3",
  bcol = "grey",
  p.adj = "holm"
)
table(total_sfn_nafld[is.na(total_sfn_nafld$"电刺激左足Lat")==FALSE,]$fibgroup)  #得到每个指标每组有几个人做了

library(xlsx)
write.xlsx(total_sfn_withbiopsy2,"SFN有病理19人数据0418.xlsx")

library(ggplot2)
#ggplot(total_sfn_withbiopsy2[which(is.na(total_sfn_withbiopsy2$matchgroup)==FALSE),],aes(matchgroup,冷阈值均数大鱼际,color=matchpair))+geom_point(size=3)+geom_smooth(method = lm)+theme_grey(base_family = "STHeiti")
ggplot(total_sfn_withbiopsy2,aes(as.numeric(matchgroup),振动感觉踇趾,color=matchpair))+geom_point(size=3)+geom_line()+theme_grey(base_family = "STHeiti")+ scale_x_continuous("正常/轻度NAFLD vs 重度NAFLD", breaks=NULL)+theme(legend.title=element_blank())
#######主要产图
#geom_text(aes(label=姓名), vjust=1.5, colour="black")+

#plot(c(1,2,3),c(2,3,4),xlab="啊")
par(family='STKaiti')  #使plot里显示中文

table(total_sfn_withbiopsy2[which(total_sfn_withbiopsy2$nas4group!=0),]$"NAS（total）")  #观察nas情况


colnames(total_sfn_withbiopsy2)[16]<-"振动感觉中指"                                                                                                    
colnames(total_sfn_withbiopsy2)[17]<-"振动感觉踇趾" 

write_meansd<-function(vname,n,dataname){ 
  i<-which(colnames(dataname)==vname)
  if (n==1){
    v<-c(vname,paste(round(mean(dataname[,i], na.rm = TRUE),2),"±",round(sd(dataname[,i], na.rm = TRUE),2),sep=""))
  }else{
    v<-c(vname,paste(round(quantile(dataname[,i], na.rm = TRUE)[3],2),"(",round(quantile(dataname[,i], na.rm = TRUE)[2],2),"-",round(quantile(dataname[,i], na.rm = TRUE)[4],2),")",sep=""))
  }  
  return(v)
}
write_meansd("年龄",1,mydata4_nodrink)
write_v_table1_anova<-function(vname,groupname,n,groupcount,data){
  library(agricolae)
  i<-which(colnames(data)==vname)
  j<-which(colnames(data)==groupname)
  group<-data[,which(colnames(data)==groupname)]  #####新加 把分组变量填在这里
  model <- aov(data[,i] ~ group, data)
  p.anova<-summary(model)[[1]][1,5]  # anova p
  
  out <- LSD.test(model, "group", p.adj = "bonferroni" ) # multiple comparision
  #out$groups$groups    #group difference a b c ab
  v=c(vname)
  for (gn in c(1:groupcount)){
    if (n==1){
      v<-append(v,paste(round(out$means[gn,1],2),"±",round(out$means[gn,2],2)))
      
    }else{
      v<-append(v,paste(round(out$means[gn,9],2),"(",round(out$means[gn,8],2),"-",round(out$means[gn,10],2),")"))
    }  
    v<-append(v,paste(out$groups$groups[which(out$groups[,1]==out$means[gn,1])]))   #group difference a b c ab
  }
  v<-append(v,round(p.anova,3))    # anova p
  v<-append(write_meansd(vname,n,data),v)  #total count
  return(v)
}
# example
write_v_table1_anova("冷阈值均数大鱼际","fibrosis",1,4,total_sfn_withbiopsy2[which(total_sfn_withbiopsy2$nas4group!=0),])  #vname, #groupname  #1=mean sd,2=(25~75)   #how many groups   #dataframe
# rbind(write_v_table1("bmi2",1,3,data3),write_v_table1("age",1,3,data3))

table1<-write_v_table1_anova("冷阈值均数大鱼际","fibrosis",1,4,total_sfn_withbiopsy2[which(total_sfn_withbiopsy2$nas4group!=0),])  #vname, #groupname  #1=mean sd,2=(25~75)   #how many groups   #dataframe
table1<-rbind(table1,write_v_table1_anova("冷阈值均数足背","fibrosis",1,4,total_sfn_withbiopsy2[which(total_sfn_withbiopsy2$nas4group!=0),]))
table1<-rbind(table1,write_v_table1_anova("冷痛阈值均数大鱼际","fibrosis",2,4,total_sfn_withbiopsy2[which(total_sfn_withbiopsy2$nas4group!=0),]))
table1<-rbind(table1,write_v_table1_anova("冷痛阈值均数足背","fibrosis",2,4,total_sfn_withbiopsy2[which(total_sfn_withbiopsy2$nas4group!=0),]))
table1<-rbind(table1,write_v_table1_anova("热阈值均数大鱼际","fibrosis",1,4,total_sfn_withbiopsy2[which(total_sfn_withbiopsy2$nas4group!=0),]))
table1<-rbind(table1,write_v_table1_anova("热阈值均数足背","fibrosis",1,4,total_sfn_withbiopsy2[which(total_sfn_withbiopsy2$nas4group!=0),]))
table1<-rbind(table1,write_v_table1_anova("热痛阈值均数大鱼际","fibrosis",1,4,total_sfn_withbiopsy2[which(total_sfn_withbiopsy2$nas4group!=0),]))
table1<-rbind(table1,write_v_table1_anova("热痛阈值均数足背","fibrosis",1,4,total_sfn_withbiopsy2[which(total_sfn_withbiopsy2$nas4group!=0),]))
table1<-rbind(table1,write_v_table1_anova("振动感觉中指","fibrosis",1,4,total_sfn_withbiopsy2[which(total_sfn_withbiopsy2$nas4group!=0),]))
table1<-rbind(table1,write_v_table1_anova("振动感觉踇趾","fibrosis",1,4,total_sfn_withbiopsy2[which(total_sfn_withbiopsy2$nas4group!=0),]))
library(xlsx)
write.xlsx(table1,"table1temp.xlsx")
table12<-write_v_table1_anova("冷痛阈值均数大鱼际","fibrosis",2,4,total_sfn_withbiopsy2[which(total_sfn_withbiopsy2$nas4group!=0),])
table12<-rbind(table12,write_v_table1_anova("冷痛阈值均数足背","fibrosis",2,4,total_sfn_withbiopsy2[which(total_sfn_withbiopsy2$nas4group!=0),]))
table12<-rbind(table12,write_v_table1_anova("振动感觉踇趾","fibrosis",2,4,total_sfn_withbiopsy2[which(total_sfn_withbiopsy2$nas4group!=0),]))

write.xlsx(table12,"table1temp.xlsx",sheetName = "sheet2",append = TRUE)
