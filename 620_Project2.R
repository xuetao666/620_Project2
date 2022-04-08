##########Clean up and import package:

rm(list=ls(all=TRUE))  #same to clear all in stata
cat("\014")

x<-c("dplyr","ggplot2","tidyr","xlsx","stringr","stringi")

new.packages<-x[!(x %in% installed.packages()[,"Package"])]
if(length(new.packages)) install.packages(new.packages)


lapply(x, require, character.only=T)
coalesce <- function(...) {
  apply(cbind(...), 1, function(x) {
    x[which(!is.na(x))[1]]
  })
}
##########################################
#Functions:

#Unadjusted:

getSLR=function(outcome,contvarlist,catvarlist,data){
  coef=c()
  for(var in contvarlist){
    fit=lm(as.formula(paste0(outcome,"~",var)),data=data)
    conf=round(confint(fit),2)[2,]
    summary_fit=round(summary(fit)$coef[2,],3)
    p=ifelse(summary_fit[4]<=0.001,"<0.001",ifelse(summary_fit[4]<0.05,paste0(summary_fit[4],"*"),summary_fit[4]))
    conf=paste0(round(summary_fit[1],2),"(",conf[1],",",conf[2],")")
    coef=rbind(coef,c(var,"",conf,p))
  }
  
  for(var in catvarlist){
    fit=lm(as.formula(paste0(outcome,"~as.factor(",var,")")),data=data)
    summary_fit=round(summary(fit)$coef[-1,],digits = 3)
    names= rownames(summary(fit)$coef)[-1]
    ##extract names
    names=gsub(paste0("as.factor\\(",var,"\\)"),"",names)
    ref_level=unique(data[[var]])[!(unique(data[[var]]) %in% names)]
    conf=round(confint(fit)[-1,],2)
    
    ref=c(var,ref_level,"Reference","-")
    if(length(nrow(summary_fit))==0){ #One line
      p=ifelse(summary_fit[4]<=0.001,"<0.001",ifelse(summary_fit[4]<0.05,paste0(summary_fit[4],"*"),summary_fit[4]))
      coef=rbind(coef,ref,c("",names,paste0(round(summary_fit[1],2),"(",conf[1],",",conf[2],")"),p))
    } else{
      p=ifelse(summary_fit[,4]<=0.001,"<0.001",ifelse(summary_fit[,4]<0.05,paste0(summary_fit[,4],"*"),summary_fit[,4]))
      conf=paste0(summary_fit[,1],"(",conf[,1],",",conf[,2],")")
      coef=rbind(coef,ref,cbind("",names,conf,p))
    }
  }
  colnames(coef)=c("Factor","Level","Coefficient(95%CI)","P value")
  row.names(coef)=NULL
  return(coef)
  
}

getMLR=function(fit,data,digit){
  tbl=c()
  
  #get intercept:
  intercept=summary(fit)$coef[1,]
  p=round(intercept[length(intercept)],3)
  intercept=round(intercept,digits = digit)
  if(p==0){
    p="<0.001**"
  }
  ci=paste0(round(confint(fit)[1,],digits=digit),collapse = ",")
  intercept=c(paste0(intercept[1],"(",ci,")"),p)
  tbl=c("Intercept","",intercept)
  
  
  
  varlist=names(fit$model)[-1]
  lm_name=rownames(summary(fit)$coef)
  
  ##Add the variables:
  for(var in varlist){
    #Continous
    if(!grepl("as.factor",var)){
      coef=summary(fit)$coef[var,]
      p=round(coef[length(coef)],3)
      coef=round(coef,digits = digit)
      if(p==0){
        p="<0.001**"
      } 
      if(p<=0.05){
        p=paste0(p,"*")
      }
      ci=paste0(round(confint(fit)[var,],digits=digit),collapse = ",")
      
      coef=c(paste0(coef[1],"(",ci,")"),p)
      ##Add 95%CI
      tbl=rbind(tbl,c(var,"",coef))
    }
    else { #Categorical variables:
      
      #Overall levels of this category:
      
      
      catvarlist=lm_name[stri_detect_fixed(lm_name,var)]
      var_clean=str_remove_all(var,"as.factor\\(|\\)")
      level_cat=unique(data[[var_clean]])
      compare_group=str_remove_all(catvarlist,paste0("as.factor\\(",var_clean,"\\)"))
      ref_group=level_cat[!(level_cat %in% compare_group)]
      
      #Reference group:
      ref=c(var_clean,ref_group,"Reference","-")
      for(cat in compare_group){
        catvar=paste0(var,cat)
        coef=summary(fit)$coef[catvar,]
        p=round(coef[length(coef)],3)
        coef=round(coef,digits = digit)
        if(p==0){
          p="<0.001**"
        }
        summary_fit=round(summary(fit)$coef[catvar,],digits = digit)
        p=ifelse(p<=0.001,"<0.001**",ifelse(p<0.05,paste0(p,"*"),p))
        ci=paste0(round(confint(fit)[catvar,],digits=digit),collapse = ",")
        tbl=rbind(tbl,ref,c("",cat,paste0(round(summary_fit[1],2),"(",ci,")"),p))
      }
    }
  }
  
  row.names(tbl)=NULL
  colnames(tbl)=c("Factor","Level","Coefficient(95%CI)","P value")
  return(tbl)
}


#########################################

setwd("C:/Users/xueti/Dropbox (University of Michigan)/Umich/class/term2/620/Group Project/620_Project2/Data/")


data_screen=read.xlsx("620W22-Project2-Data.xlsx",sheetName = "ScreenTime")
data_baseline=read.xlsx("620W22-Project2-Data.xlsx",sheetName ="Baseline")


##Data cleaning:
#Exclude missing and impute variables:
#Flip the total screen time and social screen time:

pickup.1st.min=str_split(data_screen$Pickup.1st,pattern = ":")

pickup.1st.min2=unlist(lapply(pickup.1st.min, function(x){as.numeric(x[1])*60+as.numeric(x[2])}))
data_screen$Pickup.1st.min=c(pickup.1st.min2[-1],NA)
data_screen=data_screen %>%
  mutate(Pickup.1st.min=ifelse(Time==30,NA,Pickup.1st.min)) %>%
  filter(!is.na(Pickup.1st.min))


data_screen=data_screen %>%
  filter(Pickup.1st!="NA" & Day!="NA" & Tot.Scr.Time !="NA" & Tot.Soc.Time!="NA" & Pickups!="NA") %>%
  mutate(Tot.Scr.Time=as.numeric(Tot.Scr.Time),
         Tot.Soc.Time=as.numeric(Tot.Soc.Time)) %>%
  mutate(Tot.Scr.Time2=ifelse(Tot.Scr.Time<Tot.Soc.Time,Tot.Soc.Time,Tot.Scr.Time),
         Tot.Soc.Time2=ifelse(Tot.Scr.Time<Tot.Soc.Time,Tot.Scr.Time,Tot.Soc.Time)) %>%
  mutate(Tot.Scr.Time=Tot.Scr.Time2,
         Tot.Soc.Time=Tot.Soc.Time2) %>%
  mutate(Phase=ifelse(Time<15,"Baseline",ifelse(Time>=15 & Time<=22,"A","B")),
         InterventionA=ifelse(Time<15,0,ifelse(Time>=15 & Time<=22,1,NA)),
         InterventionB=ifelse(Time<15,0,ifelse(Time>22,1,NA))) %>%
  select(!Tot.Scr.Time2) %>%
  select(!Tot.Soc.Time2) %>%
  mutate(Time=as.numeric(Time)) %>%
  mutate(Day=ifelse(Day=="Mon","Monday",
                    ifelse(Day=="Tue","Tuesday",
                           ifelse(Day=="Wed","Wednesday",
                                  ifelse(Day=="Thu","Thursday",
                                         ifelse(Day=="Fri","Friday",
                                                ifelse(Day=="Sat","Saturday",
                                                       ifelse(Day=="Sun","Sunday",Day))))))))


data_baseline=data_baseline %>%
  mutate(age=ifelse(age=="NA",NA,age)) %>%
  mutate(age=as.numeric(age)) %>%
  mutate(age=ifelse(is.na(age),mean(age,na.rm = T),age))

#Export in Excel

setwd("C:/Users/xueti/Dropbox (University of Michigan)/Umich/class/term2/620/Group Project/620_Project2/Results/")
date<-gsub("-","_",Sys.Date())
dir.create(date)
setwd(date)
resultpath<-getwd()

wb<-createWorkbook(type="xlsx")
TITLE_STYLE <- CellStyle(wb)+ Font(wb,  heightInPoints=16,  isBold=TRUE, underline=1)
SUB_TITLE_STYLE <- CellStyle(wb) +
  Font(wb,  heightInPoints=14,
       isItalic=TRUE, isBold=FALSE)
TEXT_STYLE <- CellStyle(wb) +
  Font(wb,  heightInPoints=12,
       isItalic=FALSE, isBold=FALSE)

TABLE_ROWNAMES_STYLE <- CellStyle(wb) + Font(wb, isBold=TRUE)
TABLE_COLNAMES_STYLE <- CellStyle(wb) + Font(wb, isBold=TRUE) +
  Alignment(wrapText=TRUE, horizontal="ALIGN_CENTER") +
  Border(color="black", position=c("TOP", "BOTTOM"),
         pen=c("BORDER_THIN", "BORDER_THIN"))

#Analysis 2:
sheet <- createSheet(wb, sheetName = "Table3")
xlsx.addTitle<-function(sheet, rowIndex, title, titleStyle,colIndex=1){
  rows <-createRow(sheet,rowIndex=rowIndex)
  sheetTitle <-createCell(rows, colIndex=colIndex)
  setCellValue(sheetTitle[[1,1]], title)
  setCellStyle(sheetTitle[[1,1]], titleStyle)
}


#Table 3a:
data_3=data_screen %>%
  group_by(ID,Phase) %>%
  summarize(WakeUp_mean=mean(Pickup.1st.min)) %>%
  spread(key=Phase,value=WakeUp_mean) %>%
  mutate(InterventionA_diff=A-Baseline,
         InterventionB_diff=B-Baseline) %>%
  left_join(data_baseline,by="ID")

data_3=data_3[,1:19]

data_4=data_screen %>%
  mutate(schoolday=ifelse(Day=="Monday" | Day=="Tuesday"| Day=="Wednesday" | Day=="Thursday",1,0)) %>%
  group_by(ID,Phase,schoolday) %>%
  summarize(WakeUp_mean=mean(Pickup.1st.min)) %>%
  spread(key=Phase,value=WakeUp_mean) %>%
  mutate(InterventionA_diff=A-Baseline,
         InterventionB_diff=B-Baseline) %>%
  left_join(data_baseline,by="ID")

data_4=data_4[,1:20]
data_4_schoolday=data_4 %>%
  filter(schoolday==1)
data_4_nonschoolday=data_4 %>%
  filter(schoolday==0)


tA_overall=t.test(data_3$A, data_3$Baseline, paired = TRUE, alternative = "two.sided")
tB_overall=t.test(data_3$B, data_3$Baseline, paired = TRUE, alternative = "two.sided") 
tA_school=t.test(data_4_schoolday$A, data_4_schoolday$Baseline, paired = TRUE, alternative = "two.sided")
tA_nonschool=t.test(data_4_nonschoolday$A, data_4_nonschoolday$Baseline, paired = TRUE, alternative = "two.sided")
tB_school=t.test(data_4_schoolday$B, data_4_schoolday$Baseline, paired = TRUE, alternative = "two.sided")
tB_nonschool=t.test(data_4_nonschoolday$B, data_4_nonschoolday$Baseline, paired = TRUE, alternative = "two.sided")


diff_mean=c(tA_overall$estimate,tA_school$estimate,tA_nonschool$estimate,
            tB_overall$estimate,tB_school$estimate,tB_nonschool$estimate)
diff_t=c(tA_overall$statistic,tA_school$statistic,tA_nonschool$statistic,
         tB_overall$statistic,tB_school$statistic,tB_nonschool$statistic)
diff_p=c(tA_overall$p.value,tA_school$p.value,tA_nonschool$p.value,
         tB_overall$p.value,tB_school$p.value,tB_nonschool$p.value)
diff_ul=c(tA_overall$conf.int[[1]],tA_school$conf.int[[1]],tA_nonschool$conf.int[[1]],
          tB_overall$conf.int[[1]],tB_school$conf.int[[1]],tB_nonschool$conf.int[[1]])
diff_ll=c(tA_overall$conf.int[[2]],tA_school$conf.int[[2]],tA_nonschool$conf.int[[2]],
          tB_overall$conf.int[[2]],tB_school$conf.int[[2]],tB_nonschool$conf.int[[2]])
ci=paste0(round(diff_ul,3),",",round(diff_ll,3))
type=c("Overall","School day","Non-School day","Overall","School day","Non-School day")
type2=c("Intervention A","","","Intervention B","","")
diff_tbl=cbind(type2,type,round(diff_mean,3),round(diff_t,3),ci,round(diff_p,3))
colnames(diff_tbl)=c(NA,NA,"Mean","T value","95% CI","P value")
diff_tbl=as.data.frame(diff_tbl)

xlsx.addTitle(sheet, rowIndex=1,
              title="Table3a Mean difference of Wake-up time before and after intervention",
              titleStyle = TEXT_STYLE)

addDataFrame(diff_tbl, sheet, startRow=3, startColumn=1,row.names  = F,
             colnamesStyle =TABLE_COLNAMES_STYLE)

rowindex=3+nrow(diff_tbl)+2

#Unadjusted analysis

contvarlist=c("workmate","academic","non.academic","age","course.hours","siblings","apps","devices","procrastination")
catvarlist=c("pets","sex","degree","job")
unadjusted_tblA=getSLR(outcome="InterventionA_diff",data=data_3,contvarlist = contvarlist,catvarlist = catvarlist)
unadjusted_tblAS=getSLR(outcome="InterventionA_diff",data=data_4_schoolday,contvarlist = contvarlist,catvarlist = catvarlist)
unadjusted_tblANS=getSLR(outcome="InterventionA_diff",data=data_4_nonschoolday,contvarlist = contvarlist,catvarlist = catvarlist)

unadjusted_tblB=getSLR(outcome="InterventionB_diff",data=data_3,contvarlist = contvarlist,catvarlist = catvarlist)
unadjusted_tblBS=getSLR(outcome="InterventionB_diff",data=data_4_schoolday,contvarlist = contvarlist,catvarlist = catvarlist)
unadjusted_tblBNS=getSLR(outcome="InterventionB_diff",data=data_4_nonschoolday,contvarlist = contvarlist,catvarlist = catvarlist)

unadjusted_tblA=cbind(unadjusted_tblA,unadjusted_tblAS[,3:4],unadjusted_tblANS[,3:4])
unadjusted_tblA=rbind(c("","","Overall","","School day","","Non-School day",""),colnames(unadjusted_tblA),unadjusted_tblA)

unadjusted_tblB=cbind(unadjusted_tblB,unadjusted_tblBS[,3:4],unadjusted_tblBNS[,3:4])
unadjusted_tblB=rbind(c("","","Overall","","School day","","Non-School day",""),colnames(unadjusted_tblB),unadjusted_tblB)


xlsx.addTitle(sheet, rowIndex=rowindex,
              title="Table3b Factors crude influence on the mean difference of Wake-up time before and after intervention A",
              titleStyle = TEXT_STYLE)
addDataFrame(unadjusted_tblA, sheet, startRow=rowindex+2, startColumn=1,row.names = F,
             col.names = F)
rowindex=rowindex+4+nrow(unadjusted_tblA)


xlsx.addTitle(sheet, rowIndex=rowindex,
              title="Table3c Factors crude influence on the mean difference of Wake-up time before and after intervention B",
              titleStyle = TEXT_STYLE)
addDataFrame(unadjusted_tblB, sheet, startRow=rowindex+2, startColumn=1,row.names = F,
             col.names = F)
rowindex=rowindex+4+nrow(unadjusted_tblB)



##Adjusted model:
formu=as.formula(paste0("InterventionA_diff","~",paste0(contvarlist,collapse = " + ")," + ",paste0("as.factor(",catvarlist,")",collapse = " + ")))
fit=lm(formu,data=data_3)
adjusted_tblA=getMLR(fit,data=data_3,digit=2)
formu=as.formula(paste0("InterventionA_diff","~",paste0(contvarlist,collapse = " + ")," + ",paste0("as.factor(",catvarlist,")",collapse = " + ")))
fit=lm(formu,data=data_4_schoolday)
adjusted_tblAS=getMLR(fit,data=data_4_schoolday,digit=2)
formu=as.formula(paste0("InterventionA_diff","~",paste0(contvarlist,collapse = " + ")," + ",paste0("as.factor(",catvarlist,")",collapse = " + ")))
fit=lm(formu,data=data_4_nonschoolday)
adjusted_tblANS=getMLR(fit,data=data_4_nonschoolday,digit=2)
adjusted_tblA=cbind(adjusted_tblA,adjusted_tblAS[,3:4],adjusted_tblANS[,3:4])
adjusted_tblA=rbind(c("","","Overall","","School day","","Non-School day",""),colnames(adjusted_tblA),adjusted_tblA)

xlsx.addTitle(sheet, rowIndex=rowindex,
              title="Table3d Factors adjusted influence on the mean difference of Wake-up time before and after intervention A",
              titleStyle = TEXT_STYLE)
addDataFrame(adjusted_tblA, sheet, startRow=rowindex+2, startColumn=1,row.names = F,
             col.names = F)
rowindex=rowindex+4+nrow(adjusted_tblA)


formu=as.formula(paste0("InterventionB_diff","~",paste0(contvarlist,collapse = " + ")," + ",paste0("as.factor(",catvarlist,")",collapse = " + ")))
fit=lm(formu,data=data_3)
adjusted_tblB=getMLR(fit,data=data_3,digit=2)
formu=as.formula(paste0("InterventionB_diff","~",paste0(contvarlist,collapse = " + ")," + ",paste0("as.factor(",catvarlist,")",collapse = " + ")))
fit=lm(formu,data=data_4_schoolday)
adjusted_tblBS=getMLR(fit,data=data_4_schoolday,digit=2)
formu=as.formula(paste0("InterventionB_diff","~",paste0(contvarlist,collapse = " + ")," + ",paste0("as.factor(",catvarlist,")",collapse = " + ")))
fit=lm(formu,data=data_4_nonschoolday)
adjusted_tblBNS=getMLR(fit,data=data_4_nonschoolday,digit=2)
adjusted_tblB=cbind(adjusted_tblB,adjusted_tblBS[,3:4],adjusted_tblBNS[,3:4])
adjusted_tblB=rbind(c("","","Overall","","School day","","Non-School day",""),colnames(adjusted_tblB),adjusted_tblB)

xlsx.addTitle(sheet, rowIndex=rowindex,
              title="Table3e Factors adjusted influence on the mean difference of Wake-up time before and after intervention B",
              titleStyle = TEXT_STYLE)
addDataFrame(adjusted_tblB, sheet, startRow=rowindex+2, startColumn=1,row.names = F,
             col.names = F)
rowindex=rowindex+4+nrow(adjusted_tblB)




# 
# 
# formu=as.formula(paste0("InterventionB_diff","~",paste0(contvarlist,collapse = " + ")," + ",paste0("as.factor(",catvarlist,")",collapse = " + ")))
# fit=lm(formu,data=data_3)
# adjusted_tblB=getMLR(fit,data=data_3,digit=2)
# 
# adjusted_tbl=cbind(adjusted_tblA,adjusted_tblB[,3:4])
# 
# 
# xlsx.addTitle(sheet, rowIndex=rowindex,
#               title="Table3c Factors adjusted influence on the mean difference of Wake-up time before and after intervention",
#               titleStyle = TEXT_STYLE)
# 
# name_col=as.data.frame(t(c("","","Intervention A","","Intervention B")))
# 
# addDataFrame(name_col, sheet, startRow=rowindex+1, startColumn=1,row.names = F,col.names = F)
# 
# 
# addDataFrame(adjusted_tbl, sheet, startRow=rowindex+2, startColumn=1,row.names = F,
#              colnamesStyle =TABLE_COLNAMES_STYLE)
# 
# #Table 4:
# #Sheet format
# sheet <- createSheet(wb, sheetName = "Table4")
# 
# xlsx.addTitle<-function(sheet, rowIndex, title, titleStyle,colIndex=1){
#   rows <-createRow(sheet,rowIndex=rowIndex)
#   sheetTitle <-createCell(rows, colIndex=colIndex)
#   setCellValue(sheetTitle[[1,1]], title)
#   setCellStyle(sheetTitle[[1,1]], titleStyle)
# }
# 
# #t-test:
# 
# 
# xlsx.addTitle(sheet, rowIndex=1,
#               title="Table4a Mean difference of Wake-up time before and after intervention",
#               titleStyle = TEXT_STYLE)
# 
# 
# addDataFrame(diff_tbl, sheet, startRow=3, startColumn=1,
#              rownamesStyle = TABLE_ROWNAMES_STYLE,colnamesStyle =TABLE_COLNAMES_STYLE)
# 
# rowindex=3+nrow(diff_tbl)+2
# 
# #Undajusted
# 
# ###School day:
# contvarlist=c("workmate","academic","non.academic","age","course.hours","siblings","apps","devices","procrastination")
# catvarlist=c("pets","sex","degree","job")
# unadjusted_tblA=getSLR(outcome="InterventionA_diff",data=data_4_schoolday,contvarlist = contvarlist,catvarlist = catvarlist)
# unadjusted_tblB=getSLR(outcome="InterventionB_diff",data=data_4_schoolday,contvarlist = contvarlist,catvarlist = catvarlist)
# 
# unadjusted_tbl=cbind(unadjusted_tblA,unadjusted_tblB[,3:4])
# 
# xlsx.addTitle(sheet, rowIndex=rowindex,
#               title="Table4b Factors crude influence on the mean difference of Wake-up time before and after intervention among school days",
#               titleStyle = TEXT_STYLE)
# 
# name_col=as.data.frame(t(c("","","Intervention A","","Intervention B")))
# addDataFrame(name_col, sheet, startRow=rowindex+1, startColumn=1,row.names = F,col.names = F)
# 
# addDataFrame(unadjusted_tbl, sheet, startRow=rowindex+2, startColumn=1,row.names = F,
#              colnamesStyle =TABLE_COLNAMES_STYLE)
# 
# rowindex=rowindex+4+nrow(unadjusted_tbl)
# 
# ##Adjusted model:
# formu=as.formula(paste0("InterventionA_diff","~",paste0(contvarlist,collapse = " + ")," + ",paste0("as.factor(",catvarlist,")",collapse = " + ")))
# fit=lm(formu,data=data_3)
# adjusted_tblA=getMLR(fit,data=data_3,digit=2)
# 
# formu=as.formula(paste0("InterventionB_diff","~",paste0(contvarlist,collapse = " + ")," + ",paste0("as.factor(",catvarlist,")",collapse = " + ")))
# fit=lm(formu,data=data_3)
# adjusted_tblB=getMLR(fit,data=data_3,digit=2)
# 
# adjusted_tbl=cbind(adjusted_tblA,adjusted_tblB[,3:4])
# 
# 
# xlsx.addTitle(sheet, rowIndex=rowindex,
#               title="Table3c Factors adjusted influence on the mean difference of Wake-up time before and after intervention",
#               titleStyle = TEXT_STYLE)
# 
# name_col=as.data.frame(t(c("","","Intervention A","","Intervention B")))
# 
# addDataFrame(name_col, sheet, startRow=rowindex+1, startColumn=1,row.names = F,col.names = F)
# 
# 
# addDataFrame(adjusted_tbl, sheet, startRow=rowindex+2, startColumn=1,row.names = F,
#              colnamesStyle =TABLE_COLNAMES_STYLE)
# 
# 
# #Adjusted
# 
# 
# 
# 



saveWorkbook(wb, paste("620_Project2",date,".xlsx",sep=""),)








