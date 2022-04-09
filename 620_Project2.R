##########Clean up and import package:

rm(list=ls(all=TRUE))  #same to clear all in stata
cat("\014")

x<-c("dplyr","ggplot2","tidyr","xlsx","stringr","stringi","MASS","reshape2","data.table","cowplot")

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
    coef=rbind(coef,c(var,"",conf,p,var))
  }
  
  for(var in catvarlist){
    fit=lm(as.formula(paste0(outcome,"~as.factor(",var,")")),data=data)
    summary_fit=round(summary(fit)$coef[-1,],digits = 3)
    names= rownames(summary(fit)$coef)[-1]
    ##extract names
    names=gsub(paste0("as.factor\\(",var,"\\)"),"",names)
    ref_level=unique(data[[var]])[!(unique(data[[var]]) %in% names)]
    conf=round(confint(fit)[-1,],2)
    
    ref=c(var,ref_level,"Reference","-",paste0(var,ref_level))
    if(length(nrow(summary_fit))==0){ #One line
      p=ifelse(summary_fit[4]<=0.001,"<0.001",ifelse(summary_fit[4]<0.05,paste0(summary_fit[4],"*"),summary_fit[4]))
      coef=rbind(coef,ref,c("",names,paste0(round(summary_fit[1],2),"(",conf[1],",",conf[2],")"),p,paste0(var,names)))
    } else{
      p=ifelse(summary_fit[,4]<=0.001,"<0.001",ifelse(summary_fit[,4]<0.05,paste0(summary_fit[,4],"*"),summary_fit[,4]))
      conf=paste0(summary_fit[,1],"(",conf[,1],",",conf[,2],")")
      coef=rbind(coef,ref,cbind("",names,conf,p,paste0(var,names)))
    }
  }
  colnames(coef)=c("Factor","Level","Coefficient(95%CI)","P value","ID")
  row.names(coef)=NULL
  return(coef)
  
}

getMLR=function(fit,data,digit){
  tbl=c()
  sig_list=c()
  pos_list=c()
  neg_list=c()
  
  #get intercept:
  intercept=summary(fit)$coef[1,]
  p=round(intercept[length(intercept)],3)
  intercept=round(intercept,digits = digit)
  if(p==0){
    p="<0.001**"
  }
  if(p<=0.05){
    p=paste0(p,"*")
  }
  ci=paste0(round(confint(fit)[1,],digits=digit),collapse = ",")
  intercept=c(paste0(intercept[1],"(",ci,")"),p)
  tbl=c("Intercept","",intercept,"Intercept")
  
  
  
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
        sig_list=c(sig_list,var)
        if(coef[1]>0){
          pos_list=c(pos_list,var)
        }else{
          neg_list=c(neg_list,var)
        }
        
      } 
      if(p<=0.05){
        p=paste0(p,"*")
        sig_list=c(sig_list,var)
        if(coef[1]>0){
          pos_list=c(pos_list,var)
        }else{
          neg_list=c(neg_list,var)
        }
      }
      ci=paste0(round(confint(fit)[var,],digits=digit),collapse = ",")
      
      coef=c(paste0(coef[1],"(",ci,")"),p)
      ##Add 95%CI
      tbl=rbind(tbl,c(var,"",coef,var))
    }
    else { #Categorical variables:
      
      #Overall levels of this category:
      
      
      catvarlist=lm_name[stri_detect_fixed(lm_name,var)]
      var_clean=str_remove_all(var,"as.factor\\(|\\)")
      level_cat=unique(data[[var_clean]])
      compare_group=str_remove_all(catvarlist,paste0("as.factor\\(",var_clean,"\\)"))
      ref_group=level_cat[!(level_cat %in% compare_group)]
      
      #Reference group:
      ref=c(var_clean,ref_group,"Reference","-",paste0(var_clean,ref_group))
      for(cat in compare_group){
        catvar=paste0(var,cat)
        coef=summary(fit)$coef[catvar,]
        p=round(coef[length(coef)],3)
        coef=round(coef,digits = digit)
        if(p==0){
          p="<0.001**"
          sig_list=c(sig_list,var)
          if(coef[1]>0){
            pos_list=c(pos_list,var)
          }else{
            neg_list=c(neg_list,var)
          }
        }
        if(p<=0.05){
          p=paste0(p,"*")
          sig_list=c(sig_list,var)
          if(coef[1]>0){
            pos_list=c(pos_list,var)
          }else{
            neg_list=c(neg_list,var)
          }
        }
        summary_fit=round(summary(fit)$coef[catvar,],digits = digit)
        
        ci=paste0(round(confint(fit)[catvar,],digits=digit),collapse = ",")
        tbl=rbind(tbl,ref,c("",cat,paste0(round(summary_fit[1],2),"(",ci,")"),p,paste0(var_clean,cat)))
      }
    }
  }
  
  row.names(tbl)=NULL
  colnames(tbl)=c("Factor","Level","Coefficient(95%CI)","P value","ID")
  returnlist=list()
  returnlist[["tbl"]]=tbl
  returnlist[["sig_list"]]=sig_list
  returnlist[["pos_list"]]=pos_list
  returnlist[["neg_list"]]=neg_list

  return(returnlist)
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
  mutate(Pickup.1st.min=ifelse(Time==30,NA,Pickup.1st.min)) 


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
  dplyr::select(!Tot.Scr.Time2) %>%
  dplyr::select(!Tot.Soc.Time2) %>%
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
data_screen=data_screen %>%
  mutate(schoolday="Overall")


data_3=data_screen %>%
  mutate(schoolday=ifelse(Day=="Monday" | Day=="Tuesday"| Day=="Wednesday" | Day=="Thursday","School Day","Non-School Day")) 
data_31=data_screen %>%
  mutate(schoolday=Day)
data_3=rbind(data_3,data_screen,data_31)


data_31=data_3 %>%
  group_by(ID,Phase,schoolday) %>%
  summarize(Tot.Scr.Time_mean=mean(Tot.Scr.Time,na.rm=T)) %>%
  spread(key=Phase,value=Tot.Scr.Time_mean) %>%
  mutate(Tot.Scr_A_diff=A-Baseline,
         Tot.Scr_B_diff=B-Baseline) %>%
  dplyr::select(ID,Tot.Scr_A_diff,Tot.Scr_B_diff,schoolday)



data_32=data_3 %>%
  group_by(ID,Phase,schoolday) %>%
  summarize(Tot.Soc.Time_mean=mean(Tot.Soc.Time,na.rm=T)) %>%
  spread(key=Phase,value=Tot.Soc.Time_mean) %>%
  mutate(Tot.Soc_A_diff=A-Baseline,
         Tot.Soc_B_diff=B-Baseline) %>%
  dplyr::select(ID,Tot.Soc_A_diff,Tot.Soc_B_diff,schoolday)


data_3=data_3 %>%
  group_by(ID,Phase,schoolday) %>%
  summarize(Pickups_mean=mean(as.numeric(Pickups),na.rm=T)) %>%
  spread(key=Phase,value=Pickups_mean) %>%
  mutate(Pickups_A_diff=A-Baseline,
         Pickups_B_diff=B-Baseline) %>%
  dplyr::select(ID,Pickups_A_diff,Pickups_B_diff,schoolday) %>%
  left_join(data_31,by=c("ID","schoolday")) %>%
  left_join(data_32,by=c("ID","schoolday")) %>% 
  left_join(data_baseline,by=c("ID"))

data_3=data_3[,1:21]

title=c("","","Total Screen Time","","Social Screen Time","","Numbers of pick ups","")
tbl=c("","","Mean(95%CI)","P value","Mean(95%CI)","P value","Mean(95%CI)","P value")

for(intervention in c("A","B")){
  for(day in c("Overall","School Day","Non-School Day",
               "Monday","Tuesday","Wednesday","Thursday","Friday","Saturday","Sunday")){
    data_temp=data_3 %>%
      filter(schoolday==day)
    
    if(day=="Overall"){
      out=c(paste0("Intervention ",intervention),day)
    } else {
      out=c("",day) 
    }
    for(var in c("Tot.Scr","Tot.Soc","Pickups")){
      t=t.test(data_temp[[paste0(var,"_",intervention,"_diff")]], mu = 0, alternative = "two.sided")
      
      t_mean=round(t$estimate,2)
      t_p=round(t$p.value,3)
      if(t_p==0){
        t_p="<0.001**"
      }else if(t_p<0.05){
        t_p=paste0(t_p,"*")
      }
      t_ul=round(t$conf.int[[1]],2)
      t_ll=round(t$conf.int[[2]],2)
      t_mean=paste0(t_mean,"(",t_ul,",",t_ll,")")
      
      out=c(out,t_mean,t_p)  
    }
    tbl=rbind(tbl,out)
  }
}

tbl=rbind(title,tbl)

xlsx.addTitle(sheet, rowIndex=1,
              title="Table3. Mean difference of Screen Usage before and after intervention stratified by School Day",
              titleStyle = TEXT_STYLE)

addDataFrame(tbl, sheet, startRow=3, startColumn=2,row.names  = F,col.names = F)

rowindex=3+nrow(tbl)+2


##Table4:


contvarlist=c("workmate","academic","non.academic","age","course.hours","siblings","apps","devices","procrastination")
catvarlist=c("pets","sex","degree","job")

outcome="Tot.Scr"
subtitlelist=c("a","b","c","d","e","f","g")
i=1

for(intervention in c("A","B")){
  
  for(day in c("Overall","School Day","Non-School Day")){
    subtitle=subtitlelist[i]
    unadjusted_tbl=getSLR(outcome=paste0(outcome,"_",intervention,"_diff"),
                          data=data_3[data_3$schoolday==day,],contvarlist = contvarlist,catvarlist = catvarlist)
    
    formu=as.formula(paste0(outcome,"_",intervention,"_diff","~",paste0(contvarlist,collapse = " + ")," + ",paste0("as.factor(",catvarlist,")",collapse = " + ")))
    fit=lm(formu,data=data_3[data_3$schoolday==day,])
    adjusted_tbl=as.data.frame(getMLR(fit,data=data_3[data_3$schoolday==day,],digit=2)$tbl)
    adjusted_tbl$ID2=seq.int(nrow(adjusted_tbl))
    #Model selection:
    result=stepAIC(fit,direction="both")
    selected_var=attr(result$terms , "term.labels")
    formu=as.formula(paste0(outcome,"_",intervention,"_diff","~",paste0(selected_var,collapse = " + ")))
    fit2=lm(formu,data=data_3[data_3$schoolday==day,])
    adjusted_tbl2=getMLR(fit2,data=data_3[data_3$schoolday==day,],digit=2)$tbl
    unadjusted_tbl=rbind(c("Intercept","","","","Intercept"),unadjusted_tbl)
    
    tbl=merge(unadjusted_tbl,adjusted_tbl,by="ID",all=T)
    tbl=merge(tbl,adjusted_tbl2,by="ID",all=T)
    tbl=tbl[order(tbl$ID2),]
    tbl=tbl[,c(2,3,4,5,8,9,13,14)]
    tbl=rbind(c("","","Unadjusted","","Adjusted","","",""),
              c("","","","","Full Model","","StepWise Selection",""),
              c("Factor","Level","Coefficient (95%CI)","P value","Coefficient (95%CI)","P value","Coefficient (95%CI)","P value"),tbl)
    
    xlsx.addTitle(sheet, rowIndex=rowindex,
                  title=paste0("Table4.",i,". Mean difference of Total Screen Time before and after intervention ",intervention,
                               " among ",day," Data"),
                  titleStyle = TEXT_STYLE)
    addDataFrame(tbl, sheet, startRow=rowindex+2, startColumn=2,row.names  = F,col.names = F)
    rowindex=3+nrow(tbl)+2+rowindex
    i=i+1
  }
  
}

#Get appendix data:

sheet <- createSheet(wb, sheetName = "Appendix")
xlsx.addTitle<-function(sheet, rowIndex, title, titleStyle,colIndex=1){
  rows <-createRow(sheet,rowIndex=rowIndex)
  sheetTitle <-createCell(rows, colIndex=colIndex)
  setCellValue(sheetTitle[[1,1]], title)
  setCellStyle(sheetTitle[[1,1]], titleStyle)
}


rowindex=1

contvarlist=c("workmate","academic","non.academic","age","course.hours","siblings","apps","devices","procrastination")
catvarlist=c("pets","sex","degree","job")
varlist=c(contvarlist,catvarlist)

sig_varlist=c()
pos_varlist=c()
neg_varlist=c()
selected_varlist=c()

outcome_namelist=c("Total Screen Time","Social Screen Time","Numbers of pick ups")

i=1
outi=1
for(outcome in c("Tot.Scr","Tot.Soc","Pickups")){
  outcome_name=outcome_namelist[outi]
  for(intervention in c("A","B")){
    
    for(day in c("Overall","School Day","Non-School Day",
                 "Monday","Tuesday","Wednesday","Thursday","Friday","Saturday","Sunday")){
      
      unadjusted_tbl=getSLR(outcome=paste0(outcome,"_",intervention,"_diff"),
                            data=data_3[data_3$schoolday==day,],contvarlist = contvarlist,catvarlist = catvarlist)
      
      formu=as.formula(paste0(outcome,"_",intervention,"_diff","~",paste0(contvarlist,collapse = " + ")," + ",paste0("as.factor(",catvarlist,")",collapse = " + ")))
      fit=lm(formu,data=data_3[data_3$schoolday==day,])
      adjusted_tbl=as.data.frame(getMLR(fit,data=data_3[data_3$schoolday==day,],digit=2)$tbl)
      adjusted_tbl$ID2=seq.int(nrow(adjusted_tbl))
      #Model selection:
      result=stepAIC(fit,direction="both")
      selected_var=attr(result$terms , "term.labels")
      if(length(selected_var)!=0){
        formu=as.formula(paste0(outcome,"_",intervention,"_diff","~",paste0(selected_var,collapse = " + ")))
        fit2=lm(formu,data=data_3[data_3$schoolday==day,])
        adjusted_tbl2=getMLR(fit2,data=data_3[data_3$schoolday==day,],digit=2)$tbl
        unadjusted_tbl=rbind(c("Intercept","","","","Intercept"),unadjusted_tbl)
        
        tbl=merge(unadjusted_tbl,adjusted_tbl,by="ID",all=T)
        tbl=merge(tbl,adjusted_tbl2,by="ID",all=T)
        tbl=tbl[order(tbl$ID2),]
        tbl=tbl[,c(2,3,4,5,8,9,13,14)]
        tbl=rbind(c("","","Unadjusted","","Adjusted","","",""),
                  c("","","","","Full Model","","StepWise Selection",""),
                  c("Factor","Level","Coefficient (95%CI)","P value","Coefficient (95%CI)","P value","Coefficient (95%CI)","P value"),tbl)
        
        #Get significant varlist:
        getresult=getMLR(fit2,data=data_3[data_3$schoolday==day,],digit=2)
        sig_var=getresult$sig_list
        pos_var=getresult$pos_list
        neg_var=getresult$neg_list
        #
        sig_varlist=rbind(sig_varlist,
                          c(paste0("Intervention ",intervention),outcome_name,day,varlist %in% sig_var))
        pos_varlist=rbind(pos_varlist,
                          c(paste0("Intervention ",intervention),outcome_name,day,varlist %in% pos_var))
        neg_varlist=rbind(sig_varlist,
                          c(paste0("Intervention ",intervention),outcome_name,day,varlist %in% neg_var))
        selected_var=str_remove_all(selected_var,"as.factor\\(|\\)")
        selected_varlist=rbind(selected_varlist,
                               c(paste0("Intervention ",intervention),outcome_name,day,varlist %in% selected_var))
        
      } else {
        unadjusted_tbl=rbind(c("Intercept","","","","Intercept"),unadjusted_tbl)
        tbl=merge(unadjusted_tbl,adjusted_tbl,by="ID",all=T)
        tbl=tbl[order(tbl$ID2),]
        tbl=tbl[,c(2,3,4,5,8,9)]
        tbl=rbind(c("","","Unadjusted","","Adjusted",""),
                  c("","","","","Full Model",""),
                  c("Factor","Level","Coefficient (95%CI)","P value","Coefficient (95%CI)","P value"),tbl)
        
        sig_varlist=rbind(sig_varlist,
                          c(paste0("Intervention ",intervention),outcome_name,day,varlist %in% NULL))
        pos_varlist=rbind(pos_varlist,
                          c(paste0("Intervention ",intervention),outcome_name,day,varlist %in% NULL))
        neg_varlist=rbind(sig_varlist,
                          c(paste0("Intervention ",intervention),outcome_name,day,varlist %in% NULL))
      }
      
      xlsx.addTitle(sheet, rowIndex=rowindex,
                    title=paste0("Table5.",i,". Mean difference of ",outcome_name," before and after intervention ",intervention,
                                 " among ",day," Data"),
                    titleStyle = TEXT_STYLE)
      addDataFrame(tbl, sheet, startRow=rowindex+2, startColumn=2,row.names  = F,col.names = F)
      rowindex=3+nrow(tbl)+2+rowindex
      i=i+1
    }
    
  }
  outi=outi+1
}


colnames(sig_varlist)=c("Intervention","Outcome","Day",varlist)
colnames(pos_varlist)=c("Intervention","Outcome","Day",varlist)
colnames(neg_varlist)=c("Intervention","Outcome","Day",varlist)
colnames(selected_varlist)=c("Intervention","Outcome","Day",varlist)


for(filename in c("sig_varlist","pos_varlist","neg_varlist","selected_varlist")){
  file=get(filename)
  if(filename=="sig_varlist"){
    long=reshape2::melt(as.data.frame(file),measure.vars=4:16,variable.name="Variable") 
    filename=str_remove(filename,"_varlist")
    long[[filename]]=long$value
    long=long %>%
      dplyr::select(!value)
  } else {
    temp=reshape2::melt(as.data.frame(file),measure.vars=4:16,variable.name="Variable")
    filename=str_remove(filename,"_varlist")
    temp[[filename]]=temp$value
    long=long %>%
      left_join(temp,by=c("Intervention","Outcome","Day","Variable")) %>%
      dplyr::select(!value)
  }
 
}
pd=position_dodge(width=0.75)
long$Day=factor(long$Day,levels =c("Overall","School Day","Non-School Day",
                                    "Monday","Tuesday","Wednesday","Thursday","Friday","Saturday","Sunday") )
long$Day_numeric=as.numeric(long$Day)
long=long %>%
  mutate(nonsig= !as.logical(sig) & as.logical(selected))
weekday=c("Overall","School Day","Non-School Day","Monday","Tuesday","Wednesday","Thursday","Friday","Saturday","Sunday")
long$Variable_num=as.numeric(long$Variable)
var_level=levels(long$Variable)

i=1
for(outcome in outcome_namelist){
  for(intervention in c("Intervention A","Intervention B")){
    
    long_pos=long %>%
      filter(pos==T & Outcome==outcome & Intervention==intervention ) 
    long_neg=long %>%
      filter(neg==T & Outcome==outcome & Intervention==intervention )
    long_nonsig=long %>%
      filter(nonsig==T & Outcome==outcome & Intervention==intervention )
    
    if(i!=6){
      p=ggplot() + 
        geom_point(data=long_nonsig,aes(Day_numeric,Variable_num,color="Non-significant")) +
        geom_point(data=long_neg,aes(Day_numeric,Variable_num,col="Negatively significant")) +
        geom_point(data=long_pos,aes(Day_numeric,Variable_num,col="Positvely significant")) +
        scale_x_discrete(name="Day",limits=weekday) + 
        scale_y_discrete(limits=var_level) + 
        labs(colour="")+
        ggtitle(paste0(intervention,": ",outcome))+
        theme(axis.text.x =  element_text(size = 8, angle = 45,hjust=1),
              axis.title.x = element_blank(),
              axis.title.y = element_blank(),
              plot.title =element_text(size=10,face="bold",hjust=0.5))+
        theme(legend.position = "none")
    } else {
      p=ggplot() + 
        geom_point(data=long_nonsig,aes(Day_numeric,Variable_num,color="Non-significant")) +
        geom_point(data=long_neg,aes(Day_numeric,Variable_num,col="Negatively significant")) +
        geom_point(data=long_pos,aes(Day_numeric,Variable_num,col="Positvely significant")) +
        scale_x_discrete(name="Day",limits=weekday) + 
        scale_y_discrete(limits=var_level) + 
        labs(colour="")+
        ggtitle(paste0(intervention,": ",outcome))+
        theme(axis.text.x =  element_text(size = 8, angle = 45,hjust=1),
              axis.title.x = element_blank(),
              axis.title.y = element_blank(),
              plot.title =element_text(size=10,face="bold",hjust=0.5))
    }
     
    
    assign(paste0("p",i),p)
    i=i+1
  }
}

png("Selection.png",width=600, height=600,units = "px")


pt1=plot_grid(p1,p2,NULL,p3,p4,NULL,nrow=2,rel_widths =c(1,1,0.5) )

pt2=plot_grid(p5,p6,rel_widths = c(1,1.55))

plot_grid(pt1,pt2,nrow=2,rel_heights = c(2,1))

dev.off()

saveWorkbook(wb, paste("620_Project2",date,".xlsx",sep=""),)
