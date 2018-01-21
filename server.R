# server.R

## translating the EWAS stata code to R
## 
## First complete version 09 May 2017

## Revised June 20 2017
## Revised 2017-11-20
### Using Observe event 2017-11-27

### Installing required packages
pkgs<-c('XLConnect','plyr','dplyr','car','stringr','zoo','foreign','ggplot2','splines','mgcv','Hmisc','xtable','foreach','xlsx','lattice','latticeExtra',"gridExtra","grid","shiny")
for (i in length(pkgs)){
  if(pkgs[i] %in% rownames(installed.packages()) == FALSE) {
    install.packages(pkgs[i])
  }
}

library(XLConnect)
library(plyr)## This package should be loaded before dplyr
library(dplyr)
library(car)
library(stringr)
library(zoo)
library(foreign)
library(ggplot2)
library(splines)
library(mgcv)
library(Hmisc)
library(xtable)
library(foreach)
library(xlsx)## Ensure java installed on system same as one for R, e.g 64 bit R with 64 bit Java
library(lattice)
library(latticeExtra)
library(gridExtra)
library(grid)

server<-function(input,output,session) {
  
  out_x<-eventReactive(input$goButton, {
  
    dir<-input$dir
    to_fold<-paste(input$dir,"/ewars app/www",sep='')
    original_data_file_name<-input$original_data_file_name
    original_data_sheet_name<-input$original_data_sheet_name
    stop_runin<-input$stop_runin
    generating_surveillance_workbook<-input$generating_surveillance_workbook
    run_per_district<-as.numeric(str_split(input$run_per_district,',',simplify =T))
    population.a<-input$population
    number_of_cases<-input$number_of_cases
    outbreak_week_length<-input$outbreak_week_length
    alarm_indicators<-input$alarm_indicators
    alarm_window<-input$alarm_window
    alarm_threshold<-input$alarm_threshold
    graph_per_district<-input$graph_per_district
    season_length<-input$season_length
    z_outbreak<-input$z_outbreak
    outbreak_window<-input$outbreak_window
    prediction_distance<-input$prediction_distance
    outbreak_threshold<-input$outbreak_threshold
    spline<-isolate(input$spline)                            
    
    setwd(dir)
    
    files_www<-list.files(to_fold,full.names =T)
    unlink(files_www)
    
    data<- xlsx::read.xlsx2(original_data_file_name,sheetName=original_data_sheet_name,colClasses=NA,header=T)
    #data<- read.csv(original_data_file_name(),header=T)
    
    data<-data %>% filter(!week %in% c(1,53)) 
    
    
    if("population" %in% names(data)){
      data<-data %>% mutate(outbreak=with(data,(get(number_of_cases)/population)*1000))
      
    }else {
      data<-data %>% mutate(outbreak=(get(number_of_cases)/get(population.a))*1000)
    }
    
    
    if(length(run_per_district)>0){
      data<-data %>% filter(district %in% run_per_district)
    }else{
      stop("Choose atleast one district")
    }
    
    
    if(outbreak_threshold> 1|outbreak_threshold<= 0) { 
      stop( " You specified an invalid value for the threshold variable. The threshold variable should be in (0,1]. ")
      
    }
    
    
    if(!spline %in% c(TRUE,FALSE)){
      stop("You specified an invalid value for the spline. The spline option should takes value either 0(No) or 1(Yes).")
    }
    
    
    ############# Runin: Here starting the analysis
    #To capture the length of runin period and evaluation period
    
    data<-data %>% mutate(year_week=year*100+week)
    
    
    
    if(generating_surveillance_workbook==FALSE){
      data<-data %>% mutate(runin=as.numeric((stop_runin>=year_week)))
    }
    
    if(generating_surveillance_workbook==TRUE){
      data<-data %>% mutate(runin=1)
    }
    
    al_var<- str_split(alarm_indicators,pattern=",",simplify =T)
    n_alarm_indicators=length(al_var)
    
    
    
    # create matrix size of data for indicator variables
    
    
    #al_tab<-data.frame(matrix(NA,nrow(data),n_alarm_indicators))
    #names(al_tab)<-paste("alarm",1:n_alarm_indicators,sep="")
    
    #for(i in 1:n_alarm_indicators){
    # al_tab[,i]<-get(al_var[i])
    
    #}
    #data<-cbind(data,al_tab)
    
    data<-within(data,{
      for (i in 1 :n_alarm_indicators){
        
        assign(paste('alarm',i,sep=""),get(al_var[i]))
      }
      rm(i)
    })
    
    
    #Gives the date/time for the output files
    
    
    format.Date(Sys.time(),"%d%b%Y_%H%M%S")
    
    #egen district_no=group(district)
    
    data<-data %>% mutate(district_no=as.numeric(as.factor(district)))
    #sort district_no year_week
    
    data<-data %>% arrange(district_no,year_week)
    
    #gen obs_no=_n
    data<-data %>% mutate(obs_no=1:nrow(data))
    
    #1 - Smoothing
    #gen outbreak_tmp=outbreak if runin==1
    
    data$outbreak_tmp<-as.integer(NA)
    data$outbreak_tmp[data$runin==1]<-data$outbreak[data$runin==1]
    
    #tsset district_no year_week - make as time series dataset
    #tssmooth ma outbreak_moving_tmp=outbreak_tmp, weights(1 1 1 <1> 1 1 1) replace
    #test_dat<<-data
    ## made change on 2017-11-21, sum then divide by 7 
    data<- data %>% group_by(district_no) %>% 
      mutate(outbreak_moving_tmp_sum=rollapply(outbreak_tmp,FUN=sum,width=list(-3:3),align = "center",fill = NA,na.rm = T,partial=T),
             outbreak_moving_tmp_sum1=case_when(outbreak_moving_tmp_sum>0~outbreak_moving_tmp_sum,
                                                TRUE~as.double(NA)),
             
             outbreak_moving_tmp=outbreak_moving_tmp_sum1/7)
    #mutate(outbreak_moving_tmp=rollapply(data=outbreak_tmp,FUN=mean,width=list(-3:3),align = "left",fill = NA,partial=T) )
    ## replace runin=0 with NA for NA
    
    data$outbreak_tmp[data$runin==0]<-NA
    ## test a different algorith to compute running mean for lag 3 and lead 3 19 April 2017
    
    #data<-within(data,{
    #create matrix to hold values
    #mat_val2<-matrix(NA,nrow(data),length(-3:3))
    #for (j in -3:3){
    
    #if(j<0){
    
    #to_put2<-dplyr::lag(outbreak_tmp,-j)
    #}
    
    #else if(j>=0){
    #to_put2<-dplyr::lead(outbreak_tmp,j)
    #}
    
    #mat_val2[,j+4]<-to_put2
    #}
    
    #mean_v2<-apply(mat_val2,1,FUN=sum,na.rm=T)
    
    #assign(paste('outbreak_moving_tmp1',sep=""),mean_v2)
    
    #rm(j,to_put2,mean_v2,mat_val2)
    #})
    
    data<-ungroup(data)
    
    
    data<-data %>% group_by(district_no,week) %>% mutate(outbreak_moving=mean(outbreak_moving_tmp,na.rm=T),outbreak_moving_sd=sd(outbreak_moving_tmp,na.rm=T))
    
    data<-ungroup(data)
    
    data<-data %>% arrange(obs_no)
    
    
    #Estimating parameters
    
    
    data<-data %>% mutate(outbreak_moving_limit=outbreak_moving+(z_outbreak*outbreak_moving_sd))
    
    
    #data$outbreakweek<-0
    #data<-data %>% mutate(outbreakweek=as.numeric((outbreak >=outbreak_moving_limit & !is.na(outbreak))
    # |(outbreak>=outbreak_moving_limit & outbreak_moving_limit!=0 & outbreak!=0)))
    
    #replace missing outbreak to Inf
    ##replace outbreakweek=1 if outbreak>=outbreak_moving_limit&outbreak!=. | outbreak>=outbreak_moving_limit&outbreak_moving_limit!=0&outbreak!=0
    
    data$outbreakweek<-0
    data<-within(data,{
      
      #outbreak[is.na(outbreak)]<-Inf
      cond_1<-(outbreak >=outbreak_moving_limit & !is.na(outbreak))
      #removed |(outbreak>=outbreak_moving_limit & outbreak_moving_limit!=0 & outbreak!=0)
      outbreakweek[cond_1]<-1
      #outbreak[outbreak==Inf]<-NA
      rm(cond_1)
    })
    
    
    
    #data[which(cond_1),]$outbreakweek<-1
    
    data<-data %>% arrange(obs_no)
    
    #data$outbreakperiod=NA
    
    data$sum_outbreak<-0
    data$none_outbreakweek<-1
    data$none_outbreakweek[which(data$outbreakweek==1)]<-0
    data$sum_none_outbreak<-0
    
    outbreak_weeks_start_1<-outbreak_week_length-1
    
    for(i in 0:outbreak_weeks_start_1){
      p<-which(data$district_no==lag(data$district_no,i))
      data$sum_outbreak[p]<-data$sum_outbreak[p]+lag(data$outbreakweek,i)[p]
      data$sum_none_outbreak[p]<-data$sum_none_outbreak[p]+lag(data$none_outbreakweek,i)[p]
    }
    
    #data$outbreakperiod[which(data$sum_outbreak==outbreak_week_length())]<-1
    #data$outbreakperiod[which(data$sum_none_outbreak==outbreak_week_length())]<-0
    
    #outbreakperiod[_n-1]==1 & outbreakperiod[_n]==. & district[_n-1]==district[_n]
    #p1<-which(is.na(data$outbreakperiod) & lag(data$outbreakperiod,1)==1 &(data$district_no==lag(data$district_no,1)))
    #data$outbreakperiod[p1]<-1
    #data$outbreakperiod[is.na(data$outbreakperiod)]<-0
    
    ##  a within implementation
    test_dat<<-data
    
    data$outbreakperiod<-as.integer(NA)
    
    
    
    data<-within(data,{
      #outbreakperiod<-NA
      cond1<-which(sum_outbreak==outbreak_week_length)
      outbreakperiod[cond1]<-1
      cond2<-which(sum_none_outbreak==outbreak_week_length)
      outbreakperiod[cond2]<-0
      ## fill missing with previous 4 the following code
      #cond2.a<-which(is.na(outbreakperiod) & !is.na(lag(outbreakperiod,1)) &(district_no==lag(district_no,1)))
      #outbreakperiod[cond2.a]<-lag(outbreakperiod,1)[cond2.a]
      ## added cascading code to match stata cascading behavior 28 July 2017
      repeat({
        cond3<-which(is.na(outbreakperiod) & lag(outbreakperiod,1)==1 &(district_no==lag(district_no,1)))
        outbreakperiod[cond3]<-1
        if(length(cond3)==0){break}
        
      }
      )
      cond4<-which(is.na(outbreakperiod))
      outbreakperiod[cond4]<-0
      rm(cond1,cond2,cond3,cond4)
    })
    
    #dat_test<<-data
    
    ## use case_when implementation
    
    
    dat_lab1<-paste("Data_showing_runin(R)_",format.Date(Sys.Date(),"%d%b%Y"),'.dta',sep="")
    write.dta(data,dat_lab1)
    
    test_ggplot<<-data
    ## do the first plot
    ##
    if(generating_surveillance_workbook==FALSE){
      rand_n<-sample(1:1e8,1)
      for(j in 1:length(run_per_district)){
        dat<-data %>% filter(district==run_per_district[j] & runin==1)
        #replace outbreak_moving_limit=0 if outbreak_moving_limit==.
        dat$outbreak_moving_limit<-ifelse(is.na(dat$outbreak_moving_limit),0,dat$outbreak_moving_limit)
        dat$outbreakNA<-NA
        
        endmic<-paste("Endemic Channel (z outbreak=",z_outbreak,')',sep="")
        test_ggplot1<<-dat
        
        ## add condition for when no outbreak
        if(sum(dat$outbreakperiod)>0){
          
          plot_a<-ggplot(aes(x=obs_no,y=outbreak_moving_limit),data=dat)+ 
            geom_area(aes(fill="endmic"))+
            geom_line(aes(x=obs_no,y=outbreak,lty="Cases per 1000 pop"),color="tomato2")+
            geom_point(aes(x=obs_no,y=outbreak,shape="Outbreak period"),data=subset(dat,dat$outbreakperiod==1),size=2,colour="red")+
            #scale_y_continuous(breaks=seq(0,0.03,0.01),labels=seq(0,0.03,0.01),limits=c(0,0.03))+
            ylab("Cases per 1000 pop")+
            xlab("Epidemiological week (runin period)")+
            ggtitle(paste('Runin Period District:',run_per_district[j]))+
            theme_bw()+
            theme(panel.grid.minor.x=element_blank(),
                  panel.grid.major.x=element_blank(),
                  panel.grid.minor.y=element_blank(),
                  panel.background=element_rect(color="white"),
                  legend.position = "bottom",
                  legend.text.align=1,
                  legend.title=element_blank())+
            scale_fill_manual(values=c("endmic"="lightblue2"),labels=endmic)+
            scale_linetype_manual(values=c('Cases per 1000 pop'=1))+
            scale_shape_manual(values=("Outbreak period"=1))+
            scale_shape(solid =T)
        }else{
          plot_a<-ggplot(aes(x=obs_no,y=outbreak_moving_limit),data=dat)+ 
            geom_area(aes(fill="endmic"))+
            geom_line(aes(x=obs_no,y=outbreak,lty="Cases per 1000 pop"),color="tomato2")+
            geom_point(aes(y=as.integer(NA),shape="Outbreak period"),size=2,colour="red")+
            #scale_y_continuous(breaks=seq(0,0.03,0.01),labels=seq(0,0.03,0.01),limits=c(0,0.03))+
            ylab("Cases per 1000 pop")+
            xlab("Epidemiological week (runin period)")+
            ggtitle(paste('Runin Period District:',run_per_district[j]))+
            theme_bw()+
            theme(panel.grid.minor.x=element_blank(),
                  panel.grid.major.x=element_blank(),
                  panel.grid.minor.y=element_blank(),
                  panel.background=element_rect(color="white"),
                  legend.position = "bottom",
                  legend.text.align=1,
                  legend.title=element_blank())+
            scale_fill_manual(values=c("endmic"="lightblue2"),labels=endmic)+
            scale_linetype_manual(values=c('Cases per 1000 pop'=1))+
            scale_shape_manual(values=("Outbreak period"=1))+
            scale_shape(solid =T)
          
        }
        
        #assign(paste("plot_a",j,sep=''),plot_a)
        
        print(plot_a)
        
        ggsave(filename=paste('Runin_period_for_district(',run_per_district[j],')_',rand_n,'-',format(Sys.Date(),"%d%b%y"),'.png',sep=''),width=20,device="png",
               height=12,units='cm',dpi=300)
        #print(get(paste("plot_a",j,sep='')))
        #print(get(paste("plot_a",j,sep='')))
        
        #yl1<-max(dat$outbreak,na.rm=T)+max(dat$outbreak,na.rm=T)/10
        #obj0<-xyplot(outbreak_moving_limit~obs_no,dat,panel=panel.xyarea,origin=0,border=F,col="lightblue",ylim=c(0,yl1))
        #obj0<-xyplot(outbreak_moving_limit~obs_no,dat,panel=panel.xyarea,origin=0,border=F,col="lightblue")
        #obj1<-xyplot(outbreak ~ obs_no,dat, type = "l",col="tomato2",lwd=2,grid=T)
        #obj2<-xyplot(outbreak ~ obs_no,subset(dat,dat$outbreakperiod==1),
        #par.settings = list(plot.symbol=list(col = "red",cex = 0.8,pch = 16)),type = "p")
        #plot_a_x<-obj0+obj1+obj2
        
        #plot_a<-update(plot_a_x,
        #main=paste('Runin Period District:',run_per_district()[j]),
        #ylab="Cases per 1000 pop",
        # xlab="Epidemiological week",
        #grid=T,
        #key=list(space="bottom",text=list(text=c(endmic,"Cases per 1000 pop","Outbreak period")),
        #columns=1,
        #lines=list(lty=c(1,1,0),col=c("lightblue2","tomato2",NA), lwd=c(3.5,2,NA)),
        #points=list(pch=c(NA,NA,16),col=c(NA,NA,"red")),cex=0.7
        #))
        
        #assign(paste("plot_a",j,sep=''),plot_a)
        
        #png(paste('Runin_period_for_district(',run_per_district()[j],')_',format(Sys.Date(),"%d%b%y"),'.png',sep=''),width=17,
        #height=12,units='cm',res=300)
        
        #print(get(paste("plot_a",j,sep='')))
        
        
        #dev.off()
        ##copy the images to www to render on UI
        #to_fold()<-"C:\\joacim\\How To Guide-Demo Materials\\ewars app\\www"
        img_nam<-paste('Runin_period_for_district(',run_per_district[j],')_',rand_n,'-',format(Sys.Date(),"%d%b%y"),'.png',sep='')
        unlink(paste(to_fold,'/',img_nam,sep=''))
        file.copy(img_nam,to_fold,overwrite=T)
        
        
        
        
      }
    }
    
    ## add code to combine plots into 1 using facetwrap
    #dat.a1<-data %>% filter(runin==1)
    #replace outbreak_moving_limit=0 if outbreak_moving_limit==.
    #dat.a1$outbreak_moving_limit<-ifelse(is.na(dat.a1$outbreak_moving_limit),0,dat.a1$outbreak_moving_limit)
    
    #endmic<-paste("Endemic Channel (z outbreak=",z_outbreak(),')',sep="")
    
    
    #plot_a.21<-ggplot(aes(x=obs_no,y=outbreak_moving_limit),data=dat.a1)+ 
    #geom_area(aes(fill="endmic"))+
    #geom_line(aes(x=obs_no,y=outbreak,lty="Cases per 1000 pop"),color="tomato2")+
    #geom_point(aes(x=obs_no,y=outbreak,shape="Outbreak period"),data=subset(dat.a1,dat.a1$outbreakperiod==1),size=2,colour="red")+
    #scale_y_continuous(breaks=seq(0,0.03,0.01),labels=seq(0,0.03,0.01),limits=c(0,0.03))+
    #ylab("Cases per 1000 pop")+
    #xlab("Epidemiological week (runin period)")+
    #ggtitle(paste('Runin Period District:',run_per_district()[j]))+
    # theme_bw()+
    #theme(panel.grid.minor.x=element_blank(),
    # panel.grid.major.x=element_blank(),
    # panel.grid.minor.y=element_blank(),
    # panel.background=element_rect(color="white"),
    #legend.position = "bottom",
    #legend.text.align=1,
    #legend.title=element_blank())+
    #scale_fill_manual(values=c("endmic"="lightblue2"),labels=endmic)+
    #scale_linetype_manual(values=c('Cases per 1000 pop'=1))+
    #scale_shape_manual(values=("Outbreak period"=1))+
    #scale_shape(solid =T)+
    #facet_wrap(~paste('Runin Period District:',as.factor(district)),ncol=1,scales="free")
    
    
    #print(plot_a.21)
    
    
    #get(paste("plot_a",1,sep=''))
    #get(paste("plot_a",2,sep=''))
    #Below is the mean alarm for runing and evaluation
    
    #loop for calculating mean alarm for each alarm indicator
    
    data<-within(data,{
      mat_val<-matrix(NA,nrow(data),length(0 :(alarm_window-1)))
      for (i in 1 :n_alarm_indicators){
        for (j in 0 :(alarm_window-1)){
          
          #create matrix to hold values
          
          
          to_put<-dplyr::lag(get(paste('alarm',i,sep="")),j)
          to_put[which(!district==dplyr::lag(district,j))]<-NA
          
          
          mat_val[,j+1]<-to_put
          
          assign(paste('creating_mean',j,sep=""),to_put)
          
        }
        
        mean_v<-apply(mat_val,1,FUN=mean,na.rm=T)
        assign(paste('mean_alarm',i,sep=""),mean_v)
        
      }
      rm(i,j,to_put,mean_v,mat_val)
    })
    
    
    #Below is the mean of outbreakperiods 
    
    data<-within(data,{
      mat_val2<-matrix(NA,nrow(data),length(0 :(outbreak_window-1)))
      for (j in 0 :(outbreak_window-1)){
        #create matrix to hold values
        
        
        to_put2<-dplyr::lead(outbreakperiod,(j+prediction_distance))
        to_put2[which(!district==dplyr::lead(district,(j+prediction_distance)))]<-NA
        
        mat_val2[,j+1]<-to_put2
        
        assign(paste('creating_mean_outbreak',j,sep=""),to_put2)
        
      }
      
      
      mean_v2<-apply(mat_val2,1,FUN=mean,na.rm=T)
      
      assign(paste('outbreak_mean',sep=""),mean_v2)
      
      rm(j,to_put2,mean_v2,mat_val2)
    })
    
    
    
    #Below is the regression
    
    #Generating variable for merging the coefficients for each season
    data$season<-NA
    
    data<-within(data,{
      
      for(s in seq(min(week),season_length,52)){
        cond<-(week>=s & week<(s+season_length))
        season[cond]<-s
      }
      rm(cond,s)
      
    })
    
    
    #Binary variable "outbreak_mean_cutoff" (i.e.0 if sum of outbreaks < threshold and 1 if sum of outbreaks >= threshold ) 
    
    data$outbreak_mean_cutoff<-0
    
    data<-within(data,{
      cond<-which(outbreak_mean>=outbreak_threshold & !is.na(outbreak_mean))
      outbreak_mean_cutoff[cond]<-1
      rm(cond)
    })
    
    if (spline==FALSE){
      
      
      #* To generate list of all alarm indicators names (e.g. mean_alarm1 mean_alarm2 mean_alarm3 and so on  )
      loop<- n_alarm_indicators+1
      independ_var_name<-paste("mean_alarm",1:n_alarm_indicators,sep="")
      
      #* To generate list of all coefficients names(e.g. c1 c2 c3 and so on)
      cols_name<-paste("coef",1:loop,sep="")
      season_start<-min(data$week)
      
      #loop over district
      
      
      out_mat<-matrix(NA,length(run_per_district),(length(cols_name)+2))
      
      for(d in 1:length(run_per_district)){
        
        
        #loop over season
        for(s in seq(season_start,season_length,52)){
          
          #sum outbreak_mean_cutoff if week>=`s' & week<`s'+`=season_length()' & district==distr_vector[1,`d'] & runin==1
          
          
          mean_t<-with(subset(data,week>=s & week<(s+season_length) & district==run_per_district[d] & runin==1),summary(outbreak_mean_cutoff))
          
          
          
          if(as.numeric(mean_t["Mean"])==0 |as.numeric(mean_t["Mean"])==1 ){
            
            print(paste("Regression for district",run_per_district[d],"has failed"))
            out_mat[d,]<-c(s,run_per_district[d],rep(NA,n_alarm_indicators+1))
          }
          
          else if(as.numeric(mean_t["Mean"])>0){
            
            #* Logistic regression calibration based on mean outbreaks of the coming N weeks >= outbreak cut-off 
            #with the mean alarm of the last N weeks(seasonal run over all years) 
            
            dd<-c(paste(independ_var_name,collapse ="+"))
            
            
            mod<-with(subset(data,
                             week>=s & week<(s+season_length) & district==run_per_district[d] & runin==1),
                      summary(glm(as.formula(c("outbreak_mean_cutoff~",dd)),family="binomial")))
            coef1<-coefficients(mod)[,1][c("(Intercept)",independ_var_name)]
            
            #To save the season and district name and coefficients matrix
            print(paste("Regression for district",run_per_district[d]))
            out_mat[d,]<-c(s,run_per_district[d],as.numeric(coef1))
          }
          
          
        }
        
      }
      colnames(out_mat)<-c("season","district",cols_name)
      
      write.dta(data.frame(out_mat),"Out_param_r.dta")
      
      data<-merge(data,out_mat,by=c("season","district"))
    }
    
    
    if (spline==TRUE){
      
      data<-within(data,{
        
        for(j in 1 :n_alarm_indicators) {
          
          sp1<-NA
          sp2<-NA
          sp3<-NA
          sp4<-NA
          
          for(d in 1:length(run_per_district)){
            cond<-which(district==run_per_district[d])
            for_sp<-get(paste("mean_alarm",j,sep=""))
            sp1[cond]<-rcspline.eval(for_sp[cond],nk=5,inclx=T) [,1]
            sp2[cond]<-rcspline.eval(for_sp[cond],nk=5,inclx=T) [,2]
            sp3[cond]<-rcspline.eval(for_sp[cond],nk=5,inclx=T) [,3]
            sp4[cond]<-rcspline.eval(for_sp[cond],nk=5,inclx=T) [,4]
            
          }
          
          for(k in 1:4){
            assign(paste("alarm",j,'sp',k,sep=''),get(paste("sp",k,sep="")))
          }
          
        }
        rm(sp1,sp2,sp3,sp4,cond,j,k,d,for_sp)
      })
      
      #To generate list of all alarm indicators names (e.g. alarm1sp1 alarm1sp2 alarm1sp3 alarm1sp4 alarm2sp1 alarm2sp2 alarm2sp3 alarm2sp4 so on  )
      independ_var_name<-rep(NA,(n_alarm_indicators*4))
      v=0
      for(i in 1:n_alarm_indicators){
        for(j in 1:4){
          v<-v+1
          independ_var_name[v]<-paste("alarm",i,'sp',j,sep="")
        }
      }
      
      
      loop<-((n_alarm_indicators*4)+1)
      #coefficients names are alarm indicators + 4 spline  s + district(for merging )+ season(for merging ) (e.g. c1 c2 c3 and so on)
      cols_name<-paste("coef",1:loop,sep="")
      season_start<-min(data$week)
      
      out_mat<-matrix(NA,length(run_per_district),(length(cols_name)+2))
      for(d in 1:length(run_per_district)){
        
        #loop over season
        for(s in seq(season_start,season_length,52)){
          
          #sum outbreak_mean_cutoff if week>=`s' & week<`s'+`=season_length()' & district==distr_vector[1,`d'] & runin==1
          mean_t<-with(subset(data,
                              week>=s & week<(s+season_length) & district==run_per_district[d] & runin==1),
                       summary(outbreak_mean_cutoff))
          
          if(as.numeric(mean_t["Mean"])==0 |as.numeric(mean_t["Mean"])==1 ){
            
            print(paste("Regression for district",run_per_district[d],"has failed"))
            out_mat[d,]<-c(s,run_per_district[d],rep(NA,((n_alarm_indicators*4)+1)))
          } else if(as.numeric(mean_t["Mean"])>0){
            
            #* Logistic regression calibration based on mean outbreaks of the coming N weeks >= outbreak cut-off 
            #with the mean alarm of the last N weeks(seasonal run over all years) 
            
            dd<-c(paste(independ_var_name,collapse ="+"))
            
            
            mod<-with(subset(data,
                             week>=s & week<(s+season_length) & district==run_per_district[d] & runin==1),
                      summary(glm(as.formula(c("outbreak_mean_cutoff~",dd)),family="binomial")))
            coef1<-coefficients(mod)[,1][c("(Intercept)",independ_var_name)]
            
            #To save the season and district name and coefficients matrix
            print(paste("Regression for district",run_per_district[d]))
            out_mat[d,]<-c(s,run_per_district[d],as.numeric(coef1))
          }
          
          
        }
      }
      colnames(out_mat)<-c("season","district",cols_name)
      
      write.dta(data.frame(out_mat),"Out_param_r.dta")
      
      data<-merge(data,out_mat,by=c("season","district"))
    }
    
    prob_round<-round(alarm_threshold,3)
    name<-round((prob_round*1000),1)
    
    
    data<-within(data,{
      prob_outbreak<-0
      assign(paste("alarm_",name,sep=''),0)
      alarm_threshold<-alarm_threshold
      
    })
    
    dat_lab2<-paste("Parameters_runin(R)_",format.Date(Sys.Date(),"%d%b%Y"),'.dta',sep="")
    write.dta(data,dat_lab2)
    
    data1<-with(data,subset(data,runin!=1))
    data1<-within(data1,{
      rm(list=c("prob_outbreak","alarm_threshold",paste("alarm_",name,sep='')))
    })
    
    if(generating_surveillance_workbook==FALSE){
      
      if(spline==FALSE){
        
        data1<-within(data1,{
          for_eq<-rep(NA,(n_alarm_indicators+1))
          for(j in 1:(n_alarm_indicators+1)){
            if(j==1){
              for_eq[j]<-"coef1"  
            }
            else if(j>1){
              for_eq[j]<-paste("(mean_alarm",j-1,'*',"coef",j,')',sep='')
            }
            
          }
          probability_equation<-eval(parse(text=paste(for_eq,collapse='+')))
          
          prob_outbreak<-exp(probability_equation)/(1+exp(probability_equation))
          rm(for_eq,j,probability_equation)
        })
        
      }
      
      
      
      if(spline==TRUE){
        independ_var_name<-rep(NA,(n_alarm_indicators*4))
        v=0
        for(i in 1:n_alarm_indicators){
          for(j in 1:4){
            v<-v+1
            independ_var_name[v]<-paste("alarm",i,'sp',j,sep="")
          }
        }
        
        data1<-within(data1,{
          for_eq<-rep(NA,((n_alarm_indicators*4)+1))
          for(j in 1:((n_alarm_indicators*4)+1)){
            if(j==1){
              for_eq[j]<-"coef1"  
            }
            else if(j>1){
              for_eq[j]<-paste('(',independ_var_name[j-1],'*',"coef",j,')',sep='')
            }
            
          }
          probability_equation<-eval(parse(text=paste(for_eq,collapse='+')))
          
          prob_outbreak<-exp(probability_equation)/(1+exp(probability_equation))
          rm(for_eq,j,probability_equation)
        })
        
      }
      
      prob_round<-round(alarm_threshold,3)
      name<-round((prob_round*1000),1)
      
      data1<-within(data1,{
        var_1<-NA
        cond<-(prob_outbreak>=prob_round & !is.na(prob_outbreak))
        var_1<-ifelse(cond,1,0)
        assign(paste("alarm_",name,sep=''),var_1)
        rm(var_1)
      })
      
      
      data1<-within(data1,{
        
        correct_alarm<-NA
        no_alarm_no_outbreak<-NA
        missed_outbreak<-NA
        false_alarm<-NA
        
        cond1<-(outbreak_mean_cutoff==1 & get(paste("alarm_",name,sep=''))==1)
        cond2<-(get(paste("alarm_",name,sep=''))==0 & outbreak_mean_cutoff==0)
        cond3<-(get(paste("alarm_",name,sep=''))==0 & outbreak_mean_cutoff==1)
        
        
        
        correct_alarm<-ifelse(cond1,1,0)
        no_alarm_no_outbreak<-ifelse(cond2,1,0)
        missed_outbreak<-ifelse(cond3,1,0)
        
        cond4<-(correct_alarm==0 & get(paste("alarm_",name,sep=''))==1)
        false_alarm<-ifelse(cond4,1,0)
        excess_cases<-((outbreak-outbreak_moving_limit)*get(population.a)) /1000
        
        excess_cases<-ifelse(excess_cases<0,0,excess_cases)
        
        #cases_below_threshold<-NA
        #cases_below_threshold<-ifelse(excess_cases==0,(outbreak*get(population()) /1000),NA)
        #cases_below_threshold<-ifelse(excess_cases!=0 & !is.na(excess_cases),(outbreak_moving_limit*get(population()) /1000),NA)
        
        
        
        weeks<-1
        total_cases<-(outbreak*get(population.a)) / 1000
        alarm_threshold<-alarm_threshold
        
        rm(cond,cond1,cond2,cond3,cond4)
      })
      ## try case when 2017-11-20
      data1<-data1 %>% mutate(cases_below_threshold=
                                case_when(excess_cases==0~(outbreak*get(population.a) /1000),
                                          excess_cases!=0 & !is.na(excess_cases)~(outbreak_moving_limit*get(population.a) /1000),
                                          TRUE~as.double(NA)
                                ))
      
      dat_lab3<-paste("Evaluation_Data(R)_",format.Date(Sys.Date(),"%d%b%Y"),'.dta',sep="")
      write.dta(data1,dat_lab3)
      
      get_eval_res<-function(d){
        rstats<-data1 %>% filter(district==run_per_district[d]) %>% summarise(n_weeks=sum(weeks),
                                                                              n_outbreak_weeks= sum(outbreakweek),
                                                                              n_outbreak_periods=sum(outbreakperiod,na.rm =T),
                                                                              n_alarms=with(data1,sum(get(paste("alarm_",name,sep='')))),
                                                                              n_correct_alarms=sum(correct_alarm),
                                                                              n_false_alarms=sum(false_alarm),
                                                                              n_missed_outbreaks=sum(missed_outbreak),
                                                                              n_no_alarm_no_outbreak=sum(no_alarm_no_outbreak),
                                                                              all_cases =sum(total_cases),
                                                                              n_outbreak_mean_cutoff=sum(outbreak_mean_cutoff),
                                                                              n_cases_below_threshold=sum(cases_below_threshold,na.rm =T))
        rstats1<-rstats %>% select(n_weeks,n_outbreak_weeks,n_outbreak_periods,n_outbreak_mean_cutoff,n_alarms,n_correct_alarms,n_false_alarms,n_missed_outbreaks,n_no_alarm_no_outbreak,all_cases,n_cases_below_threshold)
        StatTotal<-cbind(run_per_district[d],rstats1)
        return(StatTotal)
        
      }
      
      StatTotal<-foreach(d =1:length(run_per_district),.combine=rbind) %do% get_eval_res(d)
      
      
      
      mat_nam<-c("district" ,"weeks" ,"outbreak_weeks" ,"outbreak_periods" ,"defined_outbreaks", "alarms",  
                 "correct_alarms","false_alarms" ,"missed_outbreaks", "no_alarm_no_outbreak", "all_cases" ,"cases_below_threshold")
      colnames(StatTotal)<-mat_nam
      StatTotal<-StatTotal %>% mutate(sensitivity=correct_alarms/(defined_outbreaks),
                                      PPV=correct_alarms/(correct_alarms+false_alarms))
      cat("Evaluation Result")
      StatTotal
      write.dta(data.frame(StatTotal),paste("Evaluation_Result_R.dta",sep=""))
      
      
      
      
      ## do plots
      
      if(graph_per_district==TRUE){
        
        for(j in 1:length(run_per_district)){
          dat<-data1 %>% filter(district==run_per_district[j])
          #replace outbreak_moving_limit=0 if outbreak_moving_limit==.
          dat$outbreak_moving_limit<-ifelse(is.na(dat$outbreak_moving_limit),0,dat$outbreak_moving_limit)
          endmic<-paste("Endemic Channel (z outbreak=",z_outbreak,')',sep="")
          #yl1<-max(dat$outbreak,na.rm=T)+max(dat$outbreak,na.rm=T)/10
          yl1<-max(dat$outbreak,na.rm=T)*1.2
          #obj0<-xyplot(outbreak_moving_limit~obs_no,dat,panel=panel.xyarea,origin=0,border=F,col="lightblue")
          obj0<-xyplot(outbreak_moving_limit~obs_no,dat,panel=panel.xyarea,origin=0,border=F,col="lightblue",ylim=c(0,yl1))
          obj1<-xyplot(outbreak ~ obs_no,dat, type = "l",col="tomato2",lwd=2,grid=T)
          obj2<-xyplot(outbreak ~ obs_no,subset(dat,dat$outbreakperiod==1),
                       par.settings = list(plot.symbol=list(col = "red",cex = 0.8,pch = 16)),type = "p")
          objx<-obj0+obj1+obj2
          obj3<-xyplot(prob_outbreak ~ obs_no,dat,type='l',col="darkgreen",lwd=2)
          obj4<-xyplot(prob_outbreak ~ obs_no,subset(dat,with(dat,get(paste("alarm_",name,sep=''))==1)),
                       par.settings = list(plot.symbol=list(col = "blue", cex = 0.8, pch = 16)),type = "p")
          
          obj5<-xyplot(rep(alarm_threshold,nrow(dat)) ~ obs_no,dat,type='l',col="black",lwd=2)
          objy<-obj3+obj4+obj5
          max_ax<-round(max(c(max(dat$prob_outbreak,na.rm=T),alarm_threshold))+0.03,3)
          oby.1<-update(objy,scales=list(y=list(at=c(0,max_ax))),ylim=c(0,max_ax))
          #oby.1<-update(objy,scales=list(y=list(at=c(-0.88,1.12))),ylim=c(-0.88,1.12))
          #oby.1<-objy
          objx.1<-update(objx,main=paste('Evaluation Period District:',run_per_district[j]))
          all<-doubleYScale(objx.1,oby.1)
          plot_b<-update(all,
                         ylab.right="Probability of outbreak period",
                         ylab="Cases per 1000 pop",
                         xlab="Epidemiological week",
                         grid=T,
                         key=list(space="bottom",text=list(text=c("Cases per 1000 pop","Outbreak period",endmic,"Probability of outbreak period","Alarm signal","alarm_threshold")),
                                  columns=3,
                                  lines=list(lty=c(1,0,1,1,0,1),col=c("tomato2",NA,"lightblue","darkgreen",NA,"black"), lwd=c(2,NA,3,2,NA,2)),
                                  #lines=list(lty=c(1,1,0,1,0,1),col=c("lightblue","tomato2",NA,"darkgreen",NA,"black"), lwd=c(2,3,NA,2,NA,2)),
                                  points=list(pch=c(NA,16,NA,NA,16,NA),col=c(NA,"red",NA,NA,"blue",NA)),cex=0.7
                         ))
          #assign(paste("plot_b",j,sep=''),plot_b)
          
          print(plot_b)
          
          png(paste('Evaluation_period_for_district(',run_per_district[j],')_',rand_n,'-',format(Sys.Date(),"%d%b%y"),'.png',sep=''),width=20,
              height=12,units='cm',res=300)
          print(plot_b)
          dev.off()
          
          #to_fold()<-"C:\\joacim\\How To Guide-Demo Materials\\ewars app\\www"
          img_nam<-paste('Evaluation_period_for_district(',run_per_district[j],')_',rand_n,'-',format(Sys.Date(),"%d%b%y"),'.png',sep='')
          unlink(paste(to_fold,'/',img_nam,sep=''))
          file.copy(img_nam,to_fold,overwrite=T)
          
        }
      }
      
      
      
      
      
      
      ## plot whole time period
      data_for_gra1<-data %>% filter(runin==1)
      data_for_graph<-bind_rows(data_for_gra1,data1)
      data_for_graph<-within(data_for_graph,{
        ##replace prob_outbreak=exp(`probability_equation')/(1+exp(`probability_equation')) if runin==1
        if (spline==FALSE){
          for_eq<-rep(NA,(n_alarm_indicators+1))
          for(j in 1:(n_alarm_indicators+1)){
            if(j==1){
              for_eq[j]<-"coef1"  
            }
            else if(j>1){
              for_eq[j]<-paste("(mean_alarm",j-1,'*',"coef",j,')',sep='')
            }
            
          }
        }
        
        if (spline==TRUE){
          for_eq<-rep(NA,((n_alarm_indicators*4)+1))
          for(j in 1:((n_alarm_indicators*4)+1)){
            if(j==1){
              for_eq[j]<-"coef1"  
            }
            else if(j>1){
              for_eq[j]<-paste('(',independ_var_name[j-1],'*',"coef",j,')',sep='')
            }
            
          }
        }
        
        probability_equation<-eval(parse(text=paste(for_eq,collapse='+')))
        
        prob_outbreak1<-exp(probability_equation)/(1+exp(probability_equation))
        prob_outbreak[runin==1]<-prob_outbreak1[runin==1]
        rm(for_eq,j,probability_equation)
      })
      
      write.dta(data_for_graph,paste("data_for_graph_r(",run_per_district[d],").dta",sep=""))
      if(graph_per_district==TRUE){
        
        for(j in 1:length(run_per_district)){
          dat<-data_for_graph %>% filter(district==run_per_district[j])
          #replace outbreak_moving_limit=0 if outbreak_moving_limit==.
          dat$outbreak_moving_limit<-ifelse(is.na(dat$outbreak_moving_limit),0,dat$outbreak_moving_limit)
          
          endmic<-paste("Endemic Channel (z outbreak=",z_outbreak,')',sep="")
          #yl1<-max(dat$outbreak,na.rm=T)+max(dat$outbreak,na.rm=T)/10
          yl1<-max(dat$outbreak,na.rm=T)*1.2
          obj0<-xyplot(outbreak_moving_limit~obs_no,dat,panel=panel.xyarea,origin=0,border=F,col="lightblue",ylim=c(0,yl1))
          #obj0<-xyplot(outbreak_moving_limit~obs_no,dat,panel=panel.xyarea,origin=0,border=F,col="lightblue")
          obj1<-xyplot(outbreak ~ obs_no,dat, type = "l",col="tomato2",lwd=2,grid=T)
          obj2<-xyplot(outbreak ~ obs_no,subset(dat,dat$outbreakperiod==1),
                       par.settings = list(plot.symbol=list(col = "red",cex = 0.8,pch = 16)),type = "p")
          objx<-obj0+obj1+obj2
          obj3<-xyplot(prob_outbreak ~ obs_no,dat,type='l',col="darkgreen",lwd=2)
          obj4<-xyplot(prob_outbreak ~ obs_no,subset(dat,with(dat,get(paste("alarm_",name,sep=''))==1)),
                       par.settings = list(plot.symbol=list(col = "blue", cex = 0.8, pch = 16)),type = "p")
          
          obj5<-xyplot(rep(alarm_threshold,nrow(dat)) ~ obs_no,dat,type='l',col="black",lwd=2)
          objx.1<-update(objx,main=paste('Runin Evaluation Period District:',run_per_district[j]))
          objy<-obj3+obj4+obj5
          ## get max between threshold and probabilty out outbreak 04 January 2018
          max_ax<-round(max(c(max(dat$prob_outbreak,na.rm=T),alarm_threshold))+0.03,3)
          objy.1<-update(objy,scales=list(y=list(at=c(0,max_ax))),ylim=c(0,max_ax))
          #objy.1<-objy
          all<-doubleYScale(objx.1,objy.1)
          plot_c<-update(all,
                         ylab.right="Probability of outbreak period",
                         ylab="Cases per 1000 pop",
                         xlab="Epidemiological week",
                         grid=T,
                         key=list(space="bottom",text=list(text=c("Cases per 1000 pop","Outbreak period",endmic,"Probability of outbreak period","Alarm signal","alarm_threshold")),
                                  columns=3,
                                  lines=list(lty=c(1,0,1,1,0,1),col=c("tomato2",NA,"lightblue","darkgreen",NA,"black"), lwd=c(2,NA,3,2,NA,2)),
                                  points=list(pch=c(NA,16,NA,NA,16,NA),col=c(NA,"red",NA,NA,"blue",NA)),cex=0.7
                         ))
          
          #assign(paste("plot_c",j,sep=''),plot_c)
          
          #print(get(paste("plot_c",j,sep='')))
          
          png(paste('Runin_evaluation_district(',run_per_district[j],')_',rand_n,'-',format(Sys.Date(),"%d%b%y"),'.png',sep=''),width=20,
              height=12,units='cm',res=300)
          print(plot_c)
          dev.off()
          img_nam<-paste('Runin_evaluation_district(',run_per_district[j],')_',rand_n,'-',format(Sys.Date(),"%d%b%y"),'.png',sep='')
          
          #unlink(paste(to_fold(),'/',img_nam,sep=''))
          
          file.copy(img_nam,to_fold,overwrite=T)
          
        }
        
      }
      
    }
    
    #**********************Analysis and Exporting Results************************
    text<-"A?B?C?D?E?F?G?H?I?J?K?L?M?N?O?P?Q?R?S?T?U?V?W?X?Y?Z?ABACADAEAFAGAHAIAJAKALAMANAOAPAQARASATAUAVAWAXAYAZBABBBCBDBEBFBGBHBIBJBKBLBMBNBOBPBQBRBSBTBUBVBWBXBYBZ"
    cell_row<-13
    if (generating_surveillance_workbook==TRUE){
      distr_matrix=run_per_district
      n_cols=length(run_per_district)
      for (j in 1:n_cols){
        
        fname<-paste("Surveillance workbook_R",run_per_district[j],".xlsx",sep='')
        wb_template<-"Surveillance workbook_template.xlsx"
        unlink(fname)
        wb <- XLConnect::loadWorkbook(wb_template,create =F)
        XLConnect::createSheet(wb,name="Sheet")
        cs <- XLConnect::createCellStyle(wb)
        XLConnect::setFillPattern(cs, fill = XLC$"FILL.SOLID_FOREGROUND")
        XLConnect::setFillForegroundColor(cs,color =XLC$"COLOR.LIGHT_YELLOW")
        ## use "Parameters_runin.dta", clear
        Parameters_runin<-read.dta(paste("Parameters_runin(R)_",format.Date(Sys.Date(),"%d%b%Y"),'.dta',sep=""))
        ##egen tag=tag(week district)
        ##keep if tag==1
        dat<-with(Parameters_runin,subset(Parameters_runin,district==run_per_district[j]))
        dat<-dat %>% filter(year==(min(dat$year)+1))
        dat<-dat %>% arrange(year_week)
        
        if(spline==1){
          
          for(n in 1:n_alarm_indicators){
            centile<- with(dat,as.numeric(quantile(get(paste("mean_alarm",n,sep='')),c(5,27.5,50,72.5,95)/100)))
            ## save the contents into objects
            
            for (i in 1:5){
              assign(paste('t',i,'_mean_alarm',n,sep=''),centile[i])
            }
            
            mean_<-with(dat,as.numeric(mean(get(paste("mean_alarm",n,sep='')),rm.na=T)))
            N<-nrow(dat)
            cenmin <- 5/(N+1)
            cenmax <-(N-4)/(N+1)
            
            centile1<- with(dat,as.numeric(quantile(get(paste("mean_alarm",n,sep='')),c(cenmin,cenmax))))
            
            min.a<-centile1[1]
            max.a<-centile1[2]
            
            t1.a<-get(paste('t',1,'_mean_alarm',n,sep=''))
            t5.a<-get(paste('t',5,'_mean_alarm',n,sep=''))
            
            if(t1.a>min.a){
              assign('t1',t1.a)
            }
            
            if(t5.a>max.a){
              assign('t5',t5.a)
            }
            
            Reference_alarm_indicator<-as.matrix(dat[,paste("mean_alarm",1:n_alarm_indicators,sep='')])
            n_obs_ref=nrow(Reference_alarm_indicator)
            
          }
          
        }
        
        if(spline==FALSE){
          cell_1<-(n_alarm_indicators+16)*2-1
          cell_2<-(n_alarm_indicators+15+n_alarm_indicators+6)*2-1
          cell_3<-(n_alarm_indicators+20)*2-1
          
          
          Output_matrix_cell<-substr(text,cell_1,cell_1+2)
          if(substr(Output_matrix_cell,2,2)=="?"){
            Output_matrix_cell<-substr(Output_matrix_cell,1,1)
          }
          
          intercept_cell<-substr(text,cell_2,cell_2+2)
          if(substr(intercept_cell,2,2)=="?"){
            intercept_cell<-substr(intercept_cell,1,1)
          }
          
          
          alarm_probability_cell="E"
          probability_cut_off_cell="F"
          outbreak_week_cell="G"
          alarm_week_cell="H"
          
          
          moving_limit_cell<-substr(text,cell_3,cell_3+2)
          if(substr(moving_limit_cell,2,2)=="?"){
            moving_limit_cell<-substr(moving_limit_cell,1,1)
          }
          
          
          col_label=cell_row-1
          jj<-0
          
          for(i in al_var){
            jj=jj+1
            cell_num<-(10+jj)*2-1
            
            alarm_label_cell<-substr(text,cell_num,cell_num+2)
            if(substr(alarm_label_cell,2,2)=="?"){
              alarm_label_cell<-substr(alarm_label_cell,1,1)
              XLConnect::createName(wb, name ="reg", formula = paste("Sheet!",alarm_label_cell,col_label,sep=''),overwrite =T)
              XLConnect::writeNamedRegion(wb,i, name = "reg",header =F)
            }
          }
          
          
          #mkmat district week outbreak_moving outbreak_moving_sd outbreak_moving_limit coef*, matrix(Output_matrix) 
          
          names_mat<-c(c("district","week","outbreak_moving","outbreak_moving_sd","outbreak_moving_limit"),paste("coef",2:(n_alarm_indicators+1),sep=''),"coef1")
          dat_put<-dat[,names_mat]
          
          XLConnect::createName(wb, name ="reg1", formula = paste("Sheet!",Output_matrix_cell,col_label,sep=''),overwrite =T)
          XLConnect::writeNamedRegion(wb,dat_put, name = "reg1",header =T)
          
          ## update the first column
          
          
          a<-c("District :" ,"Alarm indicator :","z_outbreak :","Prediction distance:","Outbreak window size:","Alarm window size:","outbreak_threshold:",
               "alarm_threshold:","Outbreak week length:")
          a1<-c("District:" ,"Alarm indicators:","z_outbreak:","Prediction distance:","Outbreak window:","Alarm window:","outbreak_threshold:",
                "alarm_threshold:","Outbreak week length:")
          
          a2<-gsub(" ",'_',tolower(gsub(':',"",a1)))
          
          mat_initial<-matrix(NA,length(a),2)
          mat_initial[,1]<-a
          for(i in 1:length(a)){
            if(i==1){
              mat_initial[i,2]<-run_per_district[j]
            }
            if(i>1){
              mat_initial[i,2]<-paste(eval(parse(text=a2[i])),collapse=" ")
            }
            
          }
          
          XLConnect::createName(wb, name ="reg2", formula = paste("Sheet!","A",1,sep=''),overwrite =T)
          XLConnect::writeNamedRegion(wb,mat_initial, name = "reg2",header =F)
          
          new_vars<-c("Year","Week","Outbreak indicator","Endemic Channel","Outbreak probability","alarm_threshold",
                      "Outbreak period","Alarm signal",number_of_cases,"population")
          ## get the positions
          pos<-c("A","B","C","D",alarm_probability_cell,probability_cut_off_cell,outbreak_week_cell,alarm_week_cell,"I","J")
          
          
          # populate these column names into excel
          
          XLConnect::createName(wb, name ="reg3", formula = paste("Sheet!","A",col_label,sep=''),overwrite =T)
          XLConnect::writeNamedRegion(wb,t(new_vars), name = "reg3",header =F)
          
          cell_4<-(n_alarm_indicators+15+n_alarm_indicators+6)
          cell_5<-(10+n_alarm_indicators+1)
          
          loop_end=cell_row+50
          
          
          for (cell in cell_row:loop_end){
            
            for_eq<-rep(NA,n_alarm_indicators)
            
            for(n in 1:n_alarm_indicators){
              
              alarm_coef_cell<-substr(text,(cell_4-n)*2-1,((cell_4-n)*2-1)+2)
              if(substr(alarm_coef_cell,2,2)=="?"){
                alarm_coef_cell<-substr(alarm_coef_cell,1,1)
              }
              
              alarm_cell<-substr(text,(cell_5-n)*2-1,((cell_5-n)*2-1)+2)
              if(substr(alarm_cell,2,2)=="?"){
                alarm_cell<-substr(alarm_cell,1,1)
              }
              for_eq[n]<-paste(alarm_coef_cell,cell,'*',alarm_cell,cell,sep='')
              ## put background on the alarm cell
              
              XLConnect::setCellStyle(wb, sheet = "Sheet", row =cell, col = col2idx(alarm_cell), cellstyle = cs)
              
            }
            probability_equation<-paste(c(paste(intercept_cell,cell,sep=''),for_eq),collapse = "+")
            prob_outbreak_formula<-paste('exp(',probability_equation,')/(1+exp(',probability_equation,'))',sep='')
            
            ##putexcel `alarm_probability_cell'`cell' = formula("`prob_outbreak_formula'")
            
            XLConnect::setCellFormula(wb, "Sheet", row=cell, col=col2idx(alarm_probability_cell), prob_outbreak_formula)
            ##putexcel D`cell' = formula("`moving_limit_cell'`cell'")
            XLConnect::setCellFormula(wb, "Sheet", row=cell, col=col2idx("D"), paste(moving_limit_cell,cell,sep=''))
            ##putexcel `probability_cut_off_cell'`cell' = formula(B8)
            XLConnect::setCellFormula(wb, "Sheet", row=cell, col=col2idx(probability_cut_off_cell), paste("B",8,sep=''))
            ##putexcel C`cell' = formula((I`cell'/J`cell')*1000)
            XLConnect::setCellFormula(wb, "Sheet", row=cell, col=col2idx("C"), paste("(I",cell,'/','J',cell,")*1000",sep=''))
            #putexcel `outbreak_week_cell'`cell' = formula("IF(C`cell'>D`cell';1;0)") 
            XLConnect::setCellFormula(wb, "Sheet", row=cell, col=col2idx(outbreak_week_cell), paste("IF(C",cell,'>D',cell,',1,0)',sep=''))
            #putexcel `alarm_week_cell'`cell' = formula("IF(`alarm_probability_cell'`cell'>`probability_cut_off_cell'`cell';1;0)")
            XLConnect::setCellFormula(wb, "Sheet", row=cell, col=col2idx(alarm_week_cell), paste("IF(",alarm_probability_cell,cell,'>',probability_cut_off_cell,cell,',1,0)',sep=''))
            #putexcel B`cell' = ((`cell'-`=cell_row')+2)
            XLConnect::setCellFormula(wb, "Sheet", row=cell, col=col2idx("B"), paste('(',cell,'-',cell_row,')+2',sep=''))
            
            #putexcel I`cell' =(""), fpattern("solid","242 220 219", "")
            XLConnect::setCellStyle(wb, sheet = "Sheet", row =cell, col = col2idx("I"), cellstyle = cs)
            #putexcel J`cell' =(""), fpattern("solid","242 220 219", "")
            
            XLConnect::setCellStyle(wb, sheet = "Sheet", row =cell, col = col2idx("J"), cellstyle = cs)
            
          }
          
          
          
        }
        
        
        ## add for spline()==1
        if(spline==TRUE){
          cella.1<-(1*n_alarm_indicators+16)*2-1
          cella.2<-(1*n_alarm_indicators+21+n_alarm_indicators*4)*2-1
          cella.3<-(1*n_alarm_indicators+20)*2-1
          
          Output_matrix_cell<-substr(text,cella.1,cella.1+2)
          if(substr(Output_matrix_cell,2,2)=="?"){
            Output_matrix_cell<-substr(Output_matrix_cell,1,1)
          }
          
          intercept_cell<-substr(text,cella.2,cella.2+2)
          if(substr(intercept_cell,2,2)=="?"){
            intercept_cell<-substr(intercept_cell,1,1)
          }
          
          
          alarm_probability_cell="E"
          probability_cut_off_cell="F"
          outbreak_week_cell="G"
          alarm_week_cell="H"
          
          moving_limit_cell<-substr(text,cella.3,cella.3+2)
          if(substr(moving_limit_cell,2,2)=="?"){
            moving_limit_cell<-substr(moving_limit_cell,1,1)
          }
          
          col_label<-cell_row-1
          jj<-1
          pp<-1
          
          for(i in al_var){
            cell_num<-(10+jj)*2-1
            alarm_label_cell<-substr(text,cell_num,cell_num+2)
            if(substr(alarm_label_cell,2,2)=="?"){
              alarm_label_cell<-substr(alarm_label_cell,1,1)
            }
            jj=jj+1
            for(sp in 1:4){
              cell_num_sp<-(22+(1*n_alarm_indicators)+(4*n_alarm_indicators)+pp)*2-1
              
              spline_label_cell<-substr(text,cell_num_sp,cell_num_sp+2)
              if(substr(spline_label_cell,2,2)=="?"){
                spline_label_cell<-substr(spline_label_cell,1,1)
                createName(wb, name ="reg", formula = paste("Sheet!",spline_label_cell,col_label,sep=''),overwrite =T)
                writeNamedRegion(wb,paste(i,"_spline",sp,sep=''), name = "reg",header =F)
              }
              pp<-pp+1
              createName(wb, name ="reg", formula = paste("Sheet!",alarm_label_cell,col_label,sep=''),overwrite =T)
              writeNamedRegion(wb,i, name = "reg",header =F)
              
            }
            
          }
          
          
          
          #mkmat district week outbreak_moving outbreak_moving_sd outbreak_moving_limit coef*, matrix(Output_matrix) 
          names_mat<-c(c("district","week","outbreak_moving","outbreak_moving_sd","outbreak_moving_limit"),paste("coef",2:((n_alarm_indicators*4)+1),sep=''),"coef1")
          dat_put<-dat[,names_mat]
          XLConnect::createName(wb, name ="reg1", formula = paste("Sheet!",Output_matrix_cell,col_label,sep=''),overwrite =T)
          XLConnect::writeNamedRegion(wb,dat_put, name = "reg1",header =T)
          
          a<-c("District :" ,"Alarm indicator :","z_outbreak:","Prediction distance:","Outbreak window size:","Alarm window size:","outbreak_threshold:",
               "alarm_threshold:","Outbreak week length:")
          a1<-c("District:" ,"Alarm indicators:","z_outbreak:","Prediction distance:","Outbreak window:","Alarm window:","outbreak_threshold:",
                "alarm_threshold:","Outbreak week length:")
          
          a2<-gsub(" ",'_',tolower(gsub(':',"",a1)))
          
          mat_initial<-matrix(NA,length(a),2)
          mat_initial[,1]<-a
          for(i in 1:length(a)){
            if(i==1){
              mat_initial[i,2]<-run_per_district[j]
            }
            if(i>1){
              mat_initial[i,2]<-paste(eval(parse(text=a2[i])),collapse=" ")
            }
            
          }
          
          XLConnect::createName(wb, name ="reg2", formula = paste("Sheet!","A",1,sep=''),overwrite =T)
          XLConnect::writeNamedRegion(wb,mat_initial, name = "reg2",header =F)
          
          new_vars<-c("Year","Week","Outbreak indicator","Endemic Channel","Outbreak probability","alarm_threshold",
                      "Outbreak period","Alarm signal",number_of_cases,"population")
          
          ## get the positions
          pos<-c("A","B","C","D",alarm_probability_cell,probability_cut_off_cell,outbreak_week_cell,alarm_week_cell,"I","J")
          
          # populate these column names into excel
          
          XLConnect::createName(wb, name ="reg3", formula = paste("Sheet!","A",col_label,sep=''),overwrite =T)
          XLConnect::writeNamedRegion(wb,t(new_vars), name = "reg3",header =F)
          
          
          
          loop_end<-cell_row+50
          u<-0
          for (cell in cell_row:loop_end){
            m<-1
            k<-1
            cell_position<-cell_row+n_obs_ref+u
            u<-u+1
            for_eq<-rep(NA,(n_alarm_indicators*4))
            
            for(n in 1:n_alarm_indicators){
              alarm_cell<-substr(text,((10+n)*2-1),((10+n)*2-1)+2)
              if(substr(alarm_cell,2,2)=="?"){
                alarm_cell<-substr(alarm_cell,1,1)
              }
              
              for(sp in 1:4){
                
                alarm_coef_cell<-substr(text,((1*n_alarm_indicators+20+k)*2-1),((1*n_alarm_indicators+20+k)*2-1)+2)
                if(substr(alarm_coef_cell,2,2)=="?"){
                  alarm_coef_cell<-substr(alarm_coef_cell,1,1)
                }
                
                spline_cell<-substr(text,(((1*n_alarm_indicators)+22+(n_alarm_indicators*4)+k)*2-1),(((1*n_alarm_indicators)+22+(n_alarm_indicators*4)+k)*2-1)+2)
                if(substr(spline_cell,2,2)=="?"){
                  spline_cell<-substr(spline_cell,1,1)
                }
                nk<-5
                km1<-nk-1
                t1<-get(paste('t1_mean_alarm',n,sep=''))
                t2<-get(paste('t2_mean_alarm',n,sep=''))
                t3<-get(paste('t3_mean_alarm',n,sep=''))
                t4<-get(paste('t4_mean_alarm',n,sep=''))
                t5<-get(paste('t5_mean_alarm',n,sep=''))
                
                a_Xm1p<-paste("IF(",paste(alarm_cell,cell,sep=''),'-',t1,'<0,0,',paste(alarm_cell,cell,sep=''),'-',t1,")",sep="")
                a_Xm2p<-paste("IF(",paste(alarm_cell,cell,sep=''),'-',t2,'<0,0,',paste(alarm_cell,cell,sep=''),'-',t2,")",sep="")
                a_Xm3p<-paste("IF(",paste(alarm_cell,cell,sep=''),'-',t3,'<0,0,',paste(alarm_cell,cell,sep=''),'-',t3,")",sep="")
                a_Xm4p<-paste("IF(",paste(alarm_cell,cell,sep=''),'-',t4,'<0,0,',paste(alarm_cell,cell,sep=''),'-',t4,")",sep="")
                a_Xm5p<-paste("IF(",paste(alarm_cell,cell,sep=''),'-',t5,'<0,0,',paste(alarm_cell,cell,sep=''),'-',t5,")",sep="")
                
                
                if (sp==1){
                  spline_formula=paste(alarm_cell,cell,sep='')
                }
                if (sp==2){
                  spline_formula<-paste("((",a_Xm1p,')^3-((',get(paste('a_Xm',km1,'p',sep='')),'^3)*(',get(paste('t',nk,sep='')),'-',t1,')/(',get(paste('t',nk,sep='')),'-',get(paste('t',km1,sep='')),')+((',get(paste('a_Xm',nk,'p',sep='')),')^3)*(',get(paste('t',km1,sep='')),'-',t1,')/(',get(paste('t',nk,sep='')),'-',get(paste('t',km1,sep='')),'))/(',get(paste('t',nk,sep='')),'-',t1,')^2)',sep='')
                }
                
                if (sp==3){
                  spline_formula<-paste("((",a_Xm1p,')^3-((',get(paste('a_Xm',km1,'p',sep='')),'^3)*(',get(paste('t',nk,sep='')),'-',t2,')/(',get(paste('t',nk,sep='')),'-',get(paste('t',km1,sep='')),')+((',get(paste('a_Xm',nk,'p',sep='')),')^3)*(',get(paste('t',km1,sep='')),'-',t2,')/(',get(paste('t',nk,sep='')),'-',get(paste('t',km1,sep='')),'))/(',get(paste('t',nk,sep='')),'-',t1,')^2)',sep='')
                }
                
                if (sp==4){
                  spline_formula<-paste("((",a_Xm1p,')^3-((',get(paste('a_Xm',km1,'p',sep='')),'^3)*(',get(paste('t',nk,sep='')),'-',t3,')/(',get(paste('t',nk,sep='')),'-',get(paste('t',km1,sep='')),')+((',get(paste('a_Xm',nk,'p',sep='')),')^3)*(',get(paste('t',km1,sep='')),'-',t3,')/(',get(paste('t',nk,sep='')),'-',get(paste('t',km1,sep='')),'))/(',get(paste('t',nk,sep='')),'-',t1,')^2)',sep='')
                }
                
                
                setCellFormula(wb, "Sheet", row=cell, col=col2idx(spline_cell), spline_formula)
                
                for_eq[k]<-paste(alarm_coef_cell,cell,'*',spline_cell,cell,sep='')
                
                
                
                k=k+1
              }
              
              probability_equation<-paste(c(paste(intercept_cell,cell,sep=''),for_eq),collapse = "+")
              prob_outbreak_formula<-paste('exp(',probability_equation,')/(1+exp(',probability_equation,'))',sep='')
              XLConnect::setCellStyle(wb, sheet = "Sheet", row =cell, col = col2idx(alarm_cell), cellstyle = cs)
              
              
            }
            ##putexcel `alarm_probability_cell'`cell' = formula("`prob_outbreak_formula'")
            
            XLConnect::setCellFormula(wb, "Sheet", row=cell, col=col2idx(alarm_probability_cell), prob_outbreak_formula)
            ##putexcel D`cell' = formula("`moving_limit_cell'`cell'")
            XLConnect::setCellFormula(wb, "Sheet", row=cell, col=col2idx("D"), paste(moving_limit_cell,cell,sep=''))
            ##putexcel `probability_cut_off_cell'`cell' = formula(B8)
            XLConnect::setCellFormula(wb, "Sheet", row=cell, col=col2idx(probability_cut_off_cell), paste("B",8,sep=''))
            ##putexcel C`cell' = formula((I`cell'/J`cell')*1000)
            XLConnect::setCellFormula(wb, "Sheet", row=cell, col=col2idx("C"), paste("(I",cell,'/','J',cell,")*1000",sep=''))
            #putexcel `outbreak_week_cell'`cell' = formula("IF(C`cell'>D`cell';1;0)") 
            XLConnect::setCellFormula(wb, "Sheet", row=cell, col=col2idx(outbreak_week_cell), paste("IF(C",cell,'>D',cell,',1,0)',sep=''))
            #putexcel `alarm_week_cell'`cell' = formula("IF(`alarm_probability_cell'`cell'>`probability_cut_off_cell'`cell';1;0)")
            XLConnect::setCellFormula(wb, "Sheet", row=cell, col=col2idx(alarm_week_cell), paste("IF(",alarm_probability_cell,cell,'>',probability_cut_off_cell,cell,',1,0)',sep=''))
            #putexcel B`cell' = ((`cell'-`=cell_row')+2)
            XLConnect::setCellFormula(wb, "Sheet", row=cell, col=col2idx("B"), paste('(',cell,'-',cell_row,')+2',sep=''))
            #putexcel I`cell' =(""), fpattern("solid","242 220 219", "")
            XLConnect::setCellStyle(wb, sheet = "Sheet", row =cell, col = col2idx("I"), cellstyle = cs)
            #putexcel J`cell' =(""), fpattern("solid","242 220 219", "")
            
            XLConnect::setCellStyle(wb, sheet = "Sheet", row =cell, col = col2idx("J"), cellstyle = cs)
            
            
          }
          
        } 
        
        XLConnect::saveWorkbook(wb,fname) 
        #to_fold()<-"C:\\joacim\\How To Guide-Demo Materials\\ewars app\\www"
        
        file.copy(fname,to_fold,overwrite=T)
      }
      
    } 
    ## copy the excel files to www foldre
    
    
    len_d<-length(run_per_district)
    
    if((graph_per_district==TRUE & generating_surveillance_workbook==FALSE)){
      
      ## code to render images
      #img_nam<-paste('Runin_period_for_district(',run_per_district()[j],')_',format(Sys.Date(),"%d%b%y"),'.png',sep='')
      #img_nam<-paste('Evaluation_period_for_district(',run_per_district()[j],')_',format(Sys.Date(),"%d%b%y"),'.png',sep='')
      #img_nam<-paste('Runin_evaluation_district(',run_per_district()[j],')_',format(Sys.Date(),"%d%b%y"),'.png',sep='')
      
      ##"file:/D:/Images/TechnipLogo.jpg">
      imag_path<-"file:///C:/joacim/How To Guide-Demo Materials/ewars app/www/"
      
      img1.a<-vector(mod="character",length=len_d)
      img1.b<-vector(mod="character",length=len_d)
      img1.c<-vector(mod="character",length=len_d)
      
      for(i in 1: len_d){
        img_nam<-paste('Runin_period_for_district(',run_per_district[i],')_',rand_n,'-',format(Sys.Date(),"%d%b%y"),'.png',sep='')
        if(i!=len_d){
          img1.a[i]<-paste("tags$img(src='",img_nam,"', width = '800px', height = '500px'),br()",sep='')
        }else{img1.a[i]<-paste("tags$img(src='",img_nam,"', width = '800px', height = '500px')",sep='')}
      }
      
      img1.a1<-paste(img1.a,collapse=',')
      img1.a2<-paste("tags$div(class='header', checked=NA,",img1.a1,')',sep='')
      
      for(i in 1: len_d){
        img_nam<-paste('Evaluation_period_for_district(',run_per_district[i],')_',rand_n,'-',format(Sys.Date(),"%d%b%y"),'.png',sep='')
        if(i!=len_d){
          img1.b[i]<-paste("tags$img(src='",img_nam,"', width = '800px', height = '500px'),br()",sep='')
        }else{img1.b[i]<-paste("tags$img(src='",img_nam,"', width = '800px', height = '500px')",sep='')}
      }
      
      img1.b1<-paste(img1.b,collapse=',')
      img1.b2<-paste("tags$div(class='header', checked=NA,",img1.b1,')',sep='')
      
      for(i in 1: len_d){
        img_nam<-paste('Runin_evaluation_district(',run_per_district[i],')_',rand_n,'-',format(Sys.Date(),"%d%b%y"),'.png',sep='')
        if(i!=len_d){
          img1.c[i]<-paste("tags$img(src='",img_nam,"', width = '800px', height = '500px'),br()",sep='')
        }else{img1.c[i]<-paste("tags$img(src='",img_nam,"', width = '800px', height = '500px')",sep='')}
      }
      
      img1.c1<-paste(img1.c,collapse=',')
      img1.c2<-paste("tags$div(class='header', checked=NA,",img1.c1,')',sep='')
      
      
      sss<<-list(plot1.a=img1.a2,plot2.a=img1.b2,plot3.a=img1.c2,tab=StatTotal,xls_file=NULL)
      ret<- list(plot1.a=img1.a2,plot2.a=img1.b2,plot3.a=img1.c2,tab=StatTotal,xls_file=NULL)
      
    }
    
    if(generating_surveillance_workbook==TRUE & graph_per_district==FALSE){
      
      ##run_per_district()
      ds<-vector(mod="character",length=len_d)
      
      for(i in 1: len_d){
        fname<-paste("Surveillance workbook_R",run_per_district[i],".xlsx",sep='')
        if(i!=len_d){
          ds[i]<-paste("tags$a('",fname,"',target='_blank',href='",fname,"'),br()",sep='')
        }else{ds[i]<-paste("tags$a('",fname,"',target='_blank',href='",fname,"')",sep='')}
      }
      
      ds1<-paste(ds,collapse=',')
      ds2<-paste("tags$div(class='header', checked=NA,",ds1,')',sep='')
      
      sss<<-list(plot1.a=NULL,plot2.a=NULL,plot3.a=NULL,tab=NULL,xls_file=ds2)
      ret<-list(plot1.a=NULL,plot2.a=NULL,plot3.a=NULL,tab=NULL,xls_file=ds2)
      
    }
    
    if(generating_surveillance_workbook==TRUE & graph_per_district==TRUE){
      
      ##run_per_district()
      ds<-vector(mod="character",length=len_d)
      
      for(i in 1: len_d){
        fname<-paste("Surveillance workbook_R",run_per_district[i],".xlsx",sep='')
        if(i!=len_d){
          ds[i]<-paste("tags$a('",fname,"',target='_blank',href='",fname,"'),br()",sep='')
        }else{ds[i]<-paste("tags$a('",fname,"',target='_blank',href='",fname,"')",sep='')}
      }
      
      ds1<-paste(ds,collapse=',')
      ds2<-paste("tags$div(class='header', checked=NA,",ds1,')',sep='')
      
      img1.a<-vector(mod="character",length=len_d)
      img1.b<-vector(mod="character",length=len_d)
      img1.c<-vector(mod="character",length=len_d)
      
      for(i in 1: len_d){
        img_nam<-paste('Runin_period_for_district(',run_per_district[i],')_',format(Sys.Date(),"%d%b%y"),'.png',sep='')
        if(i!=len_d){
          img1.a[i]<-paste("tags$img(src='",img_nam,"', width = '800px', height = '500px'),br()",sep='')
        }else{img1.a[i]<-paste("tags$img(src='",img_nam,"', width = '800px', height = '500px')",sep='')}
      }
      
      img1.a1<-paste(img1.a,collapse=',')
      img1.a2<-paste("tags$div(class='header', checked=NA,",img1.a1,')',sep='')
      
      for(i in 1: len_d){
        img_nam<-paste('Evaluation_period_for_district(',run_per_district[i],')_',format(Sys.Date(),"%d%b%y"),'.png',sep='')
        if(i!=len_d){
          img1.b[i]<-paste("tags$img(src='",img_nam,"', width = '800px', height = '500px'),br()",sep='')
        }else{img1.b[i]<-paste("tags$img(src='",img_nam,"', width = '800px', height = '500px')",sep='')}
      }
      
      img1.b1<-paste(img1.b,collapse=',')
      img1.b2<-paste("tags$div(class='header', checked=NA,",img1.b1,')',sep='')
      
      for(i in 1: len_d){
        img_nam<-paste('Runin_evaluation_district(',run_per_district[i],')_',format(Sys.Date(),"%d%b%y"),'.png',sep='')
        if(i!=len_d){
          img1.c[i]<-paste("tags$img(src='",img_nam,"', width = '800px', height = '500px'),br()",sep='')
        }else{img1.c[i]<-paste("tags$img(src='",img_nam,"', width = '800px', height = '500px')",sep='')}
      }
      
      img1.c1<-paste(img1.c,collapse=',')
      img1.c2<-paste("tags$div(class='header', checked=NA,",img1.c1,')',sep='')
      
      sss<<-list(plot1.a=NULL,plot2.a=NULL,plot3.a=NULL,tab=NULL,xls_file=ds2)
      ret<- list(plot1.a=NULL,plot2.a=NULL,plot3.a=NULL,tab=NULL,xls_file=ds2)
      
    }
    
    if(generating_surveillance_workbook==FALSE & graph_per_district==FALSE){
      
      sss<<-list(plot1.a=NULL,plot2.a=NULL,plot3.a=NULL,tab=StatTotal,xls_file=NULL)
      ret<- list(plot1.a=NULL,plot2.a=NULL,plot3.a=NULL,tab=StatTotal,xls_file=NULL)
      
    }
    
    ret
    
    
    
  }) 
  

  output$plot1 <- renderUI({
    
    if(!is.null(out_x()$plot1.a))  
      eval( parse(text=out_x()$plot1.a))
    
  }) 
  output$plot2 <- renderUI({ 
    
    if(!is.null(out_x()$plot2.a))
      
      eval( parse(text=out_x()$plot2.a))
  })
  output$plot3 <- renderUI({
    
    if(!is.null(out_x()$plot3.a))
      
      eval( parse(text=out_x()$plot3.a))
  })
  output$table1 <- renderTable({
    
    if(!is.null(out_x()$tab))
      
      out_x()$tab
  })
  
  output$workbooks <- renderUI({
    
    if(!is.null(out_x()$xls_file))
      
      eval( parse(text=out_x()$xls_file))
    
  })
  
  
  
  
}
