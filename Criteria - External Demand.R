# ---------------------------------------------------------------------#
#  Purpose: Build Program Performance Matrix - External Demand Criteria
#  Requestee:
#
#  Created by: John Schliesmann
#  Created Date: 6/18/2019
#
#  Modifications:
#        1. as of 2019 this attempts to duplicate excel workbook created by Tom Walsh
#
#
#  Requirments:
#        3. Contact Ben Konruff for updated WTCS Program Code to 
#              Standard Occupational Classification (SOC) crosswalk
#        5. Update EMSI Occupation Outlook Tables: (www.economicmodeling.com)
#              filters: Intl. Employment opportunities 
#                       include new and replacement job openings 
#                       in all Milwaukee, Waukesha, Ozaukee, and Washington Counties
#              These are the same EMSI datasets used in Marketing'S Graduate Career Report
#              See `EMSI Occupation Table Filters.png` in documentation for detailed filters
#        3. Download flw500.xls - Graduate Outcomes Survey Report for MATC only.  
#               This file may not load if the report says "not final" (likely the case for most recent year)
#
#  Notes:
#        1. Placement rate for most recently completed fiscal year may not be 
#               complete until the following Jan (see Requirement #3)
#
#  ToDo:
#        1. build high level criteria comparison of all MATC programs
#        2. include apprenticship wages (APR500.xls).  is needed for Capital Project's Industry Impact Score
#
# ----------------------------------------------------------------------#

processExternalDemand <- function(){
#~~~~~~~~~~~~~~ USER EDITS ~~~~~~~~~~~~~~~~~~~~~####
#~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~####

# FISCAL_YEAR <- 2019
# FED_POVERTY <- 12490  #https://aspe.hhs.gov/poverty-guidelines last updated 8/2019 

FOLDER_ED <- paste0(FOLDER,'Criteria - External Demand\\')

#EMSI Occupational Outlook tables
EMSI_1 <- paste0(FOLDER_ED,'Occupation_Table-',FISCAL_YEAR,'.xls')
EMSI_2 <- paste0(FOLDER_ED,'Occupation_Table-',FISCAL_YEAR-1,'.xls')
EMSI_3 <- paste0(FOLDER_ED,'Occupation_Table-',FISCAL_YEAR+2,'.xls')

#WTCS Datasets
#WTCS <- paste0('FY',FISCAL_YEAR,'\\WTCS - Data Source.xlsx')
XWALK_FILE <- paste0(FOLDER_ED,'Cloud WTCS Program to SOC Crosswalk.xlsx')
#HIGH_DEMAND <- 'Top 50 List with Academic Programs 2018-19.xlsx'
FLW500_LIST <- list.files(FOLDER_ED,pattern = 'FLW500', full.names = T)  #download from portal 

#~~~~~~~~~~~~~~ Project Setup ~~~~~~~~~~~~~~~~~~####
#~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~####

# library(tidyverse)
# library(stringr)
# #library(tidyquant)
# #library(scales)
# library(zoo)
# library(readxl)
# #library(openxlsx)
# library(XLConnect)
# 
# options(scipen=999)
# 
# source('S:\\RESEARCH\\55_John Schliesmann\\MATC Custom Functions.R')
# source('S:\\RESEARCH\\55_John Schliesmann\\Client Reporting Functions.R')

#setwd(paste0('S:\\RESEARCH\\06_Program Evaluation\\R_Performance_Matrix\\FY',FISCAL_YEAR))

#~~~~~~~~~~~~~~ Custom Functions ~~~~~~~~~~~~~~~####
#~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~####

rollQuartile <- function(x,col_name,period,pct = .75){
  #val<-eval(substitute(val),data, parent.frame())
  #val <- as.name(val)
  col_name <- ensym(col_name)

  p <- as.integer(period)
  p1 <-as.integer(period) - 1
  p2 <-as.integer(period) - 2
  
  x2 <- x %>%
    filter(district_group == 'WTCS',
           fiscal_year %in% c(p,p1,p2)) %>%
    filter({if("employed_related_cnt" %in% names(.)) employed_related_cnt else 6} > 5) %>%  #this excludes programs with > 5 respondents
    mutate(val = !!enquo(col_name))
  
  q <- quantile(x2$val,probs = pct,na.rm = TRUE)
  return(q)
}

replaceAllNa <- function(x){
  #Arguments: X is a dataframe
  #USAGE:   %>% replaceAllNa() 
  df <- x %>%
    mutate_if(is.integer, replace_na, 0) %>%
    mutate_if(is.numeric, replace_na,  0) %>%
    mutate_if(is.character, replace_na, '') 
  return(df)
}
formatXl <- function(x) {
    names(x)  <- tolower(names(x))
    names(x) <- gsub(" ", "_", names(x))
    names(x) <- gsub("\\.", "_", names(x))
    return(x)
}

#~~~~~~~~~~~~~~ EXTERNAL DEMAND ~~~~~~~~~~~~~~~~####
#~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~####

#---- Get Placement Rate ----####

    #This process will differ from FLW500 for WTCS district because it will aggerate all employed 
    #  and in related field before percent calculation.  WTCS does an average on percent by program
    #Point distribution needs to be reviewed.  100+ with 0 points; 400 with 30 points.

    
df <- read_excel(WTCS,sheet = 'Job Plcmt Satify') %>%
  mutate(fiscal_year = as.integer(fiscal_year) +1,   #survey data trails current FY this applies last FY data to cur FY
         employed_related = ifelse(employed_related > 0,employed_related,NA))


plcmt <- df %>%  #same process as TSA
  group_by(fiscal_year) %>%
  mutate(pct_employed_related = (sum(employed_related_cnt) / sum(employed_cnt))) %>% #rollQuartile cannot find this if it is created in same mutate function
  mutate(measure = 'Placement Rate',
         tmp_benchmark = round(rollQuartile(.,pct_employed_related,fiscal_year),3),
         available_points = 25,
         score = case_when(employed_related > tmp_benchmark ~ available_points,     #top %25 of programs gets all points
                           employed_related <= tmp_benchmark ~ employed_related/tmp_benchmark*available_points,  
                           is.na(employed_related) ~ 0,
                           TRUE ~ 0),
         benchmark = percent(tmp_benchmark,0),
         value = percent(employed_related,0),
         capital_request_score = 0) %>%    #Placement rate is based on 30 points
  ungroup()%>%
  filter(district_group == 'MATC') %>%
  select(fiscal_year,
         program_number,
         measure,
         value,
         benchmark,
         score,
         available_points,
         capital_request_score)


  
#---- get satisfaction ----####

    #uses same dataset as placement rate
satify <- df %>%
  filter(fiscal_year > FISCAL_YEAR - 4,
         district_group == 'MATC') %>% 
  mutate(measure = 'Student Satisfaction',
         benchmark = '100%',
         available_points = 10,
         score = (satisfied_very_satisfied/1.00)*available_points,  #1.00 is benchmark
         value = percent(satisfied_very_satisfied,0),
         capital_request_score = 0) %>%  #satisfaction weight based on 10 points
  select(fiscal_year,
         program_number,
         measure,
         value,
         benchmark,
         score,
         available_points,
         capital_request_score)


rm(df)

#---- Get TSA results ----####

#note; This does not include program with no TSA. The excluded programs will see -- in final report

tsa_apvd <- read_excel(WTCS,sheet = 'TSA Ph2 Approved', col_types = 'text') %>%
  gather('fiscal_year','program_number',1:ncol(.)) %>%
  mutate(fiscal_year = as.integer(fiscal_year),
         tsa_approved = TRUE)

df <- read_excel(WTCS,sheet = 'TSA') %>%
  replaceAllNa() %>%
  mutate(fiscal_year = as.integer(fiscal_year),
         pct_tsa_passed = tsa_passed/(tsa_failed + tsa_passed))

    #WTCS average TSA Pass rate
    #used for reference, not infinal calculation
    #the 75th pctile = 95%
    #when useing this benchmark for scoring 65%+ programs get full 20 pts.
    #adjusting benchmark to 100% reduces above to 55% - 60%
    pct_pass_tsa <- df %>%
      group_by(fiscal_year) %>%
      summarise(t_fail = sum(tsa_failed),
                t_pass = sum(tsa_passed),
                t_tsa = sum(tsa_failed)+sum(tsa_passed)) %>%
      ungroup() %>%
      mutate(pct = t_pass / t_tsa)




tsa <- df %>%
  left_join(tsa_apvd) %>%
  group_by(fiscal_year) %>%
  mutate(fy_pct_tsa_passed = (sum(tsa_passed) / (sum(tsa_failed)+sum(tsa_passed)))) %>% 
  mutate(measure = 'Meet TSA Standard',
         #tmp_benchmark = rollQuartile(.,fy_pct_tsa_passed,fiscal_year),
         tmp_benchmark = 1.00,
         available_points = ifelse(tsa_approved == TRUE,20,0),
         score = (pct_tsa_passed/tmp_benchmark)*available_points,
         benchmark = percent(tmp_benchmark,0),
         value = percent(pct_tsa_passed,0),
         capital_request_score = 0) %>%   #if program has TSA but not in approved list they will get overlooked
  ungroup()%>%
  filter(district_group == 'MATC') %>%
  select(fiscal_year,
         program_number,
         measure,
         value,
         benchmark,
         score,
         available_points,
         capital_request_score)


rm(df,tsa_apvd)

#---- SETUP SOC to PROGRAM CROSSWALK ----####

xwalk <- formatXl(read_excel(XWALK_FILE)) %>%
  mutate(soc = str_extract(soc,"[^ ]+"),
         tmp = substr(program,7,7),
         program_number = ifelse(substr(program,7,7) == ' ',substr(program,1,6),substr(program,1,7)), #some programs have 2 digit sequence number
         program_title = substr(program,8,length(program)))  #titles begining with a space tend to have 2-digit sequence number

  #WTCS SOC Crosswalk codes do not always corresponed with SOC codes used by EMSI.  This builds matches
  xwalk$soc[xwalk$soc == '151142'] <- 151244 #Network and Computer Systems Administrators;  MATC Programs: 101502,311502
  xwalk$soc[xwalk$soc == '151143'] <- 151241 # Computer Network Architects; MATC Programs: 101502
  xwalk$soc[xwalk$soc == '151152'] <- 151231 # Computer Network Support Specialists; MATC Programs: 101502,101504,101513,311502
  xwalk$soc[xwalk$soc == '151121'] <- 151211 # Computer Systems Analysts; MATC Programs: 101502,101504,101513
  xwalk$soc[xwalk$soc == '151122'] <- 151221 # Information Security Analysts; MATC Programs: 101504,101513,311501
  xwalk$soc[xwalk$soc == '151132'] <- 151256 # Software Developers, Applications; MATC Programs: 101527,101528
  xwalk$soc[xwalk$soc == '151134'] <- 151257 # Web Developers; MATC Programs: 101527,102066,312062
  xwalk$soc[xwalk$soc == '151151'] <- 151232 # Computer User Support Specialists; MATC Programs: 101543,301546,311546,311547,611541,611542,611543
  xwalk$soc[xwalk$soc == '292021'] <- 291292 # Dental Hygienists; MATC Programs: 105081
  xwalk$soc[xwalk$soc == '292012'] <- 292018 # Medical and Clinical Laboratory Technicians; MATC Programs: 105131
  xwalk$soc[xwalk$soc == '292071'] <- 292098 # Medical Records and Health Information Technicians; MATC Programs: 105301,315302 
  xwalk$soc[xwalk$soc == '292099'] <- 299098 # Health Technologists and Technicians, All OtherMATC Programs: 105411,315171
  xwalk$soc[xwalk$soc == '311014'] <- 311131 # Nursing Assistants; MATC Programs: 305431
  xwalk$soc[xwalk$soc == '514011'] <- 519161 # Computer-Controlled Machine Tool Operators, Metal and Plastic; MATC Programs: 324441
  xwalk$soc[xwalk$soc == '119199'] <- 119199 # Managers, All Other; MATC Programs: 611021  (EMSI does not use: Managers, All Other)

  xwalk$program_title[xwalk$program_title == 'Supply Chain Assistant'] <- 'Supply Chain Specialist'  #corrects a naming variance causing duplication
  


#---- Calculate Employment Opportunities using EMSI ----####

#EMSI Occupational Outlook tables
emsi_1 <- readWorksheetFromFile(EMSI_1,sheet = 'Occupations') %>% mutate(year = FISCAL_YEAR) #current fiscal_year
emsi_2 <- readWorksheetFromFile(EMSI_2,sheet = 'Occupations') %>% mutate(year = FISCAL_YEAR-1) #
emsi_3 <- readWorksheetFromFile(EMSI_3,sheet = 'Occupations') %>% mutate(year = FISCAL_YEAR-2)

names(emsi_1) <- c('soc',
                   'soc_title',
                   'curr_jobs',
                   'next_jobs',
                   'jobs_chg_cnt',
                   'jobs_chg_pct',
                   #'ann_open',  #included in 2017 data
                   'annual_openings',
                   'hourly_earnings_median',
                   'hourly_earnings_ave',
                   'hourly_earnings_25_pct',
                   'hourly_earnings_75_pct',
                   'year'
)

df <- rbind(emsi_1,setNames(emsi_2, names(emsi_1)))
emsi <- rbind(df,setNames(emsi_3, names(emsi_1)))



    
opport<- emsi %>% 
  mutate_all(list(~str_remove_all(.,'[()$%\\-,]+'))) %>%
  #mutate(annual_openings = ifelse(annual_openings == 'Insf. Data','-99999',annual_openings)) %>%
  left_join(xwalk, by = 'soc') %>%
  filter(!is.na(program_number)) %>%
  group_by(program_number,program_title,year) %>%
  summarize(jobs = sum(as.integer(curr_jobs),na.rm = TRUE),
            openings = sum(as.integer(annual_openings)*10,na.rm = TRUE)) %>% #*10 is carried over from excel formulas
  ungroup() %>%
  mutate(measure = 'Employment Opportunities',
         fiscal_year = as.integer(year),
         value = paste0(round(openings/jobs,1)),
         benchmark = '1.1',  #not sure how Tom determined this but it is consistent across 2015-2018
         available_points = 20,
         score = case_when(value >= 1.1 ~ 20,  #1.1 is the benchmark
                           value >= .6 ~ 13,
                           TRUE ~ 6),
         capital_request_score = ifelse(value > 1.1,1,0)) %>%   #numeric benchmark comparison
  select(fiscal_year,
         program_number,
         measure,
         value,
         benchmark,
         score,
         available_points,
         capital_request_score) %>%
  unique()
  #spread(year,jobs)
#use emsi to review accuracy in crosswalk


rm(emsi_1,emsi_2,emsi_3,
   EMSI_1,EMSI_2,EMSI_3,
   emsi,xwalk,XWALK_FILE,df)


#---- Build Wage dataset ----####

    # prior to 2019 Tom Walsh used Graduate Outcomes Survey Upload File to 
    #   calculate an average hourly wage for all respondents then multiplied 
    #   it by 2080 to get an annaul wage instead of using the WTCS FLW500 report.
  
    #expected file name format is:  FLW500-YYYY.xls
    #TODO: run error check to verify all the files loaded
    #      The most recient year, if not final, begins in row 11 (not 9)

df <- sapply(FLW500_LIST, read_excel, simplify=FALSE , skip = 9) %>% 
  bind_rows(.id = "id") %>%
  select(1,2,4,8:12,15:18) %>%
  mutate(id = as.integer(substr(id,nchar(id)-7,nchar(id)-4)))
#chk <- read_excel('\\Criteria - External Demand\\FLW500-2019.xls', skip = 9) %>%
#   select(1,3,7:11,14:17) 

names(df) <- c('fiscal_year',
               'program_title',
               'program_number',
               'graduates',
               'responses',
               'in_labor_force',
               'employed',
               'employed_related',
               'unemployed_seeking',
               'hourly',
               'annually',
               'avg_hrs_week')

#future improvement - add emsi wage as benchmark
wages <- df %>%
  filter(!program_number %in% c('Total','Division Total')) %>%
  mutate(program_number = gsub("-", "", program_number)) %>%
  mutate(measure = 'Graduate Wages',
         annual_wage = round(annually,0),
         #pct_poverty = annually / FED_POVERTY,  # Useful for scoring
         tmp_benchmark = FED_POVERTY * 3.5,  #350% of poverty level
         available_points = 10,
         score = ifelse(annual_wage > tmp_benchmark,10,(annual_wage/tmp_benchmark)*available_points),
         #benchmark = paste0('$ ',tmp_benchmark),
         benchmark = dollar(tmp_benchmark),
         #value = ifelse(annual_wage == 0,NA,paste0('$ ',annual_wage)),
         value = ifelse(annual_wage == 0,NA,dollar(annual_wage)),
         capital_request_score = ifelse(annual_wage > tmp_benchmark,1,0)) %>%
  select(fiscal_year,
         program_number,
         measure,
         value,
         benchmark,
         score,
         available_points,
         capital_request_score)

rm(df,FLW500_LIST)

#---- Build High Demand Field ----####

high_demand <- read_excel(WTCS,sheet = 'High Demand', col_types = 'text') %>%
  gather('fiscal_year','program_number',1:ncol(.)) %>%
  mutate(value = 'Yes') %>%
  right_join(matc) %>%
  mutate(fiscal_year = as.integer(fiscal_year),
         measure = 'High Demand Field',
         value = ifelse(is.na(value),'No',value),
         benchmark = 'Yes',
         score = ifelse(value == 'Yes',5,0),
         available_points = 0,
         capital_request_score = ifelse(value == 'Yes',1,0)) %>%  #this is treated as extra credit
  unique() %>%
  select(fiscal_year,
         program_number,
         measure,
         value,
         benchmark,
         score,
         available_points,
         capital_request_score)

  #---- Build Transfer Agreement Field ----####

transfer <- read_excel(WTCS,sheet = 'Transfer Agree', col_types = 'text') %>%
  gather('fiscal_year','program_number',1:ncol(.)) %>%
  mutate(value = 'Yes') %>%
  right_join(matc) %>%
  mutate(fiscal_year = as.integer(fiscal_year),
         measure = 'Transfer Agreements',
         value = ifelse(is.na(value),'No',value),
         benchmark = 'Yes',
         score = ifelse(value == 'Yes',10,0),
         available_points = 10,
         capital_request_score = 0) %>%
  select(fiscal_year,
         program_number,
         measure,
         value,
         benchmark,
         score,
         available_points,
         capital_request_score)


#~~~~~~~~~~~~~~~~~~ Criteria 1 ~~~~~~~~~~~~~~~~~####
#~~~~~~~~~~~~~~~ External Demand ~~~~~~~~~~~~~~~####
#~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~####

df <- plcmt %>%
  bind_rows(transfer) %>%
  bind_rows(opport) %>%
  bind_rows(tsa) %>%
  bind_rows(wages) %>%
  bind_rows(satify) %>%
  bind_rows(high_demand) %>%
  filter(fiscal_year > FISCAL_YEAR -3)

m <- as.vector(unique(df$measure))
base <- data.frame(display_order = c(1,2,3,4,5,6,7), 
                   measure = m,
                   base_weight = c(25,10,20,20,10,10,5), 
                   stringsAsFactors = FALSE)


criteria_1 <- matc %>%
  mutate(fiscal_year = as.integer(fiscal_year)) %>%
  crossing(base) %>%
  left_join(df)


criteria_1_total <- criteria_1 %>%
  group_by(fiscal_year,program_number,program_title) %>%
  summarize(base_weight = sum(base_weight,na.rm = TRUE),
            score = sum(score,na.rm = TRUE),
            available_points = sum(available_points,na.rm = TRUE))

  

rm(plcmt,
   transfer,
   opport,
   tsa,
   wages,
   satify,
   high_demand,
   m,base,df,
   PED_POVERTY,
   FLW500_LIST)

return(criteria_1)
}
