# ---------------------------------------------------------------------#
#  Purpose: Build Program Performance Matrix - Main Script & Builds PDF
#  Requestee: Yan Wang
#
#  Created by: John Schliesmann
#  Created Date: 6/18/2019
#  
#  Last Modified: 1/29/2020 - tweaked scoring for FTE & Hc in Cost Analysis
#                 6/1/2021  - Updated {Data Source - PQ.xlsx} to use new server (cubes.wtcsystem.edu) and PowerQuery
#                 7/20/2021 - incorporated actuals from Accounting Office into Cost Analysis Criteria
#
#  Requirements:
#        1. External Demand: download 3 years of FLW500, APR500, and SOC Crosswalk from WTCS Portal and 5 years of EMSI Occupation Table to /FY####/Criteria - External Demand
#          a. The most recent year of graduate survey data is dependent on WTCS upload in mid Janurary
#        2. Student Success:  all data is from the WTCS Cubes.
#        3. Program Uniqueness: download 3 years of CLI330 and lastest PGM307 from WTCS Portal to /FY####/Criteria - Program Uniqueness
#
#  Notes:
#        1. as of 2019 this attempts to duplicate excel workbook created by Tom Walsh
#        2. Capital Request Industry Impact score begins near line 263. (single dataframe)
#
#  ToDo:
#        1. build high level criteria comparison of all MATC programs
#        2. Evaluate the need to have datasets external to WTCS in each FY
#        3. Add file.exists error checks
#        4. improve Capital Project Industry Impact Score (cr_score) aka Industry Impact
#        5. does it continue to compare / benchmark all 10,30,31,32,50 together?
#        6. consolidate EMSI, PGM307, CLI330, FLW & other non-cube data into R_Project to remove 
#              duplicative downloads.
#        7. automate most recient IPEDS database
#
# ----------------------------------------------------------------------#


#~~~~~~~~~~~~~~ USER EDITS ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~####
#~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~####

FISCAL_YEAR <- 2021  #change to previous FY
IPEDS_YEAR <- FISCAL_YEAR + 200000 -101  #defaults to previous FY could be same as fiscal year

#Federal Poverty level
    #https://aspe.hhs.gov/poverty-guidelines last updated 5/2021
    #Used in External Demand Criteria
    #The value below reflects a family unit of 1 in the 48 contiguous states
FED_POVERTY <- 12880   


#Deprecitated Programs
del <- c('101095', #Tourism & Travel Management 
         '000000'
)

#~~~~~~~~~~~~~~ Project Setup ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~####
#~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~####


#FOLDER <- paste0('S:\\RESEARCH\\06_Program Evaluation\\Program Performance Matrix\\FY',FISCAL_YEAR,'\\')
#FOLDER <- paste0('C:\\covid\\Program Performance Matrix\\FY',FISCAL_YEAR,'\\')
FOLDER <- paste0(dirname(getwd()),'\\FY',FISCAL_YEAR,'\\')
WTCS <- paste0('WTCS - Data Source - PQ.xlsx')  #Main dataset from cubes
#MATC_ACTUALS <- found in Cost Analysis.R  -> Academic_Cost_Analysis.xlsx   #Fincanial Data from Finance Dept (Brenda Schmidt)
                #expecting 5 columns:  Fiscal year, shortened cost center, cost center, cost center name, total actuals


IMAGES <- paste0(getwd(),'/images')

IPEDS_FOLDER <- 'S:\\RESEARCH\\55_John Schliesmann\\R_IPEDS\\data\\'
#IPEDS_FOLDER <- 'C:\\covid\\External Datasets\\IPEDS\\'


#this is iteration #3 in the archive folder
source(paste0(getwd(),'/scripts/Criteria - External Demand.R'))
source(paste0(getwd(),'/scripts/Criteria - Student Success.R'))
source(paste0(getwd(),'/scripts/Criteria - Program Uniqueness.R'))
source(paste0(getwd(),'/scripts/Criteria - Cost Analysis.R'))


#Create directory structure if it doesn't exist

dir.create(file.path(FOLDER), showWarnings = FALSE)
#dir.create(file.path(FOLDER, 'Criteria - Cost Analysis'), showWarnings = FALSE)      # not necessary
dir.create(file.path(FOLDER, 'Criteria - External Demand'), showWarnings = FALSE)     #for additional reports
dir.create(file.path(FOLDER, 'Criteria - Program Uniqueness'), showWarnings = FALSE)  #for additional reports
#dir.create(file.path(FOLDER, 'Criteria - Student Success'), showWarnings = FALSE)    # not necessary
dir.create(file.path(FOLDER, 'output'), showWarnings = FALSE)



#~~~~~~~~~~~~~~ Include Libraries ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~####
#~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~####

library(tidyverse)
library(RODBC)
library(stringr)
library(zoo)  
library(scales)
library(readxl)
library(XLConnect)
library(rmarkdown)

options(scipen=999)

#~~~~~~~~~~~~~~ Shared Functions ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~####
#~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~####

getFileName <- function(file_name = 'all',folder = getwd()) {
  ####  To Do: customize Descr column name
  ####  Notes:
  ####      1. fuzzy match file name search
  ####      2. Could result in error if multiple file types exist
  ####
  ####  Arguments:
  ####      file_name:  a string pattern in the file name  #note: a . and a _ a the same
  ####      folder:  starting point for searching.  Includes all subdirectories
  
  ####  Troubleshooting
  #file_name <- 'Table q'
  #folder <- getwd()
  
  
  if(!("tidyverse" %in% (.packages()))) { library(tidyverse) }
  if(!("readxl" %in% (.packages()))) { library(readxl) }
  if(!("logging" %in% (.packages()))) { library(logging) }
  
  
  if(file_name == 'all') {
    l <-  file.info(
      list.files(folder,
                 pattern = file_name,
                 ignore.case = TRUE,
                 full.names=TRUE,
                 recursive=TRUE,
                 include.dirs=TRUE))
    #print(rownames(l[with(l, order(desc(as.POSIXct(ctime)))),]))
  }else{
    l <- file.info(
      list.files(folder,
                 pattern = file_name,
                 ignore.case = TRUE,
                 full.names=TRUE,
                 recursive=TRUE,
                 include.dirs=TRUE))
    l$file <- basename(rownames(l))
    l <- l[!grepl("^~\\$", l$file),]   #removes open excel files from list
    l <- rownames(l[with(l, order(desc(as.POSIXct(mtime)))),])[1]  #retrieves the most reciently modified file
    if(is.na(l)){
      logwarn(paste0("FILE NOT FOUND: '",file_name," looking in",folder,'\n'),logger = 'Get')
      l <- FALSE
    }else{
      loginfo(paste0("Found File: ",l,'\n'),logger = 'Get')
    }
  }
  return(l)
}
getBenchmark <- function(.data,col,p = .75){

  #method: This first gets three year rolling average by program then 
  #        gets 75% percentile by fiscal year
  #Arguments:  .data is the data frame
  #            col is a numeric column in the data frame: the measure value
  #            p is the probs for the quantile function
  #Note:   this uses conditional group_by.  purrr nest()/map() will not work with rlang functionality
  #        the col argument may not be necessary if it is always 'value'.  Or add grp_by arg, see:
  #        https://stackoverflow.com/questions/55246913/function-to-pass-parameter-to-perform-group-by-in-r

  .data %>%
    arrange(fiscal_year) %>%
    group_by(program_number,
             {if("aid_code" %in% names(.)) aid_code else program_number},
             {if("measure" %in% names(.)) measure else program_number}
             ) %>%
    mutate(ave = as.numeric(rollapply(!!ensym(col),3, FUN = mean, 
                                      partial = TRUE, align = "right"))) %>%
    ungroup() %>%
    group_by(fiscal_year,
             {if("aid_code" %in% names(.)) aid_code else fiscal_year},
             {if("measure" %in% names(.)) measure else fiscal_year}
             ) %>%
    mutate(benchmark = quantile(ave,probs = p,na.rm = TRUE)) %>%
    ungroup() 
}

percent <- function(x, digits = 2, format = "f", ...) {
  #method: converts decimal to text with trailing %
  #Arguments: x is a non-vectored number
  #           digits is the number precision
  #           format is the C style format codes 
  #This function is also in MATC Custom Functions
  ifelse(is.na(x),
         NA,
         paste0(formatC(100 * x, format = format, digits = digits, ...), "%")
  )
}


convertYear <- function(x) {
  y <- paste0(x-1,'-',substr(x,3,4))
  
}

quiet <- function(x) { 
  #sink(tempfile()) 
  #on.exit(sink()) 
  #invisible(force(x))
  suppressMessages(x) 
} 


checkMeasureFreq <- function(x,m) {
  df <- x %>%
    filter(measure == m) 
    #group_by(fiscal_year,value) %>%
    #count()
  cat('Measure: ',m,'\n')
  print(table(df$fiscal_year,df$value))
  cat('\n\n')
}


#~~~~~~~~~~~~~~ Create Criteria Data Sets ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~####
#~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~####

matc <- read_excel(WTCS,sheet = 'MATC Programs') 
 # mutate(fiscal_year = as.integer(fiscal_year))  #not sure yet if this is important...

#source('Criteria - External Demand.R)

ed <- quiet(processExternalDemand() %>% mutate(criteria = 'External Demands'))
ss <- quiet(processStudentSuccess() %>% mutate(criteria = 'Student Success'))
pu <- quiet(processProgramUniqueness()  %>% #check script, it has custom location
   mutate(criteria = 'Program Uniqueness'))
ca <- quiet(processCostAnalysis() %>% mutate(criteria = 'Cost Analysis'))

#---- Merge PPM Catagories

m <- ed %>%
  bind_rows(ss) %>%
  bind_rows(pu) %>%
  bind_rows(ca) %>%
  select(-cip_code) %>%
  filter(fiscal_year > FISCAL_YEAR-3,
         !(program_number %in% del)) %>%
  mutate(score = ifelse(is.na(score),0,score)) %>% #this ideally should not be needed (all program rows should be created in the 4 processes
  select(program_title,
         program_number,
         criteria,
         measure,
         display_order,
         fiscal_year,
         value,
         benchmark,
         base_weight,
         available_points,
         score,
         capital_request_score)  #indicates if criteria is above benchmark (not truely a score)

#---- cleanup

#rm(matc,ed,ss,pu,ca)

#---- Build datasets for report creation and export

summary <- m %>%
  filter(fiscal_year > FISCAL_YEAR -3) %>%
  group_by(fiscal_year,program_number,program_title,criteria) %>%
  summarize(score = ifelse(sum(score) == 0,NA,sum(score,na.rm = TRUE)),  #Need NA for cut catagory
            base_weight = sum(base_weight,na.rm = TRUE),
            display_order = min(display_order)) %>%
  ungroup() %>%
  group_by(fiscal_year,criteria) %>%
  mutate(inst_rank = as.numeric(
                       as.character(
                         ifelse(is.na(score),0,
                                cut_number(score,   #higher number is higher rank; NA = 0
                                            n = 3,
                                           labels = 1:3))
                        )),
         inst_rank_descr = case_when(inst_rank == 3 ~ 'Gold',
                                     inst_rank == 2 ~ 'Silver',
                                     inst_rank == 1 ~ 'Bronze',
                                     TRUE ~ 'Inc'),
         display_year = convertYear(fiscal_year)) %>%
  ungroup() %>%
  group_by(program_number,criteria) %>%
  arrange(fiscal_year) %>%
  mutate(chg_rank = case_when(inst_rank - lag(inst_rank) > 0 ~ 'Up',
                               inst_rank - lag(inst_rank) == 0 ~ 'Flat',
                               inst_rank - lag(inst_rank)  < 0 ~ 'Down',
                               TRUE ~ 'Unk'
                               ),
         score = ifelse(is.na(score),0,round(score,1))
  ) %>%
  ungroup() %>%
  mutate_if(is.character,
            str_replace_all, pattern = "/", replacement = ", ") %>%
  arrange(program_number,display_order,fiscal_year)
  

detail <- m %>% 
  select(program_number,display_order,criteria,Measure = measure,benchmark,value,fiscal_year) %>%
  gather('type','data',c(benchmark,value)) %>%
  filter(type == 'value' | (type == 'benchmark' & fiscal_year == FISCAL_YEAR),
         fiscal_year > FISCAL_YEAR -3) %>%
  mutate(display_year = ifelse(type == 'benchmark', 
                               paste0('Benchmark (',convertYear(fiscal_year),')'),
                               convertYear(fiscal_year))) %>%
  select(-type,-fiscal_year) %>%
  spread(display_year,data) %>%
  arrange(program_number,display_order) %>%
  select(1:(ncol(.)-4),ncol(.),everything(),-display_order) %>%  #re-position benchmark
  group_by(program_number,criteria) %>%
  nest() %>%
  ungroup() %>%
  group_by(program_number) %>%
  nest()
     #Warning message:  attributes are not identical across measure variables; they will be dropped --> This is OK



 ###--------- Capital Request ------------------------------------------------###
      #used in the Construction-renovation Capital Request google sheet as
      #the Industry Impact score
      cap_rqst <- m %>%
        filter(fiscal_year == FISCAL_YEAR-1) %>%
        arrange(program_number) %>%
        group_by(fiscal_year,
                 program_number,
                 program_title) %>%
        summarize(met_benchmarks = sum(capital_request_score,na.rm = TRUE),
                  cr_score = case_when(met_benchmarks > 4 ~ 9,
                                       met_benchmarks > 2 ~ 6,
                                       met_benchmarks > 0 ~ 3,
                                       TRUE ~ 0)) %>%
        ungroup() %>%
        mutate(program_number = paste0(substr(program_number,1,2),'-',
                                     substr(program_number,3,5),'-',
                                     substr(program_number,6,6)),
               department = substr(program_number,4,6)) %>%
        group_by(department) %>%            #find max cr_score by department used by William (Bill) Smith
        mutate(cr_dept_max_score = max(cr_score)) %>%
        ungroup() %>%
        select(-met_benchmarks)
      
      
      cat('Distribution of Capital Requests Scores\n')
      table(cap_rqst$cr_score)
      write.csv(cap_rqst,paste0(FOLDER,'Capital Requests - Industy Impact Score.csv'))
###--------- Capital Request ------------------------------------------------###


write.csv(m,paste0(FOLDER,'Program Performance Matrix Dataset - ',FISCAL_YEAR,'.csv'),row.names = FALSE)
save.image(file = paste0(FOLDER,'Program Performance Matrix Complete Environment - ',FISCAL_YEAR,'.RData'))


#~~~~~~~~~~~~~~ Error Checking ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~####
#~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~####

#---- commented code used for debugging
#   tmp <- detail %>% 
#     filter(program_number == '100014') %>%
#     select(data) %>%
#     unnest()

checkMeasureFreq(m,'Placement Rate')
checkMeasureFreq(m,'Transfer Agreements')
checkMeasureFreq(m,'Part of Pathway Credential')
checkMeasureFreq(m,'Meet TSA Standard')
checkMeasureFreq(m,'Graduate Wages')


table(summary$criteria,summary$base_weight) #expect all equal values

nrow(summary %>% filter(is.na(summary$score)))  #should be zero


#Look for missing benchmarks
chk <- m %>%
  filter(fiscal_year == FISCAL_YEAR,
         is.na(benchmark),
         measure != 'Meet TSA Standard')
table(chk$measure)
table(chk$program_number)



#review score distribution
chk <- ss %>%
  filter(measure == '% Point Change in Retention Rate') 
table(chk$fiscal_year,chk$score)


#~~~~~~~~~~~~~~ Publish Criteria ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~####
#~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~####

#todo:  add a line to give screen message for the first pdf to print
#       add a eta for completion      

#get list of programs for the reporting fiscal year
list_programs <- summary %>%
#df <- summary %>%
                    filter(fiscal_year == FISCAL_YEAR) %>%
                           #substr(program_number,1,2) == '10') %>%
                    select(program_number) %>%
                    unique() %>%
                    pull()   #this function converts a df column to vector
                    
#list_programs <- df[c(12,14,24)]
#list_programs <- c('102031')

for (i in 1:length(list_programs)){  

#Get Data to fill in PPM Template
  
  #get title and program number  
  description <- summary %>%
    filter(program_number == list_programs[i]) %>%
    select(program_title,program_number) %>%
    mutate(program_number = paste0(substr(program_number,1,2),'-',
                                   substr(program_number,3,5),'-',
                                   substr(program_number,6,6))) %>%
    unique()
  
  
  #get summary table and add images
  s <- summary %>%
    filter(program_number == list_programs[i],
           fiscal_year == FISCAL_YEAR) %>%
    mutate(Level = sprintf('\\raisebox{-.3\\totalheight}{\\includegraphics[width=0.02\\textwidth, height=4mm]{%s/%s_large.png}}', IMAGES, inst_rank),
           lvl_chg = sprintf('\\raisebox{-.3\\totalheight}{\\includegraphics[width=0.02\\textwidth, height=4mm]{%s/%s.png}}', IMAGES, chg_rank)
           ) %>%
    select(Criteria = criteria,
           `Total Score` = score,
           Level,
           ` ` = inst_rank_descr,
           `Level Change From Last Year` = lvl_chg) 
  
  #get detail tables
   d <- detail %>%
    filter(program_number == list_programs[i]) %>%
    unnest(cols = c(data)) 
  
   #get capital request score
   cr_score <- cap_rqst %>%
     filter(program_number == list_programs[i]) 
   
     
     
  cat(paste0('\nPRINTING: ', description[2],' - ',description[1]  ,'  (',i,' of ',length(list_programs),')'))    
  
  render(
    input = "./scripts/PPM_Template.Rmd",  # path to the template
    #output_file = paste0(description[2],' - ',description[1], '.pdf'),  
    output_file = paste0(description[1],' (',description[2], ').pdf'),  
    output_dir = paste0(FOLDER,"/output"), 
    clean = TRUE, #removes work file (.tex, etc)
    quiet = TRUE #removes console messages
  )

  rm(description, s, d, cr_score)
}
  
 
