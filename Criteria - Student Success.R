# ---------------------------------------------------------------------#
#  Purpose: Build Program Performance Matrix - Student Success Criteria
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
#        1. Update the pivot table filters in WTCS - Data Source.xlsx
#
#  Notes:
#        1. capital_request_score is only added to criteria_2 df.
#
#  ToDo:
#        1. add base weight, score, available_points
#
# ----------------------------------------------------------------------#


#~~~~~~~~~~~~~~ USER EDITS ~~~~~~~~~~~~~~~~~~~~~####
#~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~####
# FISCAL_YEAR <- 2019


#~~~~~~~~~~~~~~ Project Setup ~~~~~~~~~~~~~~~~~~####
#~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~####

# uncomment for running script independently

# library(tidyverse)
# library(readxl)
# library(zoo)
# 
# 
# options(scipen=999)
# 
# source('S:\\RESEARCH\\55_John Schliesmann\\MATC Custom Functions.R')
# source('S:\\RESEARCH\\55_John Schliesmann\\Client Reporting Functions.R')

# WTCS <- paste0('WTCS - Data Source.xlsx')  #Main dataset from cubes


#~~~~~~~~~~~~~~ STUDENT SUCCESS ~~~~~~~~~~~~~~~~####
#~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~####

processStudentSuccess <- function(){
  

formatSS <- function(topic,p = .75,...){
  #Argument: Topic correspondes to the excel workbook tab name.

  
df <- read_excel(WTCS,sheet = topic) 

#if(ncol(df) == 4){   #finds Course 
if(topic %in% c('Course Completion','Retention')){  
  df <- df %>%
    rename(n = 3, p = 4) %>%  #value in column C (3) is count and column D is calculated percentile
    mutate(fiscal_year = as.integer(fiscal_year),
           p = round(p,2)) %>%
    select(fiscal_year,
           program_number,
           value = p)  
#} else if(ncol(df) == 6) {    #find graduation dataset
} else if(topic == 'Graduation') {
  df <- df %>%
    rename(grad_in_2 = 4,grad_in_3 = 6) %>%
    mutate(aid_code = substr(program_number,1,2),
           fiscal_year = as.integer(fiscal_year),
           value = case_when(aid_code %in% c('10','32') ~ round(grad_in_3,2),
                             aid_code %in% c('30','31') ~ round(grad_in_2,2),
                             TRUE ~ NA_real_)) %>%
    select(fiscal_year,
           program_number,
           aid_code,
           value)
} else { stop() }

rt <- df %>%
  getBenchmark(col = value,p) %>%         #benchmark is average percent of program percents
  mutate(measure = paste(topic,'Rate'),   #just renaming the excel tab
         available_points = case_when(topic == 'Course Completion' ~ 35,
                                      TRUE ~ 15),
         score = case_when(value > benchmark ~ available_points,     #top %25 of programs gets all points
                           value <= benchmark ~ value/benchmark*available_points,  
                           is.na(value) ~ 0,
                           TRUE ~ 0)) %>%  
  select(fiscal_year,
         program_number,
         measure,
         value,
         benchmark,
         score,
         available_points)

chg_rt <- df %>%
  group_by(program_number) %>%
  mutate(value = value - lag(value,order_by = fiscal_year),
         measure = paste('%age Point Change in',topic,'Rate'),
         available_points = case_when(topic == 'Course Completion' ~ 15,
                                      TRUE ~ 10)) %>%
  ungroup() %>%
  getBenchmark(col = value,p) %>%
  mutate(score = case_when(value > .15 ~ available_points,
                           value > .06 ~ available_points *.75,
                           value > -.07 ~ available_points *.5,
                           value > -.16 ~ available_points *.25,
                           is.na(value) ~ 0,
                           TRUE ~ 0
                           )) %>%
  select(fiscal_year,
         program_number,
         measure,
         value,
         benchmark,
         score,
         available_points)
  
x <- rbind(rt,chg_rt)




rm(df,rt,chg_rt)
return(x)

}





#~~~~~~~~~~~~~~~~~~ Criteria 2 ~~~~~~~~~~~~~~~~~####
#~~~~~~~~~~~~~~~ Student Success ~~~~~~~~~~~~~~~####
#~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~####


ccr <- formatSS('Course Completion',p=.8)
rtn <- formatSS('Retention')
grd <- formatSS('Graduation')  #the benchmark is for aid code 10, 32 uses 3yr, and aid code 30,31 uses 2yr


df <- ccr %>%
  bind_rows(rtn) %>%
  bind_rows(grd) %>%
  filter(fiscal_year > FISCAL_YEAR -3) %>%
  mutate(value = percent(value,0),
         benchmark = percent(benchmark,0),
         capital_request_score = 0)

m <- as.vector(unique(df$measure))
base <- data.frame(display_order = c(8,9,10,11,12,13), 
                   measure = m,
                   base_weight = c(20,10,20,10,20,10), 
                   stringsAsFactors = FALSE)


criteria_2 <- read_excel(WTCS,sheet = 'MATC Programs', col_types = 'text') %>%
  mutate(fiscal_year = as.integer(fiscal_year)) %>%
  crossing(base) %>%
  left_join(df)


#---- Error checking -------------------####

# #NEED TO REVIEW SCORE DISTRIBUTION
# library(shiny)
# library(ggplot2)
# library(plotly)
# df <- ss%>%
# #df <- criteria_2 %>%
#   filter(measure == '%age Point Change in Retention Rate',
#          fiscal_year == FISCAL_YEAR,
#          !is.na(value)) %>%
#   mutate(value = as.numeric(str_replace(value,'%',''))) %>%
#   select(fiscal_year,measure,value,score)
# 
# ui <- fluidPage(
#   plotlyOutput("distPlot")
# )
# 
# server <- function(input, output) {
#   output$distPlot <- renderPlotly({
#     #df %>% ggplot(aes(score)) +
#     df %>% ggplot(aes(value)) +
#       geom_bar(fill = "#0073c2ff")
# 
#   })
# }
# 
# shinyApp(ui = ui, server = server)
# table(df$score,df$value)


#Point distribution needs to be reviewed

#scores for % Point Change in Graduation Rate
#      0   2.5  5    7.5  10
#2017  50  35   73   23   5
#2018  26  19   104  12   7
#2019  29  22   98   18   7

#scores for % Point Change in Retention Rate
#      0   2.5  5    7.5  10
#2017  57  30   89   29   8
#2018  39  33   91   18   11
#2019  39  27   87   39   5


#check for correct assignments of base weight
#each measure should only have one weight value
table(criteria_2$measure,criteria_2$base_weight)


criteria_2_total <- criteria_2 %>%
  group_by(fiscal_year,program_number,program_title) %>%
  summarize(base_weight = sum(base_weight,na.rm = TRUE),
            score = sum(score,na.rm = TRUE),
            available_points = sum(available_points,na.rm = TRUE))



rm(m,
   df,
   base,
   ccr,
   grd,
   rtn)

return(criteria_2)
}
