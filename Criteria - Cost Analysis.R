# ---------------------------------------------------------------------#
#  Purpose: Build Program Performance Matrix - Cost Analysis Criteria
#  Requestee:
#
#  Created by: John Schliesmann
#  Created Date: 6/18/2019
#
#  Modifications:
#        1. as of 2019 this attempts to duplicate excel workbook created by Tom Walsh
#
#
#  Requirements:
#        1. get program financials from EVA
#
#  Notes:
#        1.   
#               
#               
#  ToDo:
#        1. update colu
#
# ----------------------------------------------------------------------#

processCostAnalysis <- function(){

#~~~~~~~~~~~~~~ USER EDITS ~~~~~~~~~~~~~~~~~~~~~####
#~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~####

# Uncomment to run script independently  
  
#FISCAL_YEAR <- 2021
#WTCS <- paste0('WTCS - Data Source - PQ.xlsx')  #Main dataset from cubes
#matc <- read_excel(WTCS,sheet = 'MATC Programs', skip = 4) 
# FOLDER <- paste0('S:/RESEARCH/06_Program Evaluation/Program Performance Matrix/FY',FISCAL_YEAR,'/')
#FOLDER <- paste0('c:/covid/Program Performance Matrix/FY',FISCAL_YEAR,'/')
 
  
 #Required in function
FOLDER_CA <- paste0(FOLDER,'Criteria - Cost Analysis/')  #this is not used
  
MATC_ACTUALS <- getFileName('Academic_Cost_Analysis',folder = FOLDER_CA)
#~~~~~~~~~~~~~~ Project Setup ~~~~~~~~~~~~~~~~~~####
#~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~####

#none

#~~~~~~~~~~~~~~ Custom Functions ~~~~~~~~~~~~~~~####
#~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~####

# none

#~~~~~~~~~~~~~~ Cost Analysis ~~~~~~~~~~~~~~~~~~####
#~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~####
  

  # Get data sets: Head count and FTE
  
  df <- read_excel(WTCS,sheet = 'Fte and Headcount by Program') 
  matc <- read_excel(WTCS,sheet = 'MATC Programs') 
  
  cat('Processing Cost Analysis\n')
  cat('Years retrieved\n')
  print(table(matc$fiscal_year))
  
  
  df2 <- df %>%
    arrange(program_number,fiscal_year) %>%
    group_by(program_number) %>%
    mutate(fiscal_year = as.integer(fiscal_year),
           fte_chg = (fte - lag(fte))/lag(fte),
           hc_chg = (head_count - lag(head_count))/lag(head_count)) %>%
    ungroup() %>%
    group_by(fiscal_year) %>%
    mutate(fte_lag_sum = sum(lag(fte),na.rm = TRUE),
           fte_matc = (sum(fte,na.rm = TRUE) - fte_lag_sum)/fte_lag_sum,
           hc_lag_sum = sum(lag(head_count),na.rm = TRUE),
           hc_matc = (sum(head_count,na.rm = TRUE) - hc_lag_sum)/hc_lag_sum) %>%
    ungroup() %>%
    mutate(fte_diff = case_when(fte_chg > fte_matc + .01 ~ 'Faster',
                                fte_chg < fte_matc - .01 ~ 'Slower',
                                is.na(fte_chg) ~ NA_character_,
                                TRUE ~ 'Same'),
           hc_diff = case_when(hc_chg > hc_matc + .03 ~ 'Faster',
                               hc_chg < hc_matc - .03 ~ 'Slower',
                               is.na(hc_chg) ~ NA_character_,
                                TRUE ~ 'Same'))
  
  
  #---- Build Head Count / FTE  dataset ----####
  cnt <- df2 %>%
    select(fiscal_year,
           program_number,
           `Program Total FTE` = fte,
           `Program Total Headcount (HC)` = head_count) %>%
    gather(3:4,key = 'measure',value = 'value') %>%
    getBenchmark(col = value ,p = .75) %>%
    mutate(available_points = 10,
           capital_request_score = ifelse(measure == 'Program Total FTE' &
                                            value > benchmark,1,0),
           score = case_when(value > benchmark ~ available_points,     #top %25 of programs gets all points
                             value <= benchmark ~ value/benchmark*available_points,  
                             is.na(value) ~ 0,
                             TRUE ~ 0),
           value = ifelse(is.na(value),NA_character_,as.character(round(value,0))),
           benchmark = ifelse(is.na(benchmark),NA_character_,as.character(round(benchmark,0)))) %>% 
    select(fiscal_year,
           program_number,
           measure,
           value,
           benchmark,
           score,
           available_points,
           capital_request_score)
  
  #---- Build % Change in Head Count / FTE  dataset ----####
  pct_cnt <- df2 %>%
    select(fiscal_year,
           program_number,
           `% Change in FTE from Prevous Yr` = fte_chg,
           `% Change in HC from Prevous Yr` = hc_chg) %>%
    gather(3:4,key = 'measure',value = 'value') %>%
    getBenchmark(col = value ,p = .75) %>%
    mutate(available_points = 10,
           capital_request_score = ifelse(measure == '% Change in FTE from Prevous Yr' &
                                            value > benchmark,1,0),
           score = case_when(value > .3 ~ available_points,
                             value > .1 ~ available_points *.75,
                             value > -.1 ~ available_points *.5,
                             value > -.3 ~ available_points *.25,
                             is.na(value) ~ 0,
                             TRUE ~ 0 ),
           value = ifelse(is.na(value),NA_character_,percent(value,0)),
           benchmark = ifelse(is.na(benchmark),NA_character_,percent(benchmark,0))) %>%
    select(fiscal_year,
           program_number,
           measure,
           value,
           benchmark,
           score,
           available_points,
           capital_request_score)
  
  
  #---- Build School comparison in % Change dataset ----####
  sch_comp <- df2 %>%
    select(fiscal_year,
           program_number,
           `FTE % Chg compared to MATC` = fte_diff,
           `% Chg in HC compared to MATC` = hc_diff) %>%
    gather(3:4,key = 'measure',value = 'value') %>%
    mutate(benchmark = 'Faster',
           available_points = 10,
           capital_request_score = 0,
           score = case_when(value == 'Faster'~ 10,
                             value == 'Same' ~ 6,
                             value == 'Slower' ~ 3,
                             is.na(value) ~ 0,
                             TRUE ~ 0)) %>%
    select(fiscal_year,
           program_number,
           measure,  #description of value
           value,
           benchmark,
           score,
           available_points,
           capital_request_score)
  
  rm(df,df2,cube)
  
  #---- Build % Change in Head Count / FTE  dataset ----####
  #---- Get Program Actuals ----####
  
  dept_prog_xwalk <- matc %>%
    mutate(department = substr(program_number,3,5)) %>%
    select(program_number,department) %>%
    unique()
    
  fin_hlth <- read_excel(MATC_ACTUALS,skip = 3,col_types = c(rep('text',4),'numeric')) %>%  #column A maybe blank in Excel
    select(fiscal_year = 1,
           department = 2, 
           cost_center = 3,
           cost_center_descr = 4,
           actuals = 5)  %>%  
    filter(!cost_center_descr %in% c('Office Techology - AHS',
                                     'Comprehensive Homemaking - HSC',
                                     'Culinary Arts - AHS',
                                     'Culinary Arts - HSC',
                                     'Automobile Servicing - AHS',
                                     'Automobile Servicing - HSC',
                                     'Electricity-EPD',
                                     'Mechanical Drafting - AHS',
                                     'Sm Engine & Chassis Mech - HSC'
                                     )) %>%
    group_by(fiscal_year,department) %>%
    summarize(actuals_int = as.integer(sum(actuals, na.rm = TRUE))) %>%
    ungroup() %>%
    arrange(fiscal_year) %>%
    group_by(department) %>%
    mutate(chg = actuals_int - dplyr::lag(actuals_int),
           pct_chg = ifelse(actuals_int == 0,-1,round(chg/actuals_int,2))) %>%
    ungroup() %>%
    full_join(dept_prog_xwalk) %>%
    right_join(matc) %>%
    mutate(fiscal_year = as.integer(fiscal_year),
           measure = 'Revenue Growth',
           benchmark = 'Increasing',
           available_points = 10,
           capital_request_score = 0,
           value = case_when(pct_chg > .04 ~ 'Increasing',
                             pct_chg < -.04 ~ 'Decreasing',
                             is.na(pct_chg) ~ NA_character_,
                             TRUE ~ 'Remains Same'),
           score = case_when(value == 'Increasing'~ 20,
                             value == 'Remains Same' ~ 15,
                             value == 'Decreasing' ~ 10,
                             TRUE ~ 0)) %>%
    select(fiscal_year,
           program_number,
           measure,  #description of value
           value,
           benchmark,
           score,
           available_points,
           capital_request_score)


  
  
#---- combine dataset  ----####
  
df <- cnt %>%
  bind_rows(pct_cnt) %>%
  bind_rows(sch_comp)  %>%
  bind_rows(fin_hlth) %>%
  filter(fiscal_year > FISCAL_YEAR -3) %>%
       mutate(dept = substr(program_number,3,5))
  
  
m <- sort(as.vector(unique(df$measure)))  
#SAVE FOR FUTURE INCLUSION OF FINANCIAL HEALTH
base <- data.frame(display_order = c(21,24,25,22,20,23,26), #this will reorder the sort of m
                   measure = m,
                   base_weight = c(10,10,10,10,10,10,20),
                   stringsAsFactors = FALSE)
 
# SAVE FOR USE WITHOUT FINANCIALS
# base <- data.frame(display_order = c(23,26,27,24,22,25),  #order FTE then HC
#                    measure = m, 
#                    base_weight = c(10,10,10,10,10,10), 
#                    stringsAsFactors = FALSE) 


criteria_4 <- matc %>%
  mutate(fiscal_year = as.integer(fiscal_year)) %>%
  crossing(base) %>%
  left_join(df)


rm(cnt,
   pct_cnt,
   sch_comp,
   dept_prog_xwalk,
   fin_hlth,
   df,m,
   base)



table(criteria_4$measure,criteria_4$base_weight) 
# Expecting:
#                                 10  30
#% Change in FTE from Prevous Yr ###   0
#% Change in HC from Prevous Yr  ###   0
#% Chg in HC compared to School  ###   0
#FTE % Chg compared to School    ###   0
#Program Total FTE               ###   0
#Program Total Headcount         ###   0
#hold                            ###   0
#hold                                ###

table(criteria_4$measure,criteria_4$display_order) 
# Expecting: 
#                                 22  23  24  25  26  27
#% Change in FTE from Prevous Yr   0 ###   0   0   0   0
#% Change in HC from Prevous Yr    0   0   0   0 ###   0
#% Chg in HC compared to School    0   0   0   0   0 ###
#FTE % Chg compared to School      0   0 ###   0   0   0
#Program Total FTE               ###   0   0   0   0   0
#Program Total Headcount           0   0   0 ###   0   0



  #use for review
  chk <- criteria_4 %>%
    filter(fiscal_year == '2018',
           measure %in% c('FTE % Chg compared to School',
                          '% Chg in HC compared to School'))
  table(chk$measure,chk$value)

  
  return(criteria_4)
 
}
 
# 
# #NEED TO REVIEW SCORE DISTRIBUTION
# library(shiny)
# library(ggplot2)
# library(plotly)
# df <- ca%>%
# #df <- criteria_2 %>%
#   filter(measure == '% Change in FTE from Prevous Yr',
#          #fiscal_year == FISCAL_YEAR,
#          program_number != '504209',
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
#       geom_bar(fill = "#0046AD")
# 
#   })
# }
# 
# shinyApp(ui = ui, server = server)
# table(df$score,df$value)