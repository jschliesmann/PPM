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

processLaS <- function(){
  
  #~~~~~~~~~~~~~~ USER EDITS ~~~~~~~~~~~~~~~~~~~~~####
  #~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~####
  
  # Uncomment to run script independently  
  
  FISCAL_YEAR <- 2021
  
  
  

  # FOLDER <- paste0('S:\\RESEARCH\\06_Program Evaluation\\Program Performance Matrix\\FY',FISCAL_YEAR,'\\')
  setwd(paste0('c:\\covid\\Program Performance Matrix\\FY',FISCAL_YEAR,'\\'))
  
  FOLDER_LAS <- paste0(FOLDER,'Letters and Sciences\\')  #this is not used
 
  
  #~~~~~~~~~~~~~~ Project Setup ~~~~~~~~~~~~~~~~~~####
  #~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~####
  
  #none
  
  #~~~~~~~~~~~~~~ Custom Functions ~~~~~~~~~~~~~~~####
  #~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~####
  
  # none
  
  #~~~~~~~~~~~~~~ Cost Analysis ~~~~~~~~~~~~~~~~~~####
  #~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~####
  
  cat('Processing LAS Addendum\n')
  cat('Years retrieved\n')
  print(table(matc$fiscal_year))
  
  # Get data sets: Head count and FTE
  
  crse <- read_excel(getFileName('XCCC'),sheet = 'R Script') %>%
    mutate(student_id = str_pad(student_id,9,'left','0'),
           year = substr(term,3,6),
           course_credits = as.integer(course_credits)) %>%
    replace_na(list(course_grade_point = 0,
               gender = 'U',
               zip = '00000')) %>%
    filter(year <= FISCAL_YEAR) %>%
    rename(race_ethnicity = reace_ethnicity)
  
  
  #chk for duplicate enrollments
  df <- crse %>%
    group_by(term,
             course_name,
             student_id) %>%
    count() %>%
    filter(n > 1) %>%
    group_by(term) %>%
    count()
             
  
  
  dept_fte <- crse %>%
    ungroup() %>%
    group_by(year,
             course_department,
             course_department_descr
             ) %>%
    summarize(#n = n(),
              fte = sum(course_credits)) %>%
    unique() %>%
    spread(year,fte)
  
  
  dept_cnt <- crse %>%
    ungroup() %>%
    group_by(year,
             course_department,
             course_department_descr
    ) %>%
    summarize(n = n()) %>%
    spread(year,n) 
  
  
  dept_eth <- crse %>%
    group_by(year,
             course_department,
             course_department_descr,
             race_ethnicity
    ) %>%
    count() %>%
    ungroup() %>%
    group_by(year,
             course_department,
             course_department_descr
    ) %>%
    mutate(pct = round(n/sum(n),3)*100) %>%
    select(-n) %>%
    spread(year,pct,fill = 0) 
    
  
  
  
 ####--------- DELETE BELOW ????
  
  
  df2 <- matc %>%
    left_join(df) %>%
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
  
  rm(df,df2)
  
  #---- Build % Change in Head Count / FTE  dataset ----####
  #---- Get Program Actuals ----####
  

  
  
  #---- combine dataset  ----####
  df <- cnt %>%
    bind_rows(pct_cnt) %>%
    bind_rows(sch_comp)  %>%
    filter(fiscal_year > FISCAL_YEAR -3)
  
  
  m <- sort(as.vector(unique(df$measure)))
  # SAVE FOR FUTURE INCLUSION OF NET INCOME & COMPARATIVE INCOME
  # base <- data.frame(display_order = c(20,21,22,23,24,25,26,27), 
  #                    measure = m, 
  #                    base_weight = c(30,10,10,10,10,10,10,10), 
  #                    stringsAsFactors = FALSE)
  # SAVE FOR FUTURE INCLUSION OF  NET INCOME & COMPARATIVE INCOME
  base <- data.frame(display_order = c(23,26,27,24,22,25),  #order FTE then HC
                     measure = m, 
                     base_weight = c(10,10,10,10,10,10), 
                     stringsAsFactors = FALSE) 
  
  
  criteria_4 <- matc %>%
    mutate(fiscal_year = as.integer(fiscal_year)) %>%
    crossing(base) %>%
    left_join(df)
  
  
  rm(cnt,
     pct_cnt,
     sch_comp,
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