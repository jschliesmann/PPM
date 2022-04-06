# ---------------------------------------------------------------------#
#  Purpose: Build Program Performance Matrix - Program Uniqueness Criteria
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
#        1. Must download the lastest Access DB from Survey Data option at https://nces.ed.gov/ipeds/use-the-data
#        2. Update the pivot table filters in WTCS - Data Source.xlsx
#        3. Running the script under 64-bit Version of R requires a 64-bit Access database driver. 
#             If the odbcDriverConnect() fails, install as administrator: AccessDatabaseEngine_X64.exe
#             downloaded @ https://www.microsoft.com/en-US/download/details.aspx?id=13255
#        4. Manually update the list of Pathway Credentials
#
#  Notes:
#        1.   Ipeds Package may not be available for latest version of R.  Instead
#               of installing a 'compiled' package install it from the source, but   
#               you will need RTools.exe to install the source package.
#        2. dynamically applies most recient ipeds year to most recient fiscal year 
#               after extracting ipeds data going back 4 years from most recient fiscal year
#        3. THE IPEDS PIECE NEEDS MONITORING
#               
#               
#  ToDo:
#        1. Use Filename function for IPEDS
#
# ----------------------------------------------------------------------#

processProgramUniqueness <- function(fiscal_year){

  
#~~~~~~~~~~~~~~ USER EDITS ~~~~~~~~~~~~~~~~~~~~~####
#~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~####
  
  
FOLDER_PU <- paste0(FOLDER,'\\Criteria - Program Uniqueness\\')


#PGM307 <- paste0(FOLDER_PU,'PGM307 (cipcodes).xls')  #unnecessary.. cip code is now in MATC Programs tab of data source.
CLI330_LIST <- list.files(FOLDER_PU,pattern = 'CLI330', full.names = T)  #download from portal 

#An alternative to accessing IPEDS from the Access Database directly, 
#   the IPEDS R Package could work.
#TO INSTALL IPEDS PACKAGE (not currently working with R 3.5.x)
#to install RTools.exe  http://cran.r-project.org/bin/windows/Rtools/
#to install ipeds       devtools::install_github('jbryer/ipeds')
# need mdbtools... but little support for application in windows environment


#~~~~~~~~~~~~~~ Setup ~~~~~~~~~~~~~~~~~~~~~~~~~~####
#~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~####

library(RODBC)


#~~~~~~~~~~~~~~ Functions ~~~~~~~~~~~~~~~~~~~~~~####
#~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~####

getIpedsPrograms <- function(ipeds_year,cip_detail = FALSE){
  #args: fiscal_year = 4 digit numeric
  #      cip_detail = Boolean: T uses 7-digit, F uses 5-digit 
  #      cip_detail <- FALSE

  #if it generates this error see requirments at top of script:
      #Error in sqlQuery(con, paste0("SELECT * FROM HD", ipeds_year_short), stringsAsFactors ==  : 
      #first argument is not an open RODBC channel
 #ipeds_year <- 201819; cip_detail <- FALSE
  
  #some ipeds data tables are labeld with the leading year of an academic year
  ipeds_year_short <- substr(ipeds_year,1,4)  
  
  #This will not communicate if the ODBC driver failed
  con <- tryCatch(odbcDriverConnect(paste0("Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=",
                                                  paste0(IPEDS_FOLDER,'IPEDS',ipeds_year,'.accdb'))),
                  error = function(e){
                    print('odbcDriverConnect() failed, See script requirements for more details.')
                  },
                  warning = function(w) {
                    print(paste('The' ,ipeds_year, 'IPEDS dataset ',paste0(IPEDS_FOLDER,'IPEDS',ipeds_year,'.accdb'),' was not found'));
                    return(0)})
  

  
  if(RODBC:::odbcValidChannel(con) == TRUE){
  
  cat('Retrieving the',ipeds_year, 'IPEDS dataset...\n')
  #IPEDS changed coding to include the word 'County' in countynm... adding the filter to SQL is not advised
  regional_schools <- sqlQuery(con,paste0("SELECT * FROM HD",ipeds_year_short),stringsAsFactors = FALSE) %>%
    rename_all(tolower) %>%
    filter(stabbr == 'WI',
           str_detect(countynm,
                      paste(c('Milwaukee',
                              'Waukesha',
                              'Ozaukee',
                              'Washington'),
                            collapse = '|'))
    ) %>%
    select(unitid,instnm,countynm) %>%
    unique() 
  
  regional_schools$instnm[grep('^Bryant & Stratton College',regional_schools$instnm)] <- 'Bryant & Stratton College'
    

  # This is not used
  # get CIP 2010 to MATC Program Number crosswalk
  # extractPgm307 is a function in Client Reporting Functions.R file
  # xw <- extractPgm307(PGM307) %>%
  #   mutate(cip_code = case_when(cip_detail == FALSE ~ substr(cip_code,1,5),
  #                               cip_detail == TRUE ~ cip_code,
  #                               TRUE ~ 'ERROR')) 
  
  
  programs <- sqlQuery(con,paste0("SELECT * FROM C",ipeds_year_short,"_A"),stringsAsFactors = FALSE) %>%
    rename_all(tolower) %>%  
    mutate(cip_code_7 = sub("(.{2})(.*)", "\\1.\\2",str_pad(as.character(cipcode*10000),6,'left','0')),  #convert numeric cip to character
           cip_code = case_when(cip_detail == FALSE ~ substr(cip_code_7,1,5),
                                cip_detail == TRUE ~ cip_code_7,
                                TRUE ~ 'ERROR')) #This option will combine Diploma and PS programs for counting frequency
  
 
  odbcClose(con)
  
 
   df <- regional_schools %>%
    left_join(programs) %>%
    filter(!is.na(instnm),
           unitid != 239248,  #MATC
           awlevel %in% c(1,2,3,13),
           ctotalt > 0
           ) %>% 
    select(instnm,cip_code) %>% 
    unique() %>%
    group_by(cip_code) %>%
    count()
  #Note: Bryant & Stratton College has multiple unitids, they are identified by campus
  
  #---- ERROR CHECKING
  #WARNINGS: non-matching CIP / Program Numbers
  # chk <-   programs %>%
  #   left_join(xw, by=cip_code) %>%
  #   filter(unitid == 239248) %>%
  #   select(unitid,program_number,program_title,cipcode,cip_code) %>%
  #   unique()
  #REVIEW COUNT by CIP CODE
  # chk <- programs %>%
  #   right_join(regional_schools) %>%
  #   filter(cip_code == '52.07')

  
  program_count <- matc %>%
    mutate(cip_code = substr(cip_code,1,5)) %>%
    left_join(df) %>% 
    group_by(program_number) %>%         #Needs to check that dup program numbers do not exist due to multiple CIP Code
    summarise(n = sum(n,na.rm = TRUE),   #needed for score calculation 
              #value = as.character(n),   
              ipeds_year = as.integer(ipeds_year_short)+1) %>%  #uadd 1 to match fiscal year naming convention
    ungroup() 
  
  rm(df,
     regional_schools,
     programs,
     con,
     ipeds_year_short)
  
  return(program_count)
  
  } else {return(data.frame("program_number"= '', "n" = 0, "ipeds_year" ='', stringsAsFactors = FALSE))}  #failed ODBC Connection
}

    ####----  OLD FUNCTIONS ----####
    # extractCli330 <- function(file_name){
    #   #gets special population ration from WTCS CLI330 report
    #   #agrs: file_name is the full path to the CLI330.xls report
    #   
    #   if(file.exists(file_name)){
    #     
    #     yr <- unlist( read_excel(file_name,range = 'B7:B7',col_names = FALSE) %>%
    #       select(fiscal_year = 1) %>%
    #       mutate(fiscal_year = as.numeric(gsub("[^0-9\\.]", "", fiscal_year))))
    #     
    #     df <- read_excel(file_name,skip = 8) %>%
    #       select(program_number = 1,
    #              program_title = 2,
    #              total_enrollment = 4,
    #              total_spec_pop = 21) %>%
    #       na.omit() %>%
    #       mutate(fiscal_year = yr,
    #              spec_pop_ratio = round(total_spec_pop / total_enrollment,2)) 
    #       
    #     #rm(yr, file_name)
    #     return(df)
    #   }else {stop('File not Found: Check path and file name.\n')}
    #   
    # }
  
    # extractPgm307 <- function(file_name){
    #   # Function: converts the PGM307.xls WTCS report to a data friendly format
    #   # @PARMA file_name: Full path to the PGM307.xls 
    #   # this function is also in MATC Custom Functions.R
    #   
    #   if(file.exists(file_name)){
    #     df <- read_excel(file_name,skip = 10) %>%
    #       select(program = 1,
    #              program_title = 2,
    #              cip_code = 8) %>%
    #       na.omit() %>%
    #       separate(program_title, c('program_title','remove','program_cluster','program_pathway'),sep = '[\n\r]') %>%
    #       mutate(program_number = str_replace_all(program, "-", ""),
    #              program_cluster = str_remove(program_cluster,'CLUSTER: '),
    #              program_pathway = str_remove(program_pathway,'PATHWAY: ')) %>%
    #       select(program,
    #              program_number,
    #              cip_code,
    #              program_title,
    #              program_cluster,
    #              program_pathway)
    # 
    #     
    #     return(df)
    #   }else {stop('File not Found: Check path and file name.\n')}
    #   
    # }


#~~~~~~~~~~~~~~ Program Uniqueness ~~~~~~~~~~~~~~~~####
#~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~####


#-------Identical Programs within WTCS------------####
inWtcs <- read_excel(WTCS,sheet = 'WTCS Programs') %>%
    #group_by(fiscal_year,program_number) %>%  #now completed server side (in cube)
    #count() %>%
    #ungroup() %>%
    right_join(matc) %>%
    mutate(n = ifelse(is.na(program_count),0,program_count),  
           value = as.character(n),
           fiscal_year = as.integer(fiscal_year),
           measure = 'Identical Programs within WTCS',
           benchmark = '2',
           score = case_when(n < 2 ~ 10,
                                n < 4 ~ 7,
                                n < 8 ~ 5,
                                n < 15 ~ 3,
                                TRUE ~ 0
                                ),
           available_points = 10) %>%
    select(fiscal_year,
           program_number,
           measure,
           value,
           benchmark,
           score,
           available_points)  





#-------Identical programs in MATC region---------####
#This may need a check if the most recient Ipeds year is not the same as reporting year. 
#  Which is likely if this is run prior to the new year.
#This 



df <- getIpedsPrograms(IPEDS_YEAR) %>%
   bind_rows(getIpedsPrograms(IPEDS_YEAR-101)) %>%
   bind_rows(getIpedsPrograms(IPEDS_YEAR-202)) %>%
   bind_rows(getIpedsPrograms(IPEDS_YEAR-303)) %>%
  mutate(gap = min(FISCAL_YEAR - max(as.integer(ipeds_year))),
         fiscal_year = as.character(as.integer(ipeds_year)+gap)) %>%  #links the most recient IPEDS data with most recient FY
  right_join(matc) %>%
  mutate(measure = 'Identical Programs in MATC Region',
         fiscal_year = as.integer(fiscal_year), #this is dumb, but matc df needs fiscal_year as character type
          benchmark = '0',
          score = case_when(is.na(n) ~ 10,
                            n < 1 ~ 10,
                            n < 2 ~ 7,
                            n < 6 ~ 5,
                            n < 10 ~ 3,
                            TRUE ~ 0
          ),
          value = ifelse(is.na(n),'0',as.character(n)),
          available_points = 10) %>%
  select(fiscal_year,
         ipeds_year,  #IPEDS Year is the calendaer year of the fall term 
         program_number,
         measure,
         value,
         benchmark,
         score,
         available_points)
   



# fiscal_year = case_when(IPEDS_YEAR - ipeds_year == 0 ~ FISCAL_YEAR,
#                         IPEDS_YEAR - ipeds_year == 101 ~ FISCAL_YEAR-1,
#                         IPEDS_YEAR - ipeds_year == 202 ~ FISCAL_YEAR-2,
#                         IPEDS_YEAR - ipeds_year == 303 ~ FISCAL_YEAR-3,
#                         IPEDS_YEAR - ipeds_year == 404 ~ FISCAL_YEAR-4,
#                         TRUE ~ 9999),
table(df$value,df$score)

print("IPEDS Year applied to which FY")
print("row = FY, column = ipeds year")
print(table(df$fiscal_year,df$ipeds_year))

 # chk <- df %>%
 #   filter(fiscal_year == '9999')
 # table(chk$program_number,chk$fiscal_year)

 
inRegion <- df %>%
   select(-ipeds_year)
 
rm(df, 
   IPEDS_YEAR,
   IPEDS_FOLDER)


#-------part of a Pathway Credential---------------####
df <- read_excel(WTCS,sheet = 'Pathway Cred',col_types = 'text') %>%
  mutate(tmp = program_number) %>%
  gather('fiscal_year','value',2:length(.)) %>%
  filter(!is.na(value)) 

pathCred <- read_excel(WTCS,sheet = 'MATC Programs') %>%
  left_join(df) %>%
  mutate(fiscal_year = as.numeric(fiscal_year),
         measure = 'Part of Pathway Credential',
         value = ifelse(is.na(value),'No','Yes'),
         benchmark = 'Yes',
         score = ifelse(value == 'No',0,20),
         available_points = 20) %>%
  select(fiscal_year,
         program_number,
         measure,
         value,
         benchmark,
         score,
         available_points)
  
rm(df)


#-------Has High School Dual enrolled student in a course-####
df <- read_excel(WTCS,sheet = 'MATC Programs') %>%
  mutate(aid_code = substr(program_number,1,2),
         instructional_area = substr(program_number,3,5)) %>%
  select(fiscal_year,program_number,aid_code,instructional_area) %>%
  unique()

dualEnrl <- read_excel(WTCS,sheet = 'Dual Enrollment') %>%
  full_join(df) %>%
  mutate(fiscal_year = as.numeric(fiscal_year),
         measure = 'Dual Enrolled High School Student',
         value = ifelse(is.na(course_credits),'No','Yes'),
         benchmark = 'Yes',
         score = ifelse(is.na(course_credits),0,20),
         available_points = 20 )%>%
  filter(!is.na(program_number)) %>%
  select(fiscal_year,
         program_number,
         measure,
         value,
         benchmark,
         score,
         available_points)

rm(df)

#-------Industry validated Credential--------------####
  #note: one excel sheet containes 2 tables.  Assusmes left most table is longer
df <- read_excel(WTCS,sheet = 'Industry Validated Cred') %>%
  select(fiscal_year = 1,
         program_number = 2) %>%
  na.omit()

# Old version (pivot tables) of cube data needed to seperate TSA & Apprenticeships
# df2 <- read_excel(WTCS,sheet = 'Industry Validated Cred') %>%
#   select(fiscal_year = 4,
#          program_number = 5) %>%
#   na.omit()

industCred <- df %>% 
  # bind_rows(df2) %>%
  mutate(value = 'Yes') %>%
  right_join(matc) %>%
  mutate(fiscal_year = as.numeric(fiscal_year),
         measure = 'Industry Validated Credential',
         value = ifelse(is.na(value),'No',value),
         benchmark = 'Yes',
         score = ifelse(value == 'Yes',20,0),
         available_points = 20
         ) %>%
  select(fiscal_year,
         program_number,
         measure,
         value,
         benchmark,
         score,
         available_points)

rm(df,df2)

#-------Special Populations------------------------####

#to deleete after confirming sapply works
# df <- extractCli330(paste0(FOLDER_PU,'CLI330 - ',FISCAL_YEAR,'.xls')) %>%
#   bind_rows(extractCli330(paste0(FOLDER_PU,'CLI330 - ',FISCAL_YEAR - 1,'.xls'))) %>%
#   bind_rows(extractCli330(paste0(FOLDER_PU,'CLI330 - ',FISCAL_YEAR - 2,'.xls'))) 

#all CLI330 must be final versions


warn <- sapply(CLI330_LIST, read_excel, simplify=FALSE) %>% 
    bind_rows(.id = "id") %>%
    select(id = 1,
           chk = 8) %>%
    filter(substr(chk,1,7) == 'Warning') %>%
    mutate(fiscal_year = as.integer(substr(id,nchar(id)-7,nchar(id)-4)))

if(nrow(warn) >= 1){
  warning("CLI330 - ",warn$fiscal_year[1],".xls is NOT a finalized file")
}
           
df <- sapply(CLI330_LIST, read_excel, simplify=FALSE) %>% 
  bind_rows(.id = "id") %>%
  select(id = 1,
         program_number = 2,
         total_enrollment = 5,
         total_spec_pop1 = 22,
         total_spec_pop2 = 23,  #accomidates non-final CLI330 reports
         ) %>%
  filter(!is.na(as.integer(program_number))) %>%
  mutate(fiscal_year = as.integer(substr(id,nchar(id)-7,nchar(id)-4)),
         total_spec_pop = ifelse(is.na(total_spec_pop2),total_spec_pop1,total_spec_pop2), #use col21 if col22 is NA
         spec_pop_ratio = round(as.integer(total_spec_pop) / as.integer(total_enrollment),2)) %>%
  select(-total_spec_pop1,total_spec_pop2)


#df2 for future development to source from cube
df2 <- read_excel(WTCS,sheet = 'Special Populations') %>%
  select(fiscal_year = 1,
         program_number = 2) %>%
  na.omit()

specPop <- df %>%
  getBenchmark(col = spec_pop_ratio,p = .75) %>%
  mutate(measure = 'Special Populations',
         value = percent(spec_pop_ratio,0),
         available_points = 20,
         score = case_when(spec_pop_ratio > benchmark ~ available_points,     #top %25 of programs gets all points
                           spec_pop_ratio <= benchmark ~ spec_pop_ratio/benchmark*available_points,  
                           is.na(spec_pop_ratio) ~ 0,
                           TRUE ~ 0),
         benchmark = percent(benchmark,0)) %>%
  select(fiscal_year,
         program_number,
         measure,
         value,
         benchmark,
         score,
         available_points)

rm(df,df2,warn)



#-------Bring it all together----------------------####

df <- inWtcs %>%
  bind_rows(inRegion) %>%
  bind_rows(pathCred) %>%
  bind_rows(dualEnrl) %>%
  bind_rows(industCred) %>%
  bind_rows(specPop) %>%
  mutate(capital_request_score = 0)
  

m <- sort(as.vector(unique(df$measure)))
base <- data.frame(display_order = c(17,15,14,18,16,19), 
                   measure = m, 
                   base_weight = c(20,10,10,20,20,20), 
                   stringsAsFactors = FALSE)


criteria_3 <- read_excel(WTCS,sheet = 'MATC Programs', col_types = 'text') %>%
  mutate(fiscal_year = as.integer(fiscal_year)) %>%
  crossing(base) %>%
  left_join(df) %>%
  unique()

#remove unnecessary data
# rm(inRegion,
#    pathCred,
#    dualEnrl,
#    industCred,
#    specPop,
#    m,base,df,
#    FOLDER_PU,
#    PGM307)



#-------Error Checking-----------------------------####

    table(criteria_3$measure,criteria_3$base_weight) 
    # Expecting: 
    #                                        10  20
    #Course for High School Dual Enrollment   0 ###
    #Identical Programs in MATC Region      ###   0
    #Identical Programs within WTCS         ###   0
    #Industry Validated Credential            0 ###
    #Path of Pathway Credential               0 ###
    #Special Populations                      0 ###
    
    table(criteria_3$measure,criteria_3$display_order) 
    # Expecting: 
    #                                        14  15  16  17  18  19
    #Course for High School Dual Enrollment   0   0   0 ###   0   0
    #Identical Programs in MATC Region        0 ###   0   0   0   0
    #Identical Programs within WTCS         ###   0   0   0   0   0
    #Industry Validated Credential            0   0   0   0 ###   0
    #Path of Pathway Credential               0   0 ###   0   0   0
    #Special Populations                      0   0   0   0   0 ###

return(criteria_3)
}
