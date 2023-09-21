
### Update R if needed

# install.packages("installr")
# library(installr)
# updateR()


### Update packages if needed

# update.packages(checkBuilt=TRUE, ask=FALSE)


### Loading packages

list.of.packages <- c("DBI",
                      "odbc",
                      "openxlsx", 
                      "skimr", 
                      "corrr", 
                      "gtsummary", 
                      "janitor", 
                      "tidyverse", 
                      "RODBC", 
                      "haven", 
                      "flextable", 
                      "broom", 
                      "lmtest", 
                      "tcltk", 
                      "lubridate", 
                      "officer", 
                      "gridExtra", 
                      "eeptools")
new.packages <- list.of.packages[!(list.of.packages %in% installed.packages()[,"Package"])]
if(length(new.packages)>0) {install.packages(new.packages)}

lapply(list.of.packages, require, character.only = TRUE)



################################################################################################################

### Setting directory

setwd("Y:\\________________")
# dir.create(file.path(paste0("Outputs ",today)))


### Import data

SQL_table <- sqlQuery(odbcConnect("_____"), "SELECT DISTINCT * FROM ______________")

SQL_table_from_script <- dbGetQuery(dbConnect(odbc::odbc(), "_________"), statement = read_file("____________.sql"))

dta_data <- read_dta(file = "Y:\\________________.dta")

csv_data <- read.csv("Y:\\________________.csv", header = TRUE, stringsAsFactors=FALSE)

xlsx_data <- read.xlsx("Y:\\________________.xlsx", detectDates = TRUE) %>% 
  clean_names()


### Calendar table

week_calendar_full <- data.frame(date=seq((Sys.Date()-years(5)), Sys.Date(), by="day")) %>% 
  mutate(week=isoweek(date),month=format(date, "%b"),year=year(date)) 

week_calendar <- week_calendar_full %>%
  group_by(year,week) %>% 
  summarise(weekdate=min(date))


################################################################################################################

### Setting theme for charts

nice_chart <- list(
  # theme_minimal() +
  theme_classic() +
  theme(panel.grid.major.y = element_line(size=0.1, colour="grey")) + 
  theme(panel.grid.minor = element_line(colour="white", size=0)) + 
  theme(panel.grid.major.x = element_blank()) +
  theme(axis.text.x = element_text(angle = 45, hjust = 1, size=20, margin = margin(b = 2), colour="black")) +  
  theme(axis.text.y = element_text(size=20, margin = margin(l = 10), colour="black")) +  
  theme(axis.title = element_text(size=21, face="bold")) +
  theme(legend.text = element_text(color="black", size=20, face="italic", hjust=0)) +
  theme(legend.title = element_text(color="black", size=20, face="italic", hjust=0)))


### Nice colours

two_colours <- c( "#BB5566", "#004488")

many_colours <- c( '#4363d8', "#CCEBC5", "#F0027F", "#FFFF33", "#4DAF4A", "#B3CDE3", "#FF0000", '#fffac8', "#BC80BD",
                '#f58231', "#66C2A5","#FDCDAC" , "#FFD92F",   "#B3DE69", "#FCCDE5", "#000080","#D9D9D9", '#aaffc3',
                "#FFFFFF",'#46f0f0',"#FB8072","#A52A2A","#BEBADA")


### Flextable formatting

flextable_format <- function(x) {
  x <- x %>% 
    # mutate(across(where(is.numeric),as.character)) %>%
    # flextable() %>%
    align(align = "center", part = "all") %>%
    valign(valign = "center", part = "all") %>%
    bold(part = "header") %>%
    # hline(j = NULL, part = "header") %>%
    # hline_top(part = "header") %>%
    # vline() %>%
    # vline_left() %>%
    # hline() %>%
    font(fontname = 'Calibri', part = "all") %>%
    fontsize(part = "all", size = 8) %>%
    height_all(height = 0.24, part = "body") %>%
    # hrule(rule = "exact", part = "body") %>%
    padding(padding = 0, part = "all") %>%
    autofit()
  
}


### Function to output a dataframe to word

out_func <- function(func_dataframe, name) {
  
  func_flex <- func_dataframe %>%
    flextable() %>%
    flextable_format
  
  table_doc = read_docx() %>%
    body_add_flextable(func_flex)
  
  print(table_doc, target = (paste0(getwd(),"\\",name,".docx")))

}


################################################################################################################

### Deduplication

data <- data %>%
  group_by(variable) %>%
  mutate(dup=row_number()) %>%
  ungroup() %>%
  filter(dup==1) %>%
  select(-dup)


### Join data

combined_data <- data1 %>%
  left_join(data2,by=c('________'='________'))


### Summarise data

summarised_data <- test_data %>%
  group_by(________) %>%
  summarise(n = n(), average________ = mean(________))


### Export data

write.csv(combined_data,paste0(getwd(),"\\Outputs ",today,"\\","data",".csv"))


##### Quick Analyses
### Skim data

skim(linelist)


### Histogram

ggplot(data = linelist, mapping = aes(x = age_years)) +
  geom_histogram()


### Correlations (pearsons is default)

correlation_tab <- linelist %>%
  select(where(is.numeric)) %>%
  correlate() 


### Plot correlations

rplot(correlation_tab, print_cor = T)


### Breakdown by gender

linelist %>%
  select(gender, age_years, temp, hospital) %>%
  tbl_summary(by = gender,
              digits = all_continuous() ~ 0,
              label = list(
                age_years ~ "Age (years)",
                temp ~ "Temperature (C)",
                hospital ~ "Hospital"
              )) %>%
  add_n() %>%
  add_overall()


### Median and range for variable

linelist %>%
  select(ht_cm, outcome) %>%
  tbl_summary(by = outcome,
              statistic = ht_cm ~ "{median} ({min} - {max})")


### Chi-Squared

linelist %>%
  select(gender, outcome, fever:vomit) %>%
  tbl_summary(by = outcome) %>%
  add_p() %>%
  bold_p()


### t-test

linelist %>%
  select(wt_kg, outcome) %>%
  tbl_summary(by = outcome,
              statistic = wt_kg ~ "{mean} ({sd})") %>%
  add_p(wt_kg ~ "t.test") %>%
  bold_p()


### Kruskal test

linelist %>%
  select(age_years, outcome) %>%
  tbl_summary(by = outcome,
              statistic = age_years ~ "{median} ({p25} - {p75})") %>%
  add_p(age_years ~ "kruskal.test") %>%
  bold_p()
