library(tidyverse)
library(magrittr)
library(openxlsx)
library(readxl)
library(writexl)
library(reshape2)
library(skimr)
library(janitor)
library(lubridate)


dsx <- read_excel("S:/Global Shared Folders/Large Documents/S&OP/Demand Planning/Demand Planning Team/BI Forecast Backup/DSX Forecast Backup - 2022.11.07.xlsx")


dsx[-1, ] -> dsx
colnames(dsx) <- dsx[1, ]
dsx[-1, ] -> dsx


dsx %>% 
  janitor::clean_names() %>% 
  readr::type_convert() %>% 
  data.frame()


# I'm currently defining columns. 
# C:/Users/slee/OneDrive - Ventura Foods/Ventura Work/SCE/Project/FY 23/Forecast Accuracy and Bias/Sample report for me/Forecast Accuracy by Ship-Loc_(Example for New Metrics).xlsx
# There are number of duplicated columns, and some columns I have no idea. 
# Need to discuss with Wilson before I move

