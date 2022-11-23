library(tidyverse)
library(magrittr)
library(openxlsx)
library(readxl)
library(writexl)
library(reshape2)
library(skimr)
library(janitor)
library(lubridate)

mega_data <- read_excel("C:/Users/slee/OneDrive - Ventura Foods/Ventura Work/SCE/Project/FY 23/Forecast Accuracy and Bias/Mega Data/mega_data_upto_oct_2022.xlsx")

959214

# delete NA in the first column (date divider for Wilson)



mega_data %>% 
  data.frame() -> mega_data

save(mega_data, file = "mega_data.rds")
load("mega_data.rds")


### modify mega_data
# match with dsx
# delete Oct to add my dsx from R

# How to do this? 
# dsx file into Excel (only first 10 cols)
# mega data file into Excel (only first 10 cols)
# compare.. and match these two

colnames(mega_data) <- mega_data[1, ]
mega_data[-1, ] -> mega_data

mega_data %>% 
  janitor::clean_names()



