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
  data.frame() %>% 
  dplyr::relocate(forecast_month_year_code, forecast_month_year_code, product_manufacturing_location_name, location_no, product_label_sku_code,
                  product_label_sku_name, product_category_name, product_platform_name, product_group_code, product_manufacturing_line_area_no_code,
                  abc_4_id, safety_stock_id, mto_mts_gross_requirements_calc_method_id, adjusted_forecast_pounds_lbs, adjusted_forecast_cases,
                  stat_forecast_pounds_lbs, stat_forecast_cases, product_label_sku_code, mto_mts_gross_requirements_calc_method_id, primary_channel_id,
                  sub_segment_id)



# There are number of duplicated columns, and some columns I have no idea. 
# Need to discuss with Wilson before I move

