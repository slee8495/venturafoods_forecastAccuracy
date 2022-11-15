library(tidyverse)
library(magrittr)
library(openxlsx)
library(readxl)
library(writexl)
library(reshape2)
library(skimr)
library(janitor)
library(lubridate)

# (Path Revision Needed) File read ----

dsx <- read_excel("S:/Global Shared Folders/Large Documents/S&OP/Demand Planning/Demand Planning Team/BI Forecast Backup/DSX Forecast Backup - 2022.11.14.xlsx")


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
                  stat_forecast_pounds_lbs, stat_forecast_cases, product_label_sku_code, primary_channel_id,
                  sub_segment_id) -> dsx



############ ETL ##############

# Forecast Month Name & Calendar Year

dsx %>% 
  dplyr::mutate(forecast_month_year_code = lubridate::ym(forecast_month_year_code)) %>%
  dplyr::mutate(forecast_month_name = lubridate::month(forecast_month_year_code),
                calendar_year = lubridate::year(forecast_month_year_code)) -> dsx


# MTO-MTS
dsx %>% 
  dplyr::mutate(mto_mts = recode(mto_mts_gross_requirements_calc_method_id, "Use Greater of Forecast and Customer Orders" = "MTS", "Ignore Forecast" = "MTO")) -> dsx


# Label
separate(dsx, product_label_sku_code, c("rm", "label"), sep = "-") %>% 
  dplyr::mutate(product_label_sku_code = paste0(rm, label)) %>% 
  dplyr::select(-rm) -> dsx

  

