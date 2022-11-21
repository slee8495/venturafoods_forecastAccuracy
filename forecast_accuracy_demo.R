library(tidyverse)
library(magrittr)
library(openxlsx)
library(readxl)
library(writexl)
library(reshape2)
library(skimr)
library(janitor)
library(lubridate)

########################################### Original Resources Input ###############################################
# (Path Revision Needed) dsx File read ----
# Lag 0
dsx_lag_0 <- read_excel("S:/Global Shared Folders/Large Documents/S&OP/Demand Planning/Demand Planning Team/BI Forecast Backup/2022/DSX Forecast Backup - 2022.11.01.xlsx")

# Lag 1
dsx_lag_1 <- read_excel("S:/Global Shared Folders/Large Documents/S&OP/Demand Planning/Demand Planning Team/BI Forecast Backup/2022/DSX Forecast Backup - 2022.10.03.xlsx")

# Lag 2
dsx_lag_2 <- read_excel("S:/Global Shared Folders/Large Documents/S&OP/Demand Planning/Demand Planning Team/BI Forecast Backup/2022/DSX Forecast Backup - 2022.09.02.xlsx")



# Ending Inventory
end_inv <- read_excel("C:/Users/slee/OneDrive - Ventura Foods/Ventura Work/SCE/Project/FY 23/Forecast Accuracy and Bias/Sample report for me/Forecast Accuracy by Ship-Loc_(Example for New Metrics) for R.xlsx",
                      sheet = "End Inv")





########################################################## E.T.L ##########################################################
################################################ Original resources modify ################################################


########################## Ending Inventory

end_inv[-1, ] -> end_inv
colnames(end_inv) <- end_inv[1, ]
end_inv[-1, ] -> end_inv

end_inv %>% 
  janitor::clean_names() %>% 
  readr::type_convert() %>% 
  dplyr::select(-ship_loc_2) %>% 
  dplyr::mutate(sku = gsub("-", "", sku),
                end_month = as.Date(end_month, origin = "1899-12-30"))


########################## Lag 0 

dsx_lag_0[-1, ] -> dsx_lag_0
colnames(dsx_lag_0) <- dsx_lag_0[1, ]
dsx_lag_0[-1, ] -> dsx_lag_0


dsx_lag_0 %>% 
  janitor::clean_names() %>% 
  readr::type_convert() %>% 
  data.frame() %>% 
  dplyr::relocate(forecast_month_year_code, forecast_month_year_code, product_manufacturing_location_name, location_no, product_label_sku_code,
                  product_label_sku_name, product_category_name, product_platform_name, product_group_code, product_manufacturing_line_area_no_code,
                  abc_4_id, safety_stock_id, mto_mts_gross_requirements_calc_method_id, adjusted_forecast_pounds_lbs, adjusted_forecast_cases,
                  stat_forecast_pounds_lbs, stat_forecast_cases, product_label_sku_code, primary_channel_id,
                  sub_segment_id) -> dsx_lag_0


# Forecast Month Name & Calendar Year
dsx_lag_0 %>% 
  dplyr::mutate(forecast_month_year_code = lubridate::ym(forecast_month_year_code)) %>%
  dplyr::mutate(forecast_month_name = lubridate::month(forecast_month_year_code),
                calendar_year = lubridate::year(forecast_month_year_code)) -> dsx_lag_0


# MTO-MTS
dsx_lag_0 %>% 
  dplyr::mutate(mto_mts = recode(mto_mts_gross_requirements_calc_method_id, "Use Greater of Forecast and Customer Orders" = "MTS", "Ignore Forecast" = "MTO")) -> dsx_lag_0


# Label
separate(dsx_lag_0, product_label_sku_code, c("rm", "label"), sep = "-") %>% 
  dplyr::mutate(product_label_sku_code = paste0(rm, label)) %>% 
  dplyr::select(-rm) -> dsx_lag_0


# Foreceasted Month
dsx_lag_0 %>% 
  dplyr::mutate(forcasted_month = lubridate::floor_date(Sys.Date(), unit = "month") -1,
                forcasted_month = lubridate::floor_date(forcasted_month,  unit = "month")) -> dsx_lag_0

# Forcast per 
dsx_lag_0 %>% 
  dplyr::mutate(forecast_per_lag = lubridate::floor_date(Sys.Date(), unit = "month") -1,
                forecast_per_lag = lubridate::floor_date(forecast_per_lag,  unit = "month")) -> dsx_lag_0

# Lag
dsx_lag_0 %>% 
  dplyr::mutate(lag = "Lag 0") -> dsx_lag_0



########################## Lag 1

dsx_lag_1[-1, ] -> dsx_lag_1
colnames(dsx_lag_1) <- dsx_lag_1[1, ]
dsx_lag_1[-1, ] -> dsx_lag_1


dsx_lag_1 %>% 
  janitor::clean_names() %>% 
  readr::type_convert() %>% 
  data.frame() %>% 
  dplyr::relocate(forecast_month_year_code, forecast_month_year_code, product_manufacturing_location_name, location_no, product_label_sku_code,
                  product_label_sku_name, product_category_name, product_platform_name, product_group_code, product_manufacturing_line_area_no_code,
                  abc_4_id, safety_stock_id, mto_mts_gross_requirements_calc_method_id, adjusted_forecast_pounds_lbs, adjusted_forecast_cases,
                  stat_forecast_pounds_lbs, stat_forecast_cases, product_label_sku_code, primary_channel_id,
                  sub_segment_id) -> dsx_lag_1


# Forecast Month Name & Calendar Year
dsx_lag_1 %>% 
  dplyr::mutate(forecast_month_year_code = lubridate::ym(forecast_month_year_code)) %>%
  dplyr::mutate(forecast_month_name = lubridate::month(forecast_month_year_code),
                calendar_year = lubridate::year(forecast_month_year_code)) -> dsx_lag_1


# MTO-MTS
dsx_lag_1 %>% 
  dplyr::mutate(mto_mts = recode(mto_mts_gross_requirements_calc_method_id, "Use Greater of Forecast and Customer Orders" = "MTS", "Ignore Forecast" = "MTO")) -> dsx_lag_1


# Label
separate(dsx_lag_1, product_label_sku_code, c("rm", "label"), sep = "-") %>% 
  dplyr::mutate(product_label_sku_code = paste0(rm, label)) %>% 
  dplyr::select(-rm) -> dsx_lag_1

# Foreceasted Month
dsx_lag_1 %>% 
  dplyr::mutate(forcasted_month = lubridate::floor_date(Sys.Date() - 30, unit = "month") -1,
                forcasted_month = lubridate::floor_date(forcasted_month,  unit = "month")) -> dsx_lag_1

# Forcast per 
dsx_lag_1 %>% 
  dplyr::mutate(forecast_per_lag = lubridate::floor_date(Sys.Date(), unit = "month") -1,
                forecast_per_lag = lubridate::floor_date(forecast_per_lag,  unit = "month")) -> dsx_lag_1

# Lag
dsx_lag_1 %>% 
  dplyr::mutate(lag = "Lag 1") -> dsx_lag_1




########################## Lag 2

dsx_lag_2[-1, ] -> dsx_lag_2
colnames(dsx_lag_2) <- dsx_lag_2[1, ]
dsx_lag_2[-1, ] -> dsx_lag_2


dsx_lag_2 %>% 
  janitor::clean_names() %>% 
  readr::type_convert() %>% 
  data.frame() %>% 
  dplyr::relocate(forecast_month_year_code, forecast_month_year_code, product_manufacturing_location_name, location_no, product_label_sku_code,
                  product_label_sku_name, product_category_name, product_platform_name, product_group_code, product_manufacturing_line_area_no_code,
                  abc_4_id, safety_stock_id, mto_mts_gross_requirements_calc_method_id, adjusted_forecast_pounds_lbs, adjusted_forecast_cases,
                  stat_forecast_pounds_lbs, stat_forecast_cases, product_label_sku_code, primary_channel_id,
                  sub_segment_id) -> dsx_lag_2


# Forecast Month Name & Calendar Year
dsx_lag_2 %>% 
  dplyr::mutate(forecast_month_year_code = lubridate::ym(forecast_month_year_code)) %>%
  dplyr::mutate(forecast_month_name = lubridate::month(forecast_month_year_code),
                calendar_year = lubridate::year(forecast_month_year_code)) -> dsx_lag_2


# MTO-MTS
dsx_lag_2 %>% 
  dplyr::mutate(mto_mts = recode(mto_mts_gross_requirements_calc_method_id, "Use Greater of Forecast and Customer Orders" = "MTS", "Ignore Forecast" = "MTO")) -> dsx_lag_2


# Label
separate(dsx_lag_2, product_label_sku_code, c("rm", "label"), sep = "-") %>% 
  dplyr::mutate(product_label_sku_code = paste0(rm, label)) %>% 
  dplyr::select(-rm) -> dsx_lag_2

# Foreceasted Month
dsx_lag_2 %>% 
  dplyr::mutate(forcasted_month = lubridate::floor_date(Sys.Date() - 60, unit = "month") -1,
                forcasted_month = lubridate::floor_date(forcasted_month,  unit = "month")) -> dsx_lag_2

# Forcast per 
dsx_lag_2 %>% 
  dplyr::mutate(forecast_per_lag = lubridate::floor_date(Sys.Date(), unit = "month") -1,
                forecast_per_lag = lubridate::floor_date(forecast_per_lag,  unit = "month")) -> dsx_lag_2

# Lag
dsx_lag_2 %>% 
  dplyr::mutate(lag = "Lag 2") -> dsx_lag_2




############# combine 3 files into dsx ###########

rbind(dsx_lag_0, dsx_lag_1, dsx_lag_2) -> dsx
rm(dsx_lag_0, dsx_lag_1, dsx_lag_2)



################################################ dsx calculation part ################################################

# Ending FG Inventory






