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
dsx_lag_0 <- read_excel("S:/Global Shared Folders/Large Documents/S&OP/Demand Planning/Demand Planning Team/BI Forecast Backup/2022/DSX Forecast Backup - 2022.10.03.xlsx")

# Lag 1
dsx_lag_1 <- read_excel("S:/Global Shared Folders/Large Documents/S&OP/Demand Planning/Demand Planning Team/BI Forecast Backup/2022/DSX Forecast Backup - 2022.09.02.xlsx")

# Lag 2
dsx_lag_2 <- read_excel("S:/Global Shared Folders/Large Documents/S&OP/Demand Planning/Demand Planning Team/BI Forecast Backup/2022/DSX Forecast Backup - 2022.08.01.xlsx")


# https://edgeanalytics.venturafoods.com/MicroStrategy/servlet/mstrWeb?evt=3186&src=mstrWeb.3186&subscriptionID=A0C9BDC98545BD83C168548B86BC3B28&Server=ENV-268038LAIO1USE2&Project=VF%20Intelligent%20Enterprise&Port=39321&share=1
actual_inv <- read_excel("C:/Users/slee/OneDrive - Ventura Foods/Ventura Work/SCE/Project/FY 23/Forecast Accuracy and Bias/Sales History Analysis.xlsx")



########################################################## E.T.L ##########################################################
################################################ Original resources modify ################################################


actual_inv[-1, ] -> actual_inv
colnames(actual_inv) <- actual_inv[1, ]
actual_inv[-1, ] -> actual_inv

actual_inv %>% 
  janitor::clean_names() %>% 
  readr::type_convert() %>% 
  dplyr::rename(sku = product_label_sku) %>% 
  dplyr::mutate(sku = gsub("-", "", sku)) %>% 
  tidyr::separate(calendar_month_year, c("month", "year")) %>% 
  dplyr::mutate(month = recode(month, "Jan" = 1, "Feb" = 2, "Mar" = 3, "Apr" = 4, "May" = 5, "Jun" = 6, "Jul" = 7, "Aug" = 8, "Sep" = 9, "Oct" = 10, "Nov" = 11, "Dec" = 12)) %>% 
  dplyr::mutate(date = paste0(year, "-", month, "-", 1),
                date = as.Date(date)) %>% 
  dplyr::mutate(ref = paste0(location, "_", sku, "_", date)) -> actual_inv


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
  dplyr::mutate(forecast_per = lubridate::floor_date(Sys.Date(), unit = "month") -1,
                forecast_per = lubridate::floor_date(forecast_per,  unit = "month")) -> dsx_lag_0

# Lag
dsx_lag_0 %>% 
  dplyr::mutate(lag = "Lag 0") -> dsx_lag_0


# Filter for current month
dsx_lag_0 %>% 
  dplyr::mutate(date_filter = lubridate::floor_date(Sys.Date(), unit = "month") -1,
                date_filter = lubridate::floor_date(forecast_per,  unit = "month"),
                date_filter_2 = ifelse(date_filter != forecast_month_year_code, "no", "yes")) %>% 
  dplyr::filter(date_filter_2 == "yes") %>% 
  dplyr::select(-date_filter, -date_filter_2) -> dsx_lag_0




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
  dplyr::mutate(forecast_per = lubridate::floor_date(Sys.Date(), unit = "month") -1,
                forecast_per = lubridate::floor_date(forecast_per,  unit = "month")) -> dsx_lag_1

# Lag
dsx_lag_1 %>% 
  dplyr::mutate(lag = "Lag 1") -> dsx_lag_1


# Filter for current month
dsx_lag_1 %>% 
  dplyr::mutate(date_filter = lubridate::floor_date(Sys.Date(), unit = "month") -1,
                date_filter = lubridate::floor_date(forecast_per,  unit = "month"),
                date_filter_2 = ifelse(date_filter != forecast_month_year_code, "no", "yes")) %>% 
  dplyr::filter(date_filter_2 == "yes") %>% 
  dplyr::select(-date_filter, -date_filter_2) -> dsx_lag_1

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
  dplyr::mutate(forecast_per = lubridate::floor_date(Sys.Date(), unit = "month") -1,
                forecast_per = lubridate::floor_date(forecast_per,  unit = "month")) -> dsx_lag_2

# Lag
dsx_lag_2 %>% 
  dplyr::mutate(lag = "Lag 2") -> dsx_lag_2

# Filter for current month
dsx_lag_2 %>% 
  dplyr::mutate(date_filter = lubridate::floor_date(Sys.Date(), unit = "month") -1,
                date_filter = lubridate::floor_date(forecast_per,  unit = "month"),
                date_filter_2 = ifelse(date_filter != forecast_month_year_code, "no", "yes")) %>% 
  dplyr::filter(date_filter_2 == "yes") %>% 
  dplyr::select(-date_filter, -date_filter_2) -> dsx_lag_2


############# combine 3 files into dsx ###########

rbind(dsx_lag_0, dsx_lag_1, dsx_lag_2) -> dsx
rm(dsx_lag_0, dsx_lag_1, dsx_lag_2)



################################################ dsx calculation part ################################################

# ref
dsx %>% 
  dplyr::mutate(ref = paste0(location_no, "_", product_label_sku_code, "_", forecast_per)) -> dsx

# AS - Ordered Quantity
# AS - Origianl Ordered Quantity
# AS - Actual Shipped

actual_inv %>% 
  dplyr::group_by(ref) %>% 
  dplyr::summarise(as_ordered_quantity = sum(ordered_final_qty)) -> actual_inv_pivot_1

actual_inv %>% 
  dplyr::group_by(ref) %>% 
  dplyr::summarise(as_original_ordered_quantity = sum(ordered_original_qty)) -> actual_inv_pivot_2

actual_inv %>% 
  dplyr::group_by(ref) %>% 
  dplyr::summarise(actual = sum(cases)) -> actual_inv_pivot_3

dsx %>% 
  dplyr::left_join(actual_inv_pivot_1) %>% 
  dplyr::left_join(actual_inv_pivot_2) %>% 
  dplyr::left_join(actual_inv_pivot_3) -> dsx


rm(actual_inv, actual_inv_pivot_1, actual_inv_pivot_2, actual_inv_pivot_3)


# NA to 0 for all

dsx %>% 
  dplyr::mutate(adjusted_forecast_pounds_lbs = replace(adjusted_forecast_pounds_lbs, is.na(adjusted_forecast_pounds_lbs), 0),
                adjusted_forecast_cases = replace(adjusted_forecast_cases, is.na(adjusted_forecast_cases), 0),
                stat_forecast_pounds_lbs = replace(stat_forecast_pounds_lbs, is.na(stat_forecast_pounds_lbs), 0),
                stat_forecast_cases = replace(stat_forecast_cases, is.na(stat_forecast_cases), 0),
                as_ordered_quantity = replace(as_ordered_quantity, is.na(as_ordered_quantity), 0),
                as_original_ordered_quantity = replace(as_original_ordered_quantity, is.na(as_original_ordered_quantity), 0),
                actual = replace(actual, is.na(actual), 0)) -> dsx


# Abs. Error (Actual)
dsx %>% 
  dplyr::mutate(abs_error_actual = abs(actual - adjusted_forecast_cases),
                abs_error_actual = replace(abs_error_actual, is.na(abs_error_actual), 0)) -> dsx


# MAPE % (Actual)
dsx %>% 
  dplyr::mutate(mape_percent_actual = ifelse(abs_error_actual == 0, 0, abs_error_actual / actual),
                mape_percent_actual = replace(mape_percent_actual, is.na(mape_percent_actual), 1),
                mape_percent_actual = replace(mape_percent_actual, is.nan(mape_percent_actual), 1),
                mape_percent_actual = replace(mape_percent_actual, is.infinite(mape_percent_actual), 1)) -> dsx


# Accuracy % (Actual)
dsx %>% 
  dplyr::mutate(accuracy_percent_actual = ifelse(mape_percent_actual > 1, 0, (1 - mape_percent_actual))) -> dsx


# MAPE.Dec (Actual)
dsx %>% 
  dplyr::mutate(mape_dec_actual = mape_percent_actual * 100) -> dsx


# Wgtd_Error (Actual)
dsx %>% 
  dplyr::mutate(wgtd_error_actual = mape_dec_actual * actual) -> dsx



# Abs. Error (Final Order)
dsx %>% 
  dplyr::mutate(abs_error_final_order = abs(as_ordered_quantity - adjusted_forecast_cases)) -> dsx

# MAPE % (Final Order)
dsx %>% 
  dplyr::mutate(mape_percent_final_order = ifelse(abs_error_final_order == 0, 0, abs_error_final_order / actual),
                mape_percent_final_order = replace(mape_percent_final_order, is.na(mape_percent_final_order), 1),
                mape_percent_final_order = replace(mape_percent_final_order, is.nan(mape_percent_final_order), 1),
                mape_percent_final_order = replace(mape_percent_final_order, is.infinite(mape_percent_final_order), 1)) -> dsx

# Accuracy % (Final Order)
dsx %>% 
  dplyr::mutate(accuracy_percent_final_order = ifelse(mape_percent_final_order > 1, 0, (1 - mape_percent_final_order))) -> dsx


# MAPE.Dec (Final Order)
dsx %>% 
  dplyr::mutate(mape_dec_final_order = mape_percent_final_order * 100) -> dsx

# Wgtd_Error (Final Order)
dsx %>% 
  dplyr::mutate(wgtd_error_final_order = mape_dec_final_order * actual) -> dsx


# Abs. Error (Original Order)
dsx %>% 
  dplyr::mutate(abs_error_original_order = abs(as_original_ordered_quantity - adjusted_forecast_cases)) -> dsx

# MAPE % (Original Order)
dsx %>% 
  dplyr::mutate(mape_percent_original_order = ifelse(abs_error_original_order == 0, 0, abs_error_original_order / actual),
                mape_percent_original_order = replace(mape_percent_original_order, is.na(mape_percent_original_order), 1),
                mape_percent_original_order = replace(mape_percent_original_order, is.nan(mape_percent_original_order), 1),
                mape_percent_original_order = replace(mape_percent_original_order, is.infinite(mape_percent_original_order), 1)) -> dsx


# Accuracy % (Original Order)
dsx %>% 
  dplyr::mutate(accuracy_percent_original_order = ifelse(mape_percent_original_order > 1, 0, (1 - mape_percent_original_order))) -> dsx

# MAPE.Dec (Original Order)
dsx %>% 
  dplyr::mutate(mape_dec_original_order = mape_percent_original_order * 100) -> dsx


# Wgtd_Error (Original Order)
dsx %>% 
  dplyr::mutate(wgtd_error_original_order = mape_dec_original_order * actual) -> dsx



# Abs. Error (Original Order by Stat Forecast)
dsx %>% 
  dplyr::mutate(abs_error_original_order_by_stat_fc = abs(as_original_ordered_quantity - stat_forecast_cases)) -> dsx


# MAPE % (Original Order by Stat Forecast)
dsx %>% 
  dplyr::mutate(mape_percent_original_order_by_stat_fc = ifelse(abs_error_original_order_by_stat_fc == 0, 0, abs_error_original_order_by_stat_fc / actual),
                mape_percent_original_order_by_stat_fc = replace(mape_percent_original_order_by_stat_fc, is.na(mape_percent_original_order_by_stat_fc), 1),
                mape_percent_original_order_by_stat_fc = replace(mape_percent_original_order_by_stat_fc, is.nan(mape_percent_original_order_by_stat_fc), 1),
                mape_percent_original_order_by_stat_fc = replace(mape_percent_original_order_by_stat_fc, is.infinite(mape_percent_original_order_by_stat_fc), 1)) -> dsx


# Accuracy % (Original Order by Stat Forecast)
dsx %>% 
  dplyr::mutate(accuracy_percent_original_order_by_stat_fc = ifelse(mape_percent_original_order_by_stat_fc > 1, 0, (1 - mape_percent_original_order_by_stat_fc))) -> dsx

# MAPE.Dec (Original Order by Stat Forecast)
dsx %>% 
  dplyr::mutate(mape_dec_original_order_by_stat_fc = mape_percent_original_order_by_stat_fc * 100) -> dsx


# Wgtd_Error (Original Order by Stat Forecast)
dsx %>% 
  dplyr::mutate(wgtd_error_original_order_by_stat_fc = mape_dec_original_order_by_stat_fc * actual) -> dsx



dsx %>% 
  dplyr::select(-cust_ref_forecast_pounds_lbs, -cust_ref_forecast_cases, -ref, -forecast_month_year_code) -> dsx




# re-arrange the columns
dsx %>% 
  dplyr::relocate(forecast_month_name, calendar_year, forcasted_month, forecast_per, lag,
                  product_manufacturing_location_code, product_manufacturing_location_name, location_no, location_name,
                  product_label_sku_code,	label, mto_mts,	mto_mts_gross_requirements_calc_method_id, product_label_sku_name,
                  product_category_name, product_platform_name,	product_group_code,	product_manufacturing_line_area_no_code, abc_4_id,
                  safety_stock_id, adjusted_forecast_pounds_lbs, adjusted_forecast_cases, stat_forecast_pounds_lbs, stat_forecast_cases,
                  primary_channel_id, sub_segment_id,	segmentation_id, product_group_short_name, as_ordered_quantity,	as_original_ordered_quantity,
                  actual,	abs_error_actual,	mape_percent_actual, accuracy_percent_actual,	mape_dec_actual, wgtd_error_actual,	abs_error_final_order,
                  mape_percent_final_order,	accuracy_percent_final_order,	mape_dec_final_order,	wgtd_error_final_order,	abs_error_original_order,
                  mape_percent_original_order, accuracy_percent_original_order,	mape_dec_original_order, wgtd_error_original_order,
                  abs_error_original_order_by_stat_fc, mape_percent_original_order_by_stat_fc, accuracy_percent_original_order_by_stat_fc,
                  mape_dec_original_order_by_stat_fc,	wgtd_error_original_order_by_stat_fc) -> dsx



##########################################################################################################################################################
##########################################################################################################################################################
##########################################################################################################################################################

# (Path Revision Needed)  bring previouse month mega_data ----
load("C:/Users/slee/OneDrive - Ventura Foods/Stan/R Codes/Projects/Forecast Accuracy/venturafoods_forecastAccuracy/Mega Data/mega_data_by_r_10.2022.rds")

# combind this month dsx result to mega_data
rbind(mega_data_by_r, dsx) -> mega_data_by_r

# (Path Revision Needed) save new mega_data ----
save(mega_data_by_r, file = "C:/Users/slee/OneDrive - Ventura Foods/Stan/R Codes/Projects/Forecast Accuracy/venturafoods_forecastAccuracy/Mega Data/mega_data_by_r_10.2022.rds")



