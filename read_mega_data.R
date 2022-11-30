library(tidyverse)
library(magrittr)
library(openxlsx)
library(readxl)
library(writexl)
library(reshape2)
library(skimr)
library(janitor)
library(lubridate)

# mega_data <- read_excel("C:/Users/slee/OneDrive - Ventura Foods/Ventura Work/SCE/Project/FY 23/Forecast Accuracy and Bias/Mega Data/mega_data_upto_oct_2022.xlsx")


# mega_data %>% 
#   data.frame() -> mega_data

# save(mega_data, file = "mega_data.rds")
load("C:/Users/slee/OneDrive - Ventura Foods/Stan/R Codes/Projects/Forecast Accuracy/venturafoods_forecastAccuracy/Mega Data/mega_data_by_excel_10.2022.rds")


###############################################



colnames(mega_data) <- mega_data[1, ]
mega_data[-1, ] -> mega_data


mega_data %>% 
  filter(!is.na(Type)) %>% 
  janitor::clean_names() %>% 
  dplyr::select(-type, -mfg_loc, -ship_loc, -product_platform, -abc_2, -sku_type,	-combined_label, -ending_fg_inv,
                -concatenate,	-campus,	-work_days,	-canola_sku,	-scroll) %>% 
  readr::type_convert() %>% 
  dplyr::mutate(forecast_month_name = recode(forecast_month_name, "Jan" = 1, "Feb" = 2, "Mar" = 3, "Apr" = 4, "May" = 5, "Jun" = 6, "Jul" = 7, "Aug" = 8, "Sep" = 9, "Oct" = 10, "Nov" = 11, "Dec" = 12, "JUne" = 6, "July" = 7)) %>% 
  dplyr::rename(product_manufacturing_location_name = mfg_location_city,
                location_no = ship_location,
                product_label_sku_code = product_sku,
                product_label_sku_name = product_sku_name_10206_gns,
                product_category_name = product_category_name_105,
                product_platform_name = product_platform_desc,
                product_group_code = product_group,
                product_manufacturing_line_area_no_code = line_number,
                abc_4_id = abc4_code,
                safety_stock_id = safety_stock,
                mto_mts_gross_requirements_calc_method_id = mto_mts,
                adjusted_forecast_pounds_lbs = net_pounds_adjusted_forecast_fiscal,
                adjusted_forecast_cases = cases_adjusted_forecast_fiscal,
                mto_mts = mts_mto,
                primary_channel_id = primary_channel,
                segmentation_id = segment,
                sub_segment_id = sub_segment,
                forcasted_month = forecasted_mo,
                as_ordered_quantity = as_order_quantity,
                as_original_ordered_quantity = as_original_order_quantity,
                abs_error_actual = abs_error,
                mape_percent_actual = mape_percent,
                accuracy_percent_actual = accuracy_percent,
                mape_dec_actual = mape_dec,
                wgtd_error_actual = wgtd_error) %>% 
  dplyr::mutate(location_no = gsub("-", "", location_no),
                forcasted_month = as.Date(forcasted_month, origin = "1899-12-30"),
                forecast_per = as.Date(forcasted_month, origin = "1899-12-30")) %>% 
  dplyr::mutate(forecast_month_year_code = "",
                product_manufacturing_location_code = "",
                location_name = "",
                product_group_short_name = "",
                abs_error_final_order = "",
                mape_percent_final_order = "",	
                accuracy_percent_final_order = "",	
                mape_dec_final_order = "",
                wgtd_error_final_order = "",
                abs_error_original_order = "",
                mape_percent_original_order = "",
                accuracy_percent_original_order = "",
                mape_dec_original_order = "",
                wgtd_error_original_order = "",
                abs_error_original_order_by_stat_fc = "",
                mape_percent_original_order_by_stat_fc = "",
                accuracy_percent_original_order_by_stat_fc = "",
                mape_dec_original_order_by_stat_fc = "",
                wgtd_error_original_order_by_stat_fc = "") %>% 
  dplyr::select(-forecast_month_year_code) %>% 
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
                  mape_dec_original_order_by_stat_fc,	wgtd_error_original_order_by_stat_fc) -> a

 


#####################################################################################################

a %>% filter(!(forecast_month_name == 10 & calendar_year == 2022)) -> b



rbind(b, dsx) -> mega_data_by_r

# save(mega_data_by_r, file = "mega_data_by_r.rds")
# load("mega_data_by_r_10.2022.rds")


# mega data - lag column revise
mega_data_by_r %>% 
  dplyr::mutate(lag = recode(lag, "No Lag" = "Lag 0", "1 mo Lag" = "Lag 1", "1 Mo Lag" = "Lag 1", "2 mo Lag" = "Lag 2", "2 Mo Lag" = "Lag 2")) -> mega_data_by_r

save(mega_data_by_r, file = "mega_data_by_r_10.2022.rds")

