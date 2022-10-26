library(tidyverse)
library(magrittr)
library(openxlsx)
library(readxl)
library(writexl)
library(reshape2)
library(skimr)
library(janitor)
library(lubridate)


sample_data <- read_excel("C:/Users/slee/OneDrive - Ventura Foods/Ventura Work/SCE/Project/FY 23/Forecast Accuracy and Bias/Sample report for me/Data sample (Kunal) - Forecast  Accuracy.xlsx")

colnames(sample_data) <- sample_data[1, ]
sample_data[-1, ] -> sample_data

sample_data %>% 
  data.frame() %>% 
  janitor::clean_names() %>% 
  readr::type_convert() -> sample_data

sample_data %>% 
  tidyr::pivot_wider(names_from = c(forecast_month_name, lag), values_from = c(net_pounds_adjusted_forecast_fiscal,
                                                                               cases_adjusted_forecast_fiscal,
                                                                               forecasted_mo, 
                                                                               ending_fg_inv,
                                                                               as_order_quantity,
                                                                               as_original_order_quantity,
                                                                               actual,
                                                                               abs_error,
                                                                               mape,
                                                                               accuracy,
                                                                               mape_dec,
                                                                               wgtd_error,
                                                                               abs_error_1,
                                                                               mape_1,
                                                                               accuracy_1,
                                                                               mape_dec_1,
                                                                               wgtd_error_1,
                                                                               abs_error_2,
                                                                               mape_2,
                                                                               accuracy_2,
                                                                               mape_dec_2,
                                                                               wgtd_error_2,
                                                                               abs_error_3,
                                                                               mape_3,
                                                                               accuracy_3,
                                                                               mape_dec_3,
                                                                               wgtd_error_3,
                                                                               abs_error_4)) %>% 
  data.frame() %>% 
  janitor::clean_names() %>% 
  readr::type_convert() -> sampled_data_1






# How write multiple data into one excel file 
openxlsx::createWorkbook("example") -> example
openxlsx::addWorksheet(example, "original_format")
openxlsx::addWorksheet(example, "wided_format")


openxlsx::writeDataTable(example, "original_format", sample_data)
openxlsx::writeDataTable(example, "wided_format", sampled_data)


openxlsx::saveWorkbook(example, file = "example.xlsx")
