# Author: Saeesh Mangwani
# Date: 2021-12-28

# Description: Helper functions for invoice generation process

# ==== Libraries ====
library(dplyr)
library(stringr)
library(openxlsx)

# ==== Formatting and writing the activity table ====
writeActivityTable <- function(curr_invoice, invoice_template){
  curr_activities <- curr_invoice %>% 
    mutate(`S. No.` = 1:nrow(.)) %>% 
    mutate(`Taxable Value` = Final.Amount) %>% 
    select(
      `S. No.`,
      'Service Description' = Service.Name,
      'HSN/SAC' = HSN, 
      'Rate' = Price, 
      'Pax' = Qty, 
      'Discount' = `Discount/TAC`, 
      `Taxable Value`,
      'Total' = Final.Amount)
  
  # Writing activities to sheet
  writeData(laca_invoice, 
            sheet = 1, 
            x = curr_activities,
            startCol = 1,
            startRow = 19, 
            colNames = F)
  
  # Writing total amounts
  writeData(laca_invoice, 
            sheet = 1, 
            x = sum(curr_activities$Total, na.rm = T),
            xy = c(7, 31))
  