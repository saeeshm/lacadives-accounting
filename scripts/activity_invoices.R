# Author: Saeesh Mangwani
# Date: 2021-12-27

# Description: Autogenerating invoices from the Lacadives activity tracker

# ==== Libraries ====
library(dplyr)
library(stringr)
library(openxlsx)
library(english)

# ==== Reading datas ====

# Activity sheet
activities <- openxlsx::read.xlsx('files/lacadives_december_invoice_reference.xlsx',
                                'Activities')

# Prices sheet
prices <- openxlsx::read.xlsx('files/lacadives_december_invoice_reference.xlsx',
                              'Rates')

# ==== Formatting activity sheet for invoice generation ====

# For rows where invoice number is missing, replacing the empty invoice, client
# and name columns with data from the previous columns (since these cells are
# only left empty when new events are added for the same client)
for (i in 1:nrow(activities)){
  # If invoice id is missing
  if(is.na(activities[i,1])){
    # If name and CID are also missing, this is just an added activity, so we
    # can infill it with the invoice id of the row above
    if(is.na(activities[i, 2])){
      activities[i, 1] <- activities[i-1, 1]
      activities[i, 2] <- activities[i-1, 2]
      activities[i, 3] <- activities[i-1, 3]
    }
  }
}

# For rows where prices are missing, adding a default price if available
for (i in 1:nrow(activities)){
  if(is.na(activities[i,7])){
    # Reading activity name
    aname <- activities[i, 4]
    # Checking for default price
    if(aname %in% prices$Service.Name){
      def_price <- prices %>% 
        filter(prices$Service.Name %in% aname) %>% 
        pull(Price)
    }else if(aname %in% prices$Service.Short.Name){
      def_price <- prices %>% 
        filter(prices$Service.Short.Name %in% aname) %>% 
        pull(Price)
    } 
    
    # If it exists, adding to the activity sheet
    if(!is.na(def_price)) activities[i, 7] <- def_price
  }
}

# Default tax rate
for (i in 1:nrow(activities)){
  if(is.na(activities[i,9])){
    # Reading activity name
    aname <- activities[i, 4]
    # Checking for default tax
    if(aname %in% prices$Service.Name){
      def_taxrate <- prices %>% 
        filter(prices$Service.Name %in% aname) %>% 
        pull(Default.Tax.Rate)
    }else if(aname %in% prices$Service.Short.Name){
      def_taxrate <- prices %>% 
        filter(prices$Service.Short.Name %in% aname) %>% 
        pull(Default.Tax.Rate)
    } 
    # If it exists, adding to the activity sheet
    if(!is.na(def_taxrate)) activities[i, 9] <- def_taxrate
  }
}

# Default HSN Numbers
for (i in 1:nrow(activities)){
  if(is.na(activities[i,5])){
    # Reading activity name
    aname <- activities[i, 4]
    # Checking for default HSN
    if(aname %in% prices$Service.Name){
      def_HSN <- prices %>% 
        filter(prices$Service.Name %in% aname) %>% 
        pull(Service.HSN.Number)
    }else if(aname %in% prices$Service.Short.Name){
      def_HSN <- prices %>% 
        filter(prices$Service.Short.Name %in% aname) %>% 
        pull(Service.HSN.Number)
    } 
    # If it exists, adding to the activity sheet
    if(!is.na(def_HSN)) activities[i, 5] <- def_HSN
  }
}

# Calculating total amounts
activities <- activities %>% 
  # Ensuring numeric types
  mutate(`Discount/TAC` = ifelse(is.na(`Discount/TAC`), 0, `Discount/TAC`)) %>% 
  mutate(Price = ifelse(is.na(Price), 0, Price)) %>% 
  mutate(Credit = ifelse(is.na(Credit), 0, Credit)) %>% 
  # Ensuring if Qty is missing, it is set to 1
  mutate(Qty = ifelse(is.na(Qty), 1, Qty)) %>% 
  # Calculating totals
  mutate(Final.Amount = ((Price*Qty) - `Discount/TAC`)) %>% 
  mutate(Final.Amount.with.GST = (Final.Amount + (Final.Amount*Tax.Rate))) 

# ==== Function that creates an invoice for a given activity ====

createInvoice <- function(invoice_template, invoice_activities, out_path){
  # Client name
  client_name <- unique(invoice_activities$Client.Name)
  writeData(invoice_template, 
            sheet = 1, 
            x = paste('Bill To:', client_name), 
            xy = c(1,12))
  
  
  # Formatting activity table
  activities_summary <- invoice_activities %>% 
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
  writeData(invoice_template, 
            sheet = 1, 
            x = activities_summary,
            startCol = 1,
            startRow = 19, 
            colNames = F)
  
  # Total amounts
  gross_amount <- sum(activities_summary$Total, na.rm = T)
  writeData(invoice_template, 
            sheet = 1, 
            x = gross_amount,
            xy = c(7, 31))
  
  # Total amount in words
  writeData(invoice_template,
            sheet = 1, 
            x = paste(english(gross_amount), 'only'),
            xy = c(3, 32))
  
  # Tax calculations
  perc18_taxable_value <- activities_summary %>% 
    filter(!`HSN/SAC` == 6109) %>% 
    pull(Total) %>% 
    sum(na.rm = T)
  tax_amount_9 <- perc18_taxable_value*0.09
  
  perc5_taxable_value <- activities_summary %>% 
    filter(`HSN/SAC` == 6109) %>% 
    pull(Total) %>% 
    sum(na.rm = T)
  tax_amount_2.5 <- perc5_taxable_value*0.025
  
  # Writing tax amounts
  writeData(invoice_template, sheet = 1, x = tax_amount_9, xy = c(8, 35))
  writeData(invoice_template, sheet = 1, x = tax_amount_9, xy = c(8, 36))
  writeData(invoice_template, sheet = 1, x = tax_amount_9*2, xy = c(8, 37))
  writeData(invoice_template, sheet = 1, x = tax_amount_2.5, xy = c(8, 38))
  writeData(invoice_template, sheet = 1, x = tax_amount_2.5, xy = c(8, 39))
  writeData(invoice_template, sheet = 1, x = tax_amount_2.5*2, xy = c(8, 40))
  
  # Final amount
  final_amount <- gross_amount + (tax_amount_9*2) + (tax_amount_2.5*2)
  writeData(invoice_template, sheet = 1, x = final_amount, xy = c(8, 41))
  
  # Saving workbook
  saveWorkbook(invoice_template, out_path, overwrite = T)
}

# ==== Iteratively generating invoices for each client ====
activities %>% 
  # Not generating invoices for those that haven't been assigned an invoice id
  filter(!is.na(Invoice.ID)) %>% 
  mutate(InvoiceIDGrp = Invoice.ID) %>% 
  # Grouping by invoice ID
  group_by(InvoiceIDGrp) %>% 
  # Iterating over each invoice
  group_walk(~{
    # setting filename as invoice id
    print(.x)
    fname <- paste0(str_replace_all(unique(.x$Invoice.ID), '\\/', '_'), '.xlsx')
    print(fname)
    out_path <- file.path('invoices', 'xlsx', fname)
    print(out_path)
    invoice_template <- openxlsx::loadWorkbook('files/lacadives_invoice_template.xlsx')
    invoice <- createInvoice(laca_invoice, .x, out_path)
  })


# ==== Writing to disk ====





