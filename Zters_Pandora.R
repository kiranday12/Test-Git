#####Append all the rows in Zters Pandora excel #######

library(rJava)
library(xlsx)
library(dplyr)
library(tidyverse)
library(ggplot2)

# ZTRS_pandora_ads <- read.xlsx("C:/Users/kiran/Desktop/Zters Pandora Ads.xlsx",
#                               sheetIndex = 10,rowIndex = 12:33,colIndex = 2:10,
#                               header = T)


# installing the required libraries
library(readxl)


# specifying the path for file
path <- "C:/Users/kiran/Desktop"

# set the working directory
setwd(path)

# accessing all the sheets
sheet = excel_sheets("Zters Pandora Ads.xlsx")

# applying sheet names to dataframe names
data_frame = lapply(setNames(sheet, sheet),
                    function(x) read_excel("Zters Pandora Ads.xlsx", sheet=x,skip = 11,col_names = T,range = cell_rows(12:33),
                                           col_types = c("text","text","date","date","numeric","numeric","numeric","numeric","numeric")))

# attaching all dataframes together
ZTRS_pandora_ads = bind_rows(data_frame, .id="Sheet") %>%
  rename(`Collection Date` = Sheet) 

ZTRS_pandora_ads_insights <- ZTRS_pandora_ads %>%
  mutate(`Impressions Delivered` = coalesce(`Impressions Delivered`, 0),
         Clicks = coalesce(Clicks,0),
         Reach = coalesce(Reach,0)) %>%
  mutate(`Collection Date` = as.Date(`Collection Date`,format = "%m.%d.%y"))%>%
  group_by(`Component Name`,`Ad Comments`) %>%
  mutate(`Weekly Impressions Delivered` = `Impressions Delivered`-lag(`Impressions Delivered`,default = `Impressions Delivered`[1],order_by = `Collection Date`),
         `Weekly Clicks` = Clicks - lag(Clicks,default = Clicks[1],order_by = `Collection Date`),
         `Weekly Reach` = Reach - lag(Reach,default = Reach[1],order_by = `Collection Date`))
  

# temp1 <- ZTRS_pandora_ads_insights %>% filter(`Collection Date` == "2022-01-09") %>%
#   select(`Impressions Delivered`)
# 
# temp2 <- ZTRS_pandora_ads_insights %>% filter(`Collection Date` == "2022-01-16") %>%
#   select(`Impressions Delivered`)

###### Have a look at the data

ZTRS_pandora_ads_insights %>% glimpse()


###### Make charts for the data 

#####Impression Goal

ZTRS_pandora_ads_insights_chart1 <- ZTRS_pandora_ads_insights %>%
  ggplot()+
  geom_line(aes(x = `Collection Date`,y = `Impressions Delivered`,color = "red"))+
  facet_wrap(~`Component Name`+`Ad Comments`)

ZTRS_pandora_ads_insights_chart2 <- ZTRS_pandora_ads_insights %>%
  ggplot()+
  geom_line(aes(x = `Collection Date`,y = `Clicks`,color = "red"))+
  facet_wrap(~`Component Name`+`Ad Comments`)

ZTRS_pandora_ads_insights_chart3 <- ZTRS_pandora_ads_insights %>%
  ggplot()+
  geom_line(aes(x = `Collection Date`,y = `Reach`,color = "red"))+
  facet_wrap(~`Component Name`+`Ad Comments`)
