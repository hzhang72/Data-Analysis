library(tidyverse)
library(readxl)
library(writexl)
library(openxlsx)

dfao <- read_excel("InvoiceReport_WithContainerNo_2021_Others.xlsx", 
                   sheet = "Sheet1")

df1 <- dfao %>% select(ThirdParty_CustomerID, ThirdParty_SiteID, Cost)
df1$Cost <- as.numeric(df1$Cost)

df11 <- df1 %>% replace(is.na(.), 0) %>% 
  group_by(ThirdParty_CustomerID, ThirdParty_SiteID) %>% 
  summarise('Auto Handling Revenue' = sum(Cost)) %>% 
  arrange(ThirdParty_SiteID, ThirdParty_CustomerID)

df11$ThirdParty_SiteID <- str_to_title(df11$ThirdParty_SiteID)
df11




dfav <- read_excel("InvoiceReport_WithContainerNo_2021_Others_valleyview.xlsx", 
                   sheet = "Sheet1")

df2 <- dfav %>% select(ThirdParty_CustomerID, ThirdParty_SiteID, Cost)
df2$Cost <- as.numeric(df2$Cost)

df22 <- df2 %>% replace(is.na(.), 0) %>% 
  group_by(ThirdParty_CustomerID, ThirdParty_SiteID) %>% 
  summarise('Auto Handling Revenue' = sum(Cost))

df22$ThirdParty_SiteID <- str_to_title(df22$ThirdParty_SiteID)
df22

dfa <- rbind(df11, df22)
dfa <- dfa %>% arrange(ThirdParty_SiteID, ThirdParty_CustomerID)
dfa







dfmn <- read_excel("ManualInvoice_2021.xlsx", 
                   sheet = "Sheet1")

df3 <- dfmn %>% select(ThirdParty_CustomerID, ThirdParty_SiteID, Amount)
df3$Amount <- as.numeric(df3$Amount)

df33 <- df3 %>% replace(is.na(.), 0) %>% 
  group_by(ThirdParty_CustomerID, ThirdParty_SiteID) %>% 
  summarise('Manual Handling Revenue' = sum(Amount)) %>% 
  arrange(ThirdParty_SiteID, ThirdParty_CustomerID)

df33$ThirdParty_SiteID <- str_to_title(df33$ThirdParty_SiteID)
df33


dfr <- merge(dfa, df33, all = TRUE)
dfr['Total Handling Revenue'] <- rowSums(dfr[sapply(dfr, is.numeric)], na.rm = TRUE)
dfr <- dfr %>% arrange(ThirdParty_SiteID, ThirdParty_CustomerID)
dfr








dfct <- read_excel("CustomerContainer_2021.xlsx", 
                   sheet = "Sheet3")
dfct$Facility <- str_to_title(dfct$Facility)

names(dfr)[1:2] <- names(dfct)[1:2]
dfr

dfct$`Container Qty` <- as.numeric(dfct$`Container Qty`)


df <- merge(dfr, dfct, all = TRUE)
df['Total Handling Revenue per Container'] = df$`Total Handling Revenue`/df$`Container Qty`
df <- df[, c(2, 1, 3:ncol(df))]
df <- df[order(df$Facility, df$Customer), ]
df


write.xlsx(df, 
           file = "C:\\Users\\mazhang\\Documents\\Jupyter Notebook\\Inbound Outbound Report\\Rate of Handling per Container\\Handling per Container Report.xlsx")
