library(tidyverse)
library(readxl)
library(writexl)

df <- read_excel("BillingRecord_03_2022.xlsx")

df['Creator'] <- ifelse(df$`Created By` == 'system', 'SYSTEM', 'USER')
unique(df$Creator)
df['Mode'] <- NA

df0 <- df
df1 <- df0 %>% distinct(DocNo, Creator, .keep_all = TRUE)
df1

dup <- df1[duplicated(df1$DocNo), ]
dup
no_dup <- df1[!duplicated(df1$DocNo), ]
no_dup

overlap_dup <- no_dup$DocNo %in% dup$DocNo

no_dup[!overlap_dup,]['Mode'] <- ifelse(no_dup[!overlap_dup,]$Creator == 'USER', 'MANUAL', 'AUTO')

no_dup[overlap_dup,]['Mode'] <- c('SEMI-AUTO')

no_dup

colnames(df1)

# df2 <- full_join(df1, no_dup, by = 'DocNo')
# df2[, c(which(colnames(df2) == 'DocNo'), ncol(df2) - 1, ncol(df2))]

df3 <- full_join(df0, no_dup, by = 'DocNo')

df4 <- df3[, c(1:(ncol(df0) - 1), ncol(df3))]

colnames(df4) <- colnames(df0)

table(df4$Mode)

write_xlsx(df4, "C:\\Users\\mazhang\\Documents\\Jupyter Notebook\\Billing Record Report\\2022 Billing Record Report\\03_2022 Billing Record Report.xlsx")

