library(readxl)
getwd()
setwd(file.path('C:', 'Users', 'Thanakrit.B', 'Documents', 'R Projects', 'Sales similarity'))

# load data , filter only tl with stats = 'N'
df <- read_excel("~/R Projects/Sales similarity/outsource_data.xlsx")
tl <- df[substr(df$Agent_Code, 1, 1) == '4' & df$Status == 'N',]
tl <- na.omit(tl)

# tl data create age from tl.birthday

tl$age <- 2018 - as.integer(format(tl$Birthday, "%Y"))

# load zip lat long
latlong <- read_excel("zip_latlong.xlsx", col_types = c("numeric", "text", "text", 
                                                        "text", "numeric", "numeric"))
# avg lat long for duplicated zip 

library(dplyr)
latlong %>% 
  select(zip, lat, lng) %>% 
  group_by(zip) %>%
  summarise(a_lat = mean(lat), a_lng = mean(lng)) -> a_latlong

# merge tl with zip latlong
tl_lat <- merge(tl, a_latlong, by.x = 'Zip_Code', by.y = 'zip', all.x = TRUE)

# select df for distance data calculation

tl_lat %>% 
  select(Agent_Code, age, a_lat, a_lng) -> df

# prompt for target data
t_age <- readline(prompt = "Target age : ")
t_zip <- readline(prompt = "Target zipcoe : ")

y = data.frame(age = as.integer(t_age), zip = as.character(t_zip))
y_latlng <- merge(y, a_latlong)
y_latlng$zip <- NULL

# Calculate proximity
library(proxy)
p_prox <- dist(x = df[,-c(1)], y = y_latlng)

# change matrix to df
p_prox <- as.data.frame.matrix(p_prox)
top10 <- head(order(p_prox$V1))
tl_lat[top10, ]

# load amsup data
library(RODBC)
file <- file.path('C:', 'Users', 'Thanakrit.B', 'Downloads', 'OSS DataSales 201803.mdb')
con <- odbcConnectAccess(file)
sqlTables(con)
data_all_type <- sqlFetch(con, "DataAllTypeTeam", as.is = T)
colnames(data_all_type) <- make.names(colnames(data_all_type))

# summary matched amSup & calculate top tl
data_all_type %>% 
  select(AMSup.NAME, Agent_Code) %>%
  right_join(tl_lat[top10, ]) %>% 
  group_by(AMSup.NAME) %>% 
  summarise(n = n()) %>% 
  arrange(desc(n))
