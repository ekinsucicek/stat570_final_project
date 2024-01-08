######################### READING - CLEANING ###################################
## Downlading and Reading the Data
library(readxl)
library(dplyr)
zipF <- "https://prod-dcd-datasets-cache-zipfiles.s3.eu-west-1.amazonaws.com/d2s9yz65mm-4.zip"
local_zip_filename <- "Dataset of Women Directors Engagement.zip"
download.file(zipF, local_zip_filename, mode = "wb", method = "auto")
unzip(local_zip_filename)
excel_files <- c("CID Scores (Table A).xlsx", "WDs Engagement (Table B).xlsx", "EWDs Aggregated Score (Table C).xlsx")

## Reading Each Excel Files
### CID SCORES
CID_2018 <- read_excel("CID Scores (Table A).xlsx", sheet = "FYE 2018")
CID_2019 <- read_excel("CID Scores (Table A).xlsx", sheet = "FYE 2019")
CID_2020 <- read_excel("CID Scores (Table A).xlsx", sheet = "FYE 2020")

CID_2018 <- CID_2018[-c(1,2),]
CID_2019 <- CID_2019[-c(1,2),]
CID_2020 <- CID_2020[-c(1,2),]

#set column common colnames
set_column_names <- function(data, year) {
  colnames(data) <- c("company_name", "strategy_policy", "climate_change_opportunities",
                      "corporate_ghg_emissions_targets", "company_wide_carbon_footprint",
                      "ghg_emissions_change_over_time", "energy_related_reporting",
                      "emission_reduction_initiatives_implementation",
                      "carbon_emission_accountability", "quality_of_disclosure",
                      paste0("quality_of_DisclosureTotal_cid_scores"),"year")
  
  data$year <- year
  
  return(data)
}

CID_2018 <- set_column_names(CID_2018,"2018")
CID_2019 <- set_column_names(CID_2019,"2019")
CID_2020 <- set_column_names(CID_2020,"2020")

CID_scores <- rbind(CID_2018, CID_2019,CID_2020)


CID_scores <- CID_scores %>%
  mutate_at(vars(-company_name, -year), as.numeric)

### WDS ENGAGEMENT
WDs_2018 <- read_xlsx("WDs Engagement (Table B).xlsx", sheet = "FYE 2018")
WDs_2019 <- read_xlsx("WDs Engagement (Table B).xlsx", sheet = "FYE 2019")
WDs_2020 <- read_xlsx("WDs Engagement (Table B).xlsx", sheet = "FYE 2020")

WDs_2018 <- WDs_2018[-1,]
WDs_2019 <- WDs_2019[-1,]
WDs_2020 <- WDs_2020[-1,]

set_column_names_2 <- function(data, year) {
  colnames(data) <- c("company_name", "number_of_wd", "per_of_wd_on_board", 
                      "per_of_wd_industry_expert", 
                      "per_of_wd_advisors","per_of_wd_community_leader")
  
  data$year <- year
  
  return(data)
}

WDs_2018 <- set_column_names_2(WDs_2018,"2018")
WDs_2019 <- set_column_names_2(WDs_2019,"2019")
WDs_2020 <- set_column_names_2(WDs_2020,"2020")

WDs_engagement <- rbind(WDs_2018, WDs_2019, WDs_2020)

# Convert columns to numeric (excluding company_name)
WDs_engagement <- WDs_engagement |> 
  mutate_at(vars(-company_name,-year), function(x) round(as.numeric(sub("%", "", x)), 2))

### EWDS SCORES
EWDs_2018 <- read_xlsx("EWDs Aggregated Score (Table C).xlsx", sheet = "FYE 2018")
EWDs_2019 <- read_xlsx("EWDs Aggregated Score (Table C).xlsx", sheet = "FYE 2019")
EWDs_2020 <- read_xlsx("EWDs Aggregated Score (Table C).xlsx", sheet = "FYE 2020")
EWDs_scores <- rbind(EWDs_2018, EWDs_2019, EWDs_2020)
str(EWDs_scores)
set.column.names <- function(data,year) {
  colnames(data) <- c("company_name", "code", "WDsName", "industry_expert", 
                      "advisor", "community_leaders", "energy_experiments", "log_of_energy_experiment")
  data$year <- year
  
  return(data)
}
EWDs_2018 <- set.column.names(EWDs_2018,"2018")
EWDs_2019 <- set.column.names(EWDs_2019,"2019")
EWDs_2020 <- set.column.names(EWDs_2020,"2020")
#Replace name in empty space
for (i in 1:4) {EWDs_2018 <- EWDs_2018 |> 
  mutate('company_name' = ifelse(is.na(company_name), lag(company_name), company_name))}
for (i in 1:5) {EWDs_2019 <- EWDs_2019 |> 
  mutate('company_name' = ifelse(is.na(company_name), lag(company_name), company_name))}
for (i in 1:6) {EWDs_2020 <- EWDs_2020 |> 
  mutate('company_name' = ifelse(is.na(company_name), lag(company_name), company_name))}

EWDs_scores$energy_experiments <- as.numeric(EWDs_scores$energy_experiments)
EWDs_scores <- EWDs_scores |> 
  mutate(wd_title = cut(energy_experiments, breaks = c(-Inf, 15, 30, 49),
                        labels = c("Assistant Director", "Director","Senior Director"), include.lowest = TRUE ))

################################## ANALYSIS ###################################
### What are the companies with the highest average quality of disclosure 
### total of CID scores of three years 2018,2019,2020?
library(dplyr)
library(ggplot2)
library(plotly)

mean_cid <- CID_scores %>%
  filter(year %in% c("2018", "2019", "2020")) %>%
  group_by(company_name) %>%
  summarize(avg_cid_score = mean(quality_of_DisclosureTotal_cid_scores, na.rm = TRUE)) |> 
  slice_max(order_by=avg_cid_score,n=10)

ggplot(mean_cid, aes(x = reorder(company_name, -avg_cid_score), y = avg_cid_score)) +
  geom_bar(stat = "identity", fill = "skyblue") +
  labs(title = "Top 10 Companies by Average Score",
       x = "Company Name",
       y = "Average Score") +
  theme(axis.text.x = element_text(angle = 45, hjust = 1)) 

plot_ly(mean_cid, x = ~company_name, y = ~avg_cid_score, type = 'bar', color = ~avg_cid_score) %>%
  layout(title = "Top 10 Companies by Average Score",
         xaxis = list(title = "Company Name", categoryorder = 'total descending'),
         yaxis = list(title = "Average Score", range = c(75, 85))) %>%
  config(displayModeBar = FALSE)

### Which are the 5 companies with the highest and lowest ratio of woman 
### advisors in 2020? Where are they located on the map?
WDs_engagement %>%
  filter(year == 2020) %>%
  group_by(company_name) %>%
  summarize(Ratio = sum(per_of_wd_advisors)) %>%
  arrange(desc(Ratio)) |> 
  head(5)

WDs_engagement %>%
  filter(year == 2020) %>%
  group_by(company_name) %>%
  summarize(Ratio = sum(per_of_wd_advisors)) %>%
  arrange(desc(Ratio)) |> 
  tail(5)

mapquestion <- WDs_engagement |>
  filter(company_name %in% c("Phillips 66", "PTT Public Company Limited",
                             "Siemens Gamesa Renewable Energy", "Royal Dutch Shell",
                             "Snam","Suncor Energy","Santos", "Tenaris SA",
                             "Vestas","Ultrapar Participações SA"))

long_lat <- data.frame(company_name = c("Phillips 66", "PTT Public Company Limited",
                                        "Siemens Gamesa Renewable Energy", "Royal Dutch Shell",
                                        "Snam","Suncor Energy","Santos", "Tenaris SA",
                                        "Vestas","Ultrapar Participações SA"),
                       longitude = c(-95.3698028, 100.523186, 9.993682, -0.118092, 9.26838000, -114.066666, 138.599503, 6.131935, 10.21076, -46.625290),
                       lattitude = c(29.7604267, 13.736717, 53.551086, 51.509865, 45.41047000, 51.049999, -34.921230, 49.611622, 56.15674, -23.533773))

mapquestion2 <- merge(mapquestion,long_lat, by=c("company_name"))

center_lon <- median(mapquestion2$longitude, na.rm = TRUE)
center_lat <- median(mapquestion2$lattitude, na.rm = TRUE)
library(leaflet)
leaflet() %>%
  addProviderTiles("Esri") %>%
  addMarkers(
    data = mapquestion2,
    lng = ~longitude,
    lat = ~lattitude,
    popup = ~company_name
  ) %>%
  setView(lng = center_lon, lat = center_lat, zoom = 1)

library(sf)

greenLeafIcon <- makeIcon(
  iconUrl = "https://static.vecteezy.com/system/resources/thumbnails/024/624/044/small/african-black-woman-watercolor-ai-generative-png.png",
  iconWidth = 38, iconHeight = 95,
  iconAnchorX = 22, iconAnchorY = 94,
  shadowWidth = 50, shadowHeight = 64,
  shadowAnchorX = 4, shadowAnchorY = 62
)

leaflet() %>%
  addProviderTiles("Esri") %>%
  addMarkers(
    data = mapquestion2,
    lng = ~longitude,
    lat = ~lattitude,
    popup = ~company_name,
    icon = greenLeafIcon
  ) %>%
  setView(lng = center_lon, lat = center_lat, zoom = 1)




View(EWDs_scores)
View(WDs_engagement)
View(CID_scores)
sqldf::sqldf("SELECT wd_title,COUNT(*) FROM EWDs_scores GROUP BY wd_title")
WDs_engagement$per_of_wd_on_board
WDs_engagement$per_of_wd_industry_expert
WDs_engagement$per_of_wd_advisors
WDs_engagement$per_of_wd_community_leader

sqldf::sqldf("SELECT *
             FROM WDs_engagement
             GROUP BY per_of_wd_community_leader
             GROUP BY per_of_wd_advisors
             GROUP BY per_of_wd_industry_expert
             GROUP BY per_of_wd_on_board")

alikoc <- sqldf::sqldf("SELECT year, 
AVG(per_of_wd_community_leader) AS per_of_wd_community_leader,
AVG(per_of_wd_advisors) AS per_of_wd_advisors,
AVG(per_of_wd_industry_expert) AS per_of_wd_industry_expert,
AVG(per_of_wd_on_board) AS per_of_wd_on_board
             FROM WDs_engagement
             GROUP BY year")

fig1 <- plot_ly(x = alikoc$year, y = alikoc$per_of_wd_community_leader, type = 'scatter', mode = 'lines+markers') 
fig2 <- plot_ly(x = alikoc$year, y = alikoc$per_of_wd_advisors, type = 'scatter', mode = 'lines+markers') 
fig3 <- plot_ly(x = alikoc$year, y = alikoc$per_of_wd_industry_expert, type = 'scatter', mode = 'lines+markers') 
fig4 <- plot_ly(x = alikoc$year, y = alikoc$per_of_wd_on_board, type = 'scatter', mode = 'lines+markers') 
fig_all <- subplot(fig1, fig2, fig3, fig4, nrows = 2)


fig <-plotly::plot_ly(data = alikoc,y = ~per_of_wd_community_leader,
                      x = ~year,name = "Community Leader",
                      type = "scatter",mode = "lines") %>%
  add_trace(y = ~per_of_wd_advisors, name = "Advisors") %>% 
  add_trace(y = ~per_of_wd_industry_expert, name = "Industry Expert") %>%
  add_trace(y = ~per_of_wd_on_board, name = "Board Member")
fig



View(sqldf::sqldf("SELECT DISTINCT(company_name) FROM WDs_engagement"))


###############################################################################

library(httr)
library(rvest)

url_ulke <- "https://developers.google.com/public-data/docs/canonical/countries_csv?hl=en"
res <- GET(url_ulke)
html_con <- content(res, "text")
html_ulke <- read_html(html_con)

# Extract all tables from the webpage
tables <- html_ulke %>%
  html_nodes("table") %>%
  html_table()

long_lat_country <- tables[[1]]
long_lat_country <- long_lat_country[,-1]

colnames(long_lat_country)[3] <- "Country"


company <- read.csv('company.csv', sep = ";")
table(company$Continent)

# Extract all tables from the webpage
tables <- html_ulke %>%
  html_nodes("table") %>%
  html_table()

long_lat_country <- tables[[1]]
long_lat_country <- long_lat_country[,-1]
colnames(long_lat_country)[3] <- "Country"
toplu_son_mu <- merge(x = company, y = long_lat_country, by = "Country", all.x = TRUE) 

df <- sqldf::sqldf("SELECT Country, continent, latitude, longitude, COUNT(*) AS Freq 
             FROM toplu_son_mu
             GROUP BY Country
                   ORDER BY Freq DESC")
df_son <- (sqldf::sqldf("SELECT Freq,A.* FROM df B LEFT JOIN toplu_son_mu A ON A.latitude=B.latitude"))



center_lon <- median(df$longitude, na.rm = TRUE)
center_lat <- median(df$lattitude, na.rm = TRUE)
library(leaflet)

getColor <- function(df) {
  sapply(df$Freq, function(mag) {
    if(mag <= 2) {
      "green"
    } else if(mag <= 4) {
      "orange"
    } else {
      "red"
    } })
}

icons <- awesomeIcons(
  icon = 'ios-close',
  iconColor = 'black',
  library = 'ion',
  markerColor = getColor(df)
)

str(df)
leaflet() %>%
  addProviderTiles(providers$Stadia.StamenTonerLite,
                   options = providerTileOptions(noWrap = TRUE)) %>%
  addAwesomeMarkers(
    data = df_son,
    lng = ~longitude,
    lat = ~latitude,
    label = ~Country,
    icon = icons,
    popup = ~paste("<br>Number of Company:", Freq,
                   "<br>Company Names:", company_name),
    clusterOptions = markerClusterOptions()
  ) 

