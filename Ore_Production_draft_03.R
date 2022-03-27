library(purrr)
library(tidyverse)
library(readxl)
library(writexl)
library(dplyr)
library(ggplot2)
library(lubridate)
library(plotly)
library(RColorBrewer)
library(stringr)    
library(data.table)
library(xlsx)
library(reshape2)
library(lubridate)
library(scales)
library(formattable)
library(sqldf)

##### Set WD
setwd("~/current work/Ore_Production_2022_laptop")

##### File list
file.list <- list.files(getwd(),pattern='.xlsx', recursive = TRUE)

##### Read XLSX in File List

df <- lapply(file.list, function(i){
  x = read_xlsx(i, sheet = 1)
}) %>% as.data.frame()


###### Remove Empty Columns ######
emptycols <- sapply(df, function (k) all(is.na(k))) 
df <- df[!emptycols]




##### Promote First Row as Header
names(df) <- df[1,]


##### Remove NA Columns
df <- df[!is.na(names(df))]

names(df) <-make.unique(names(df))


colnames(df)[1] <- "DATE"
# colnames(df)[4] <- "SHIFT"
# colnames(df)[6] <- "HEADINGS"
# colnames(df)[7] <- "AREA"
# colnames(df)[8] <- "VEIN"
# colnames(df)[9] <- "LEVEL"
# colnames(df)[10] <- "METHODS"
# colnames(df)[12] <- "CHAPA"
# colnames(df)[13] <- "TONS_TF"
# colnames(df)[14] <- "GRADE_TF"
# colnames(df)[15] <- "SAMPLE#"
# colnames(df)[16] <- "DESTINATION"
# colnames(df)[17] <- "DATE_DELIVERED"
# colnames(df)[18] <- "STOCKPILE_ID"
# colnames(df)[19] <- "TONS_TS"
# colnames(df)[20] <- "GRADE_TS"


##### Identify Duplicate Columns
which( duplicated( names( df ) ) )



df <- df[-1,] %>% 
  transmute(
    DATE = DATE,
    SHIFT = SHIFT,
    HEADINGS = Column3,
    AREA = AREA,
    VEIN = VEIN,
    LEVEL = LEVEL,
    METHODS = METHODS,
    CHAPA = `CHAPA #`,
    TONS_TF = TONS_TF,
    GRADE_TS = Column1,
    SAMPLE = `SAMPLE #`,
    DESTINATION = DESTINATION,
    DATE_DELIVERED = `DATE DELIVERED`,
    SHIFT_DELIVERED = `SHIFT DELIVERED`,
    STOCKPILE_ID = `STOCKPILE ID`,
    TONS_TS = TONS_TS,
    GRADE_TS = GRADE_TS
  )


# ifelse(grepl("500",HEADINGS),"500","-")

for  (i in c("440","455","470", "485" ,"496", "500" ,"509", "515" ,"525", "530" , "545" , "560" , "575" , "590" , "605" , "620" , "635" , "650","661", "665","676", "680", "695","700","702" ,"710", "725" ,"741", "745", "840","900")) {
  df[grepl(i,df$HEADINGS),"LEVEL"] <- i
}


for  (j in c( "SDN", "SDN3", "SDN2","SDN2 SPLIT",  "SDN4", "SDN4 SPLIT","MST2", "MAS","MHWS" ,"MAI","MAI HWS", "MAIS", "BNZ", "BHWS" , "JES" , "SDY", "BBK", "BIBAK","SDN SPLIT","SDY SPLIT", "SDNS")) {
  df[grepl(j,df$HEADINGS),"VEIN"] <- j
}


for  (k in c("ODW", "ODE", "0DE", "BOS" ,"SCOURING", "C&f","XC", "RSE", "LH", "C&F", "RSE", "SHR", "S/S", "SS", "RAISE", "HWDW", "HWDE", "XC", "DXC","STOPE", "CD", "TURN OUT" , "X-CUT", "C/D", "WIDENING", "FWS","INCLINE", "DECLINE", "METERAGE", "CUTING", "CUTIING", "HWS", "SLOT")) {
  df[grepl(k,df$HEADINGS),"METHOD_3"] <- k
}


df <- df %>% mutate(LEVEL = as.numeric(LEVEL),
                    VEIN = ifelse(grepl("BIBAK",VEIN),"BBK",
                                  ifelse(grepl("SDY SPLIT",VEIN),"SDN SPLIT",
                                         ifelse(grepl("SD4",VEIN),"SDN4",
                                                ifelse(grepl("SDY",VEIN),"SDN",
                                                       ifelse(grepl("SDNS",VEIN),"SDN SPLIT",VEIN))))),
                    METHOD_M = ifelse(grepl("S/S",METHOD_3),"SS",
                                      ifelse(grepl("RAISE",METHOD_3),"RSE",METHOD_3))) %>% 
  filter(!grepl("COARSE|MIXED|BRN|brn|MIX|MARGINAL|SILT|RESTORATION|BOULDERS|RESTORARION", HEADINGS)) %>%
  arrange(desc(VEIN))


POS <- rbind(paste0(seq(0, 200, by=1),"N") %>% as.data.frame(),paste0(seq(0, 200, by=1),"S")%>% as.data.frame(),paste0(seq(0, 200, by=1),"E")%>% as.data.frame())
colnames(POS)[1] <- "POS"



##### Total number of NA in Column ######
sapply(df, function(x) sum(is.na(x)))

length(df)


# ##### SQLDF attempt - like a vlookup but sql language ######
# df <- sqldf("select df.*, POS.POS from df left join POS on instr(df.HEADINGS,  POS.POS)")


#pawer



