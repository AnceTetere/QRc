year <- 2023
Q <- 2
path <- "F://DARBA_PLANS//Kvalitātes_ziņojums//ceturkšņa//"

#1. savieno gatavās tabulas
setwd(paste0(path, year, "\\Q", Q, "\\intermediate_tables"))
load("gatavs_T327.RData")
load("gatavs_T418.RData")
load("gatavs_T330.RData")

QRc <- rbind(gatavs_T327, gatavs_T418, gatavs_T330)
rm(gatavs_T327, gatavs_T418, gatavs_T330)

#2. Sakārto rindas un ailes order, ja vajag
# par cik sākuma tas čuš, nekārtoju - arī nevajadzēja

#3. izveido NPK vektoru izmantojot forNPK.RDS un aizpildi nmk aili - izdzēs gan vektoru, gan failu
setwd(paste0(path, year, "\\templates"))
NPK <- readRDS("forNMK.rds")
NPK <- c(NPK : (nrow(QRc) + 4))

# aizpildi NPK aili
QRc$NPK <- NPK

#4. izveido nosaukumu tabulai un noglabā to, izprintē kā RData un kā Excel manuālai iekopēšanai
setwd(paste0(path, year, "\\Q", Q, "\\intermediate_tables"))
save(QRc, file = "final_myPartQRc.RData")

setwd(paste0(path, year, "\\Q", Q, "\\manual_bind"))
library(XLConnect)
wb <- loadWorkbook(paste0("QR_", year, "c", Q, "_savienošanai.xlsx"), create = TRUE)
createSheet(wb, name = paste0("QR_", year, "c", Q))
writeWorksheet(wb, data = QRc, sheet = paste0("QR_", year, "c", Q))
saveWorkbook(wb)

rm(QRc, wb, NPK, path, Q, year)

#5. Fails, kurā jāiekopē šī mana daļa atrodams 
# T:\SMKD\MND\ApsUzn\2_darbs\!Apmaina\202G\5_kvalitate\cX\nodevums
