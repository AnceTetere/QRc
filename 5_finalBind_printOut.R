#1 savieno gatavās tabulas
QRc <- rbind(gatavs_T327, gatavs_T418, gatavs_T330)
rm(gatavs_T327, gatavs_T418, gatavs_T330)

#2 izveido NKP vektoru izmantojot forNKP un aizpilda NKP
NKP <- c(forNKP : (nrow(QRc) + forNKP-1))
# aizpilda NKP aili
QRc$NKP <- NKP

#3 izveido nosaukumu tabulai un noglabā to, izprintē kā Excel manuālai iekopēšanai
if (!dir.exists(paste0(base_path, year, "/"))) {dir.create(paste0(base_path, year, "/"))}
if (!dir.exists(paste0(base_path, year, "/Q", Q))) {dir.create(paste0(base_path, year, "/Q", Q))}

wb <- loadWorkbook(paste0(base_path, year, "/Q", Q, "/QR_", year, "c", Q, "_savienošanai.xlsx"), create = TRUE)
createSheet(wb, name = paste0("QR_", year, "c", Q))
writeWorksheet(wb, data = QRc, sheet = paste0("QR_", year, "c", Q))
saveWorkbook(wb)

rm(QRc, wb, NKP, base_path, Q, year, forNKP)
detach("package:XLConnect", unload = TRUE)

#5. Fails, kurā jāiekopē šī mana daļa atrodams 
# ..QR\cX\nodevums
