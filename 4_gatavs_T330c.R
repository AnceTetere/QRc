options(digits = 13)

# Dati ir t330(c10DI)
wb <- loadWorkbook(paste0("../c10DI/", year, "Q", Q, "/SMUD/LCI_", substr(year, 3,4), "C", Q, ".xlsx"), 
                   create = FALSE, password = NULL)
ws <- readWorksheet(wb, sheet = 1)
detach("package:XLConnect", unload = TRUE)
rm(wb)

ws <- ws[ ,c("VSN", "DAT", "nT", "NOZ2", "G1", "G2", "G3", "G4", "G5", "G6", "G7", "G8")]
if (any(ws$DAT[1] == c(paste0("0", Q * 3, substr(year, 3, 4)), paste0(Q * 3, substr(year, 3, 4))))) {
  T330 <- ws[ws$nT == "330", c("NOZ2", paste0("G", Q))]
  T330$NOZ2[T330$NOZ2 == "TOTAL"] <- "Kopā (LV)"
  rownames(T330) <- NULL
} else {
  stop("4_gatavs_T330: SMUD datnē neatbilst datums.")
}

#5 Ievieto datus
mergedDF <- merge(t330, T330, by.x = "N", by.y = "NOZ2", all.x = TRUE)
mergedDF$Value <- mergedDF[ , paste0("G", Q)]

Norder <- readRDS("../QR/ceturkšņa/r_code/templates/Norder_forQRt330.rds")
mergedDF <- mergedDF[match(Norder, mergedDF$N) , ailes_order]
rownames(mergedDF) <- NULL

#7 Pārsauc
gatavs_T330 <- mergedDF

rm(t330, T330, Norder, mergedDF, ailes_order, ws)
