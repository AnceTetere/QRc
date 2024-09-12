# Paņem datus no sagatavēm IID ESTAT
library("XLConnect")
year <- "2024"
Q <- 2
base_path <- paste0("../KS/ceturkšņa/") 

# ICL_ESTAT
dName <- paste0("..ESTAT/", year, "/ICL", year, "_Q", Q, "/pasutijums/ICL_ESTAT_", substr(year, 3, 4), "c", Q)
wb <- loadWorkbook(paste0(dName, ".xlsx"), create = FALSE, password = NULL)
ws <- readWorksheet(wb, sheet = 1)
rm(wb)
detach("package:XLConnect", unload = TRUE)

ICL_dati <- ws[ ,c("DAT", "VSN", "nT", "NOZ2", "G1", "G2", "G3", "G4", "G5", "G6", "G7", "G8")]
rm(ws)

# T327
#tName <- paste0("T327_", substr(year, 3, 4), "Q", Q)
T327 <- ICL_dati[ICL_dati$nT == "327", ]
#assign(tName, T327)
#rm(tName)

#mape <- paste0(paste0(path, "data_tables"))
#if (!file.exists(mape)) {dir.create(mape)}
#setwd(mape)
#save(list = tName, file = paste0(tName, ".RData"))

#rm(list = tName)

# T418
#tName <- paste0("T418_", substr(year, 3, 4), "Q", Q)
T418 <- ICL_dati[ICL_dati$nT == "418", ]
#assign(tName, T418)

#mape <- paste0(paste0(path, "//data_tables"))
#if (!file.exists(mape)) {dir.create(mape)}
#setwd(mape)
#save(list = tName, file = paste0(tName, ".RData"))

#rm(list = tName, tName, mape)

#------------- Vai ...
# T327
#setwd(paste0("..ESTAT\\", year, "\\ICL", year, "_Q", Q, "\\workflow\\dati\\noSMUD"))
#tName <- paste0("T327_", substr(year, 3, 4), "Q", Q)
#load(paste0(tName, ".RData"))

#setwd(paste0(path, "//data_tables"))
#save(list = tName, file = paste0(tName, ".RData"))

# T418
#tName <- paste0("T418_", substr(year, 3, 4), "Q", Q)
#load(paste0(tName, ".RData"))

#setwd(paste0(path, "//data_tables"))
#save(list = tName, file = paste0(tName, ".RData"))

#rm(tName, ICL_dati)
rm(ICL_dati)
