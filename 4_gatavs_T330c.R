options(digits = 13)

# Dati sadaļai 
# 1. Paņem datus
setwd(paste0(path, "//intermediate_tables"))
load(paste0("1_LC_TOTAL_DIS010c_", substr(year, 3, 4), "Q", Q, ".RData"))      
t <- get(paste0("LC_TOTAL_DIS010c_", substr(year, 3, 4), "Q", Q))
#fileName <- paste0("t330_", Q)
#assign(fileName, t)
rm(list = paste0("LC_TOTAL_DIS010c_", substr(year, 3, 4), "Q", Q))

#2. Paņem ceturkšņa QR šablonu T330
setwd(paste0(path, year, "//templates"))
load(paste0(year, "Q", Q, "_T330.RData")) 

#3. šablonu sadali pa ceturkšņiem
T330$Periods <- factor(T330$Periods)
t_split <- split(T330, T330$Periods)
names(t_split) <- paste0("T330_", substr(names(t_split), 1, 7))
t_split_tabs <- names(t_split)
list2env(t_split, envir = .GlobalEnv)
rm(t_split, T330)

# paņem vajadzīgo
T330 <- get(paste0("T330_", year, "C0", Q))
rm(list = t_split_tabs)
rm(t_split_tabs)

#4. Datu šablonā pārvārdo "A-S" uz "Kopā (LV)" un izņem "B-S"
t$NACE[t$NACE == "A-S"] <- "Kopā"
t <- t[t$NACE != "B-S", ]
sum(t$NACE %in% T330$NACE) == nrow(t)

#5. Ievieto datus
mergedDF <- merge(T330, t[ , c("NACE", "EUR")], by.x = "NACE", by.y = "NACE")
mergedDF$Value <- mergedDF$EUR

#6. Sakārto ailes un rindas
setwd(paste0(path, year, "//templates"))
ailes <- readRDS("ailes_order.RDS")
mergedDF<- mergedDF[order(mergedDF$NACE), ailes]
rm(t, T330, ailes)

#7. Pārsauc un saglabā
gatavs_T330 <- mergedDF

setwd(paste0(path, year, "//intermediate_tables"))
save(gatavs_T330, file = "gatavs_T330.RData")
rm(mergedDF, gatavs_T330)
