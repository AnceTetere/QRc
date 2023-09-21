#1. Nākamais atrodi datus
setwd(paste0(path, "//", year, "//data_tables"))

for (q in Q:1) {
  fileName <- paste0("T418_", substr(year, 3, 4), "Q", q)
  load(paste0(fileName, ".RData"))
  assign(paste0("t", q), get(fileName))
  rm(list = fileName, fileName)
}
rm(q)

#2. Ielādē šablonu
setwd(paste0(path, year, "//templates"))
load(paste0(year, "Q", Q, "_T418.RData"))


#3. Izformē datu tabulas

for (q in Q:1) {
  x <- get(paste0("t", q))
  T418$Value[T418$Periods == paste0(year, "C0", q, "99")] <- x[x$NOZARE2 == "TOTAL", paste0("G", q)]
  rm(list = paste0("t", q), x)
}

rm(q)

# sakārto, pārsauc un noglabā kā gatavs_T418
gatavs_T418 <- T418

setwd(paste0(path, year, "//intermediate_tables"))
save(gatavs_T418, file = "gatavs_T418.RData")
