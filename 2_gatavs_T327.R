year <- ..
Q <- ... #TODO izstrādā ievadi
path <- "..."

#1. Tabula 3.27 

setwd(paste0(path, "//", year, "//data_tables"))

for (q in Q:1) {
  fileName <- paste0("T327_", substr(year, 3, 4), "Q", q)
  load(paste0(fileName, ".RData"))
  assign(paste0("t", q), get(fileName))
  rm(list = fileName, fileName)
}
rm(q)

#2. Ielādē šablonu
setwd(paste0("path, "//", year, "//templates"))
readRDS()
load(paste0(year, "Q", Q, "_T327.RData"))


#3. Izformē datu tabulas,

for (q in Q:1) {
x <- get(paste0("t", q))
T327$Value[T327$Periods == paste0(year, "C0", q, " -")] <- x[x$NOZARE2 == "TOTAL", paste0("G", q)]
rm(list = paste0("t", q), x)
}

rm(q)

# sakārto, pārsauc un noglabā (izmanto ailes_order saglabāto vektoru)
gatavs_T327 <- T327

setwd(paste0(path, "//", year, "//intermediate_tables"))
save(gatavs_T327, file = "gatavs_T327.RData")

rm(gatavs_T327, T327)
