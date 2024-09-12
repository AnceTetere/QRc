#1 Ak_ID nes datus no T327
t327$Value[t327$Periods == paste0(year, "C0", Q, " - ceturkšņa")] <- T327[T327$NOZ2 == "TOTAL", paste0("G", Q)]
gatavs_t327 <- t327
rm(T327, t327)
