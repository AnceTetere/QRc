  t418$Value[t418$Periods == paste0(year, "C0", Q, " - ceturkšņa")] <- T418[T418$NOZ2 == "TOTAL", paste0("G", Q)]
  gatavs_T418 <- t418
  rm(T418, t418)
