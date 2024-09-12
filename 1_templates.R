# Fails atrodas ..kvalitate\cX\nodevums
wb <- loadWorkbook(paste0("..kvalitate/c", Q, "/nodevums/spD_", year, ".xlsx"))
ws <- readWorksheet(wb, sheet = "Novērtējumi")
rm(wb)
detach("package:XLConnect", unload = TRUE)

#Nomaina aiļu nosaukumus
ws <- ws[-1 , c(1:11)]
rownames(ws) <- NULL
ailes <- colnames(ws)

for (i in 1:length(ailes)) {
  colnames(ws)[i] <- switch(
    ailes[i],
    "Mainīgā.npk"  = "NPK",
    "Rādītājs" = "Rādītājs",
    "Periods" = "Periods",
    "R.k" = "R_k",
    "Detalizācija..detalizāciju.kombinācija.." = "N",
    "Mērvienība" = "Mērvienība",
    "Radītāja.aprēķinu.metodoloģija..formula."  = "Formula",
    "Rādītāja..novērtējums..skaitītājs..attiecību.tipa.rādītājiem" = "Numerator",
    "Rādītāja.novērtējums..saucējs..attiecību.tipa.rādītājiem"  = "Denominator",
    "X.Rādītāja..novērtējums..eksperta." = "Evaluation",
    "Rādītāja..novērtējums..apsekojuma." = "Value"
  )
}
rm(i, ailes)

# Izņem I sadaļu
startRows <- nrow(ws)

for (q in Q:1) {
  n <- paste0("A_Bdsaf ", q, ". ceturksnī")
  ws <- ws[ws$Rādītājs != n, ]
}

# Izņem S daļu
n <- paste0("Ak_BDV_sk ", Q, ". ceturksnī")
ws <- ws[ws$Rādītājs != n, ]

# Te vajag piemērot numuru pēc kārtas (NPK)   
forNPK <- (startRows - nrow(ws)) + 1

#mape <- paste0("..\\ceturkšņa\\", year, "\\starptabulas")
#if(!dir.exists(mape)) {dir.create(mape)}
#setwd(mape)
#saveRDS(forNPK, file = "forNPK.rds")
rm(q, startRows, n)

#1. Izveido apakštabulu "Ak_ID Q. ceturksnī"
ds <- ws[1, ] 

ds$NPK[1] <- ""
ds$Rādītājs[1] <- paste0("Ak_ID ", Q, ". ceturksnī")
ds$Periods[1] <- paste0(year, "C0", Q, " - ceturkšņa")
ds$R_k[1] <- "(rn = 1470 + 310 + 320 + 330 + 340 + 350 + 360 - 1499)"
ds$N[1] <- "Kopā (LV)"
ds$Mērvienība[1] <- "EUR"
ds$Formula[1] <- "sum((\"rn=1470_1 + 310_1 + 320_1 + 330_1 + 340_1 + 350_1 + 360_1 -1499_1\" & \"12-zīmju UUK 9.-12.zīme <>\"8100\" \")*[koef])"
ds$Numerator[1] <- ""
ds$Denominator[1] <- ""
ds$Evaluation[1] <- ""
ds$Value[1] <- ""

rownames(ds) <- NULL
t327 <- ds
rm(ds)

#setwd(mape)
#save(T327,  file = paste0(year, "Q", Q, "_T327.RData"))
#rm(T327)

#2. Izveido apakštabulu "Ns q.ceturksnī" q rindas - kopa LV tikai
ds <- ws[1, ] 

  ds$NPK[1] <- ""
  ds$Rādītājs[1] <- paste0("Ns ", Q, ". ceturksnī")
  ds$Periods[1] <- paste0(year, "C0", Q, " - ceturkšņa")
  ds$R_k[1] <- "(rn=1310; 1320; 1330)"
  ds$N[1] <- "Kopā (LV)"
  ds$Mērvienība[1] <- "stundas"
  ds$Formula[1] <- "sum([p1(\"rn=1310_1+1320_1+1330_1\" & \"12-zīmju UUK 9.-12.zīme <>\"8100\" \")]*[koef])"
  ds$Numerator[1] <- ""
  ds$Denominator[1] <- ""
  ds$Evaluation[1] <- ""
  ds$Value[1] <- ""

rownames(ds) <- NULL
t418 <- ds
rm(ds)

#save(T418,  file = paste0(year, "Q", Q, "_T418.RData"))
#rm(T418)

#3. Izveido apakštabulu "kDI_nS q. ceturksnī" - nace

#N_vec <- c("Kopā (LV)", "A", "A01", "A02", "A03", "B", "B06", "B08", "B09", "C", "C10", "C11", 
#              "C12", "C13", "C14", "C15", "C16", "C17", "C18", "C19", "C20", "C21", "C22", "C23", 
#              "C24", "C25", "C26", "C27", "C28", "C29", "C30", "C31", "C32", "C33", "D", "D35", "E",  
#              "E36", "E37", "E38", "E39", "F", "F41", "F42", "F43", "G", "G45", "G46", "G47", "H",
#              "H49", "H50", "H51", "H52", "H53", "I", "I55", "I56", "J", "J58", "J59", "J60", "J61",
#              "J62", "J63", "K", "K64", "K65", "K66", "L", "L68", "M", "M69", "M70", "M71", "M72",
#              "M73", "M74", "M75", "N", "N77", "N78", "N79", "N80", "N81", "N82", "O", "O84", "P",
#              "P85", "Q", "Q86", "Q87", "Q88", "R", "R90", "R91", "R92", "R93", "S", "S94", "S95",
#             "S96")
#saveRDS(N_vec, file = "Single_N_vec.rds")
N_vec <- readRDS(paste0("../templates/Single_N_vec.rds"))

ds <- ws[1:length(N_vec),]

  for (i in 1:length(N_vec)) {
    ds$NPK[i] <- ""
    ds$Rādītājs[i] <-
      paste0("kDI_nS ",
             Q,
             ". ceturksnī")
    ds$Periods[i] <- paste0(year, "C0", Q, " - ceturkšņa")
    ds$R_k[i] <-
      "(1470.r.1.a. + 310.r.1.a. + 320.r.1.a. + 330.r.1.a. + 340.r.1.a. + 350.r.1.a. + 360.r.1.a. - 1499.r.1.a.) : (1310.r.1.a. + 1320.r.1.a. + 1330.r.1.a.)"
    ds$N[i] <- N_vec[i]
    ds$Mērvienība[i] <- "EUR"
    ds$Formula[i] <-
      "sum((1470.r.1.a. + 310.r.1.a. + 320.r.1.a. + 330.r.1.a. + 340.r.1.a. + 350.r.1.a. + 360.r.1.a. - 1499.r.1.a.) : (1310.r.1.a. + 1320.r.1.a. + 1330.r.1.a.) <>\"8100\" \")*[koef])"
    ds$Numerator[i] <- ""
    ds$Denominator[i] <- ""
    ds$Evaluation[i] <- ""
    ds$Value[i] <- ""
  }

rm(i, ws, N_vec)

rownames(ds) <- NULL
t330 <- ds
rm(ds)

#save(T330,  file = paste0(year, "Q", Q, "_T330.RData"))
#saveRDS(ailes_order, file = "ailes_order.RDS")
#saveRDS(forNPK, file = "forNPK.RDS")
#rm(T330, ailes_order, forNPK)
