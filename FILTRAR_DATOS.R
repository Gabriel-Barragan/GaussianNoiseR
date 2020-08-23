library(openxlsx)
library(dplyr)
library(mvtnorm)
FLUJOS_OTAVALO <- read.xlsx("C:/Users/Personal/Documents/PROYECTO OTAVALO/FLUJOS_OTAVALO.xlsx", colNames=T)
names(FLUJOS_OTAVALO)
FLUJOS_OTAVALO$SECTOR <- as.factor(FLUJOS_OTAVALO$SECTOR)
FLUJOS_OTAVALO$PRIORIDAD <- as.factor(FLUJOS_OTAVALO$PRIORIDAD)
FLUJOS_OTAVALO$TIPO_DE_CALLE_RELATIVO <- as.factor(FLUJOS_OTAVALO$TIPO_DE_CALLE_RELATIVO)
var1 <- FLUJOS_OTAVALO %>%
  select(NOMBRE_DE_CALLES, PRIORIDAD, FLUJO_1) %>%
  filter(!is.na(FLUJO_1));
q = rep(0, length(var1$PRIORIDAD))
scenario <- 10
for (p in 1:length(var1$PRIORIDAD)){
  if (var1$PRIORIDAD[p]=="1"){
    q[p] <- (0.40*var1$FLUJO_1[p])
  } else if(var1$PRIORIDAD[p]=="2"){
    q[p] <- (0.25*var1$FLUJO_1[p])
  } else if(var1$PRIORIDAD[p]=="3"){
    q[p] <- (0.10*var1$FLUJO_1[p])
  }
}
Q <- diag(q)

var_f1 <- round(t(replicate(scenario, var1$FLUJO_1)) + rmvnorm(n=scenario, sigma = Q))

data1 <- rbind(t(var1$NOMBRE_DE_CALLES),t(var1$FLUJO_1),var_f1)
rownames(data1) = c()
wb <- createWorkbook()
addWorksheet(wb, "Ruido1")
writeData(wb, "Ruido1", data1, startCol = 1,
          startRow = 1, rowNames = TRUE)
saveWorkbook(wb, "RUIDO1.xlsx")
#######
var2 <- FLUJOS_OTAVALO %>%
  select(NOMBRE_DE_CALLES, PRIORIDAD, FLUJO_2) %>%
  filter(!is.na(FLUJO_2));
q = rep(0, length(var2$PRIORIDAD))

for (p in 1:length(var2$PRIORIDAD)){
  if (var2$PRIORIDAD[p]=="1"){
    q[p] <- (0.40*var2$FLUJO_2[p])
  } else if(var2$PRIORIDAD[p]=="2"){
    q[p] <- (0.25*var2$FLUJO_2[p])
  } else if(var2$PRIORIDAD[p]=="3"){
    q[p] <- (0.10*var2$FLUJO_2[p])
  }
}
Q <- diag(q)

var_f2 <- round(t(replicate(scenario, var2$FLUJO_2)) + rmvnorm(n=scenario, sigma = Q))
data2 <- rbind(t(var2$NOMBRE_DE_CALLES),t(var2$FLUJO_2),var_f2)
rownames(data2) = c()
wb <- createWorkbook()
addWorksheet(wb, "Ruido2")
writeData(wb, "Ruido2", data2, startCol = 1,
          startRow = 1, rowNames = TRUE)
saveWorkbook(wb, "RUIDO2.xlsx")
#######
var3 <- FLUJOS_OTAVALO %>%
  select(NOMBRE_DE_CALLES, PRIORIDAD, FLUJO_3) %>%
  filter(!is.na(FLUJO_3));
q = rep(0, length(var3$PRIORIDAD))

for (p in 1:length(var3$PRIORIDAD)){
  if (var3$PRIORIDAD[p]=="1"){
    q[p] <- (0.40*var3$FLUJO_3[p])
  } else if(var3$PRIORIDAD[p]=="2"){
    q[p] <- (0.25*var3$FLUJO_3[p])
  } else if(var3$PRIORIDAD[p]=="3"){
    q[p] <- (0.10*var3$FLUJO_3[p])
  }
}
Q <- diag(q)

var_f3 <- round(t(replicate(scenario, var3$FLUJO_3)) + rmvnorm(n=scenario, sigma = Q))
data3 <- rbind(t(var3$NOMBRE_DE_CALLES),t(var3$FLUJO_3),var_f3)
rownames(data3) = c()
wb <- createWorkbook()
addWorksheet(wb, "Ruido3")
writeData(wb, "Ruido3", data3, startCol = 1,
          startRow = 1, rowNames = TRUE)
saveWorkbook(wb, "RUIDO3.xlsx")
