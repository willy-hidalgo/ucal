# install.packages('rattle')
# install.packages('RColorBrewer')

library('readxl')
library('xlsx')
library('data.table')
library('rpart')
library('rpart.plot')
#library('ROCR')
library('pROC')
library('rattle')
# library('RColorBrewer')

rm(list = ls())
cat('\014')

default.path <- 'C:/Users/Willy/Google Drive/consultorías/toulouse/'

# default.path <- paste(dirname(parent.frame(2)$ofile), '/', sep = '')
# default.path <-
#   paste(dirname(rstudioapi::getActiveDocumentContext()$path),
#         '/',
#         sep = '')

# Carpeta por defecto                                                       ----
setwd(default.path)

# Evaluación de modelos
eval1 <- function(model, pred, truth) {
  print(model)
  print(table(pred,  truth))
  tab <- as.matrix(table(pred == truth,  truth))
  # print(tab)
  (tot <- colSums(tab))
  (truepos <- unname(rev(cumsum(rev(
    tab[2,]
  )))))
  (falsepos <- unname(rev(cumsum(rev(
    tab[1,]
  )))))
  (totpos <- sum(tab[2,]))
  (totneg <- sum(tab[1,]))
  (sens <- truepos / totpos)
  (omspec <- falsepos / totneg)
  sens <- c(sens, 0)
  omspec <- c(omspec, 0)
  plot(
    omspec,
    sens,
    type = "b",
    xlim = c(0, 1),
    ylim = c(0, 1),
    lwd = 2,
    xlab = "1 − specificity",
    ylab = "Sensitivity",
    main = model
  )
  grid()
  abline(0, 1, col = "red", lty = 2)
  
  print(paste('Adj: ', mean(pred == truth) * 100, sep = ''))
  print(paste('MSE: ', mean((pred - truth) ^ 2) * 100, sep = ''))
  print(paste('MAPE: ', mean(abs(pred - truth) / truth) * 100, sep = ''))
  # roc(truth, pred, direction = '<', levels = c(1, 2), col = 'red', main = model)
}

# Lee data
file.name <- 'clientes endtoend(9275).xlsx'
num.cols <- 54
alumnos <-
  as.data.table(read_excel(file.name, col_types = rep('text', num.cols)))

# Normaliza nombre de columnas
chng.cols <- names(alumnos)
chng.cols <- gsub(" ", ".", chng.cols)
chng.cols <- tolower(gsub("_", ".", chng.cols))
names(alumnos) <- chng.cols

# Selecciona columnas a utilizar
alumnos <- alumnos[, c(
  'sexo',
  'distrito',
  'medio.informa',
  'programa',
  'curso',
  'edad',
  'profesion',
  'est.civil',
  'sede',
  'programa1',
  'tipo.cliente',
  'descuento',
  'paga.mat.',
  'sede.interes'
)]

# Edad
alumnos$edad <- as.integer(alumnos$edad)
alumnos <- alumnos[edad > 0]
alumnos$r.edad <- as.factor('20 - 30')
alumnos[edad < 10]$r.edad <- '< 10'
alumnos[edad >= 10 & edad < 20]$r.edad <- '10 - 20'
alumnos[edad >= 20 & edad < 30]$r.edad <- '20 - 30'
alumnos[edad >= 30 & edad < 40]$r.edad <- '30 - 40'
alumnos[edad >= 40 & edad < 50]$r.edad <- '40 - 50'
alumnos[edad >= 50]$r.edad <- '50+'

# Los cinco cursos con más demanda
alumnos[is.na(curso)]$curso <-
  alumnos[is.na(curso)]$programa
alumnos[curso == 'Taller Adultos']$curso <-
  alumnos[curso == 'Taller Adultos']$programa
alumnos[is.na(curso)]$curso <- 'No aplicable'
t <- as.data.table(table(alumnos$curso))
setorder(t, -N)
t[, CumSum := cumsum(N)]
cursos <- t$V1[1:5]
alumnos <- alumnos[curso %in% cursos]
rm(t)

# # los cursos 2 y 3 son iguales
# alumnos[curso %in% c('Gráfica Digital (Regular)', 'Gráfica Digital (Intensivo)')]$curso <-
#   'Gráfica digital'

# Distritos
d <- as.data.table(table(alumnos$distrito))
setorder(d, -N)
d[, CumSum := cumsum(N)]
distritos <- d$V1[1:4]
alumnos[!(distrito %in% distritos)]$distrito <- 'Otro distrito'
rm(d)

# Medio información
m <- as.data.table(table(alumnos$medio.informa))
setorder(m, -N)
m[, CumSum := cumsum(N)]
medios <- m$V1[1:4]
alumnos[!(medio.informa %in% medios)]$medio.informa <- 'Otro medio'
rm(m)

# Profesión
p <- as.data.table(table(alumnos$profesion))
setorder(p, -N)
p[, CumSum := cumsum(N)]
profesiones <- p$V1[1:3]
alumnos <- alumnos[profesion %in% profesiones]
rm(p)

# Sexo
alumnos[is.na(sexo)]$sexo <- 'No indica'

# Estado civil
alumnos[is.na(est.civil)]$est.civil <- 'Otro'
alumnos$est.civil <- as.factor(alumnos$est.civil)

# Tipo cliente
alumnos[is.na(tipo.cliente)]$tipo.cliente <- 'Otro'

# Paga matrícula
alumnos[is.na(paga.mat.)]$paga.mat. <- 'Otro'

# Programa1
alumnos[is.na(programa1)]$programa1 <- 'Otro'

# NA omit
alumnos <- alumnos[, c('curso', 'sexo', 'distrito', 'medio.informa',
  'r.edad', 'profesion', 'est.civil')]
alumnos <- na.omit(alumnos)

alumnos$curso <- as.factor(alumnos$curso)
alumnos$sexo <- as.factor(alumnos$sexo)
alumnos$paga.mat. <- as.factor(alumnos$paga.mat.)
alumnos$programa1 <- as.factor(alumnos$programa1)
alumnos$tipo.cliente <- as.factor(alumnos$tipo.cliente)
alumnos$profesion <- as.factor(alumnos$profesion)
alumnos$medio.informa <- as.factor(alumnos$medio.informa)
alumnos$distrito <- as.factor(alumnos$distrito)


# chng.cols <- names(Filter(is.character, alumnos))
# id.chng.cols <- paste('id', chng.cols, sep = '.')
# alumnos[, (id.chng.cols) := lapply(.SD, as.factor), .SDcols = chng.cols]
# alumnos[, (id.chng.cols) := lapply(.SD, as.numeric), .SDcols = id.chng.cols]
# 
# a <- alumnos[, ..id.chng.cols]

train.vector <- sample(c(TRUE, FALSE), nrow(alumnos), replace = TRUE)
test.vector <- (!train.vector)

train <- alumnos[train.vector]
test <- alumnos[test.vector]

# train$y <- as.numeric(train$id.curso/max(train$id.curso))
# model.glm <-
#   glm(y ~ id.sexo + id.distrito + id.medio.informa + id.periodo + id.vendedor +
#         id.edad + id.profesion + id.est.civil  + id.campania + id.actividad +
#         id.sede.interes,
#       data = train)
#
# summary(model.glm)

# chng.cols <- names(Filter(is.character, alumnos))
# alumnos[, (chng.cols) := lapply(.SD, as.factor), .SDcols = chng.cols]
# train <- alumnos[1:round(nrow(alumnos)/2, 0)]
# test <- alumnos[(round(nrow(alumnos)/2, 0) + 1):nrow(alumnos)]

# Árbol de decisión
rp <-
  rpart(
    curso ~ sexo + distrito + medio.informa +
      r.edad + profesion + est.civil,
    data = train
  )

rpart.plot(rp)

par(mfrow = c(1, 3))
truestat <- as.integer(test$curso)
testres <- (as.integer(predict(rp, test, type = 'class')) == truestat)
print('RPart')
print(as.matrix(table(testres,  truestat)))
print(roc(testres, truestat)$auc)
plot(roc(testres, truestat), col = 'red', main = 'RPart')

# SVM
library(e1071)
tune.out <-
  tune(
    svm,
    as.integer(curso) ~ sexo + distrito + medio.informa +
      r.edad + profesion + est.civil,
    data = train,
    kernel = 'radial',
    ranges = list(
      cost = c(0.1, 1, 10, 100, 1000),
      gamma = c(0.5, 1, 2, 3, 4)
    )
  )

truestat <- as.integer(test$curso)
testres <- (as.integer(predict(tune.out$best.model, test)) == truestat)
print('SVM')
print(as.matrix(table(testres,  truestat)))
print(roc(testres, truestat)$auc)
plot(roc(testres, truestat), col = 'red', main = 'RPart')

# Random forest
library(randomForest)
rf <-
  randomForest(
    curso ~ sexo + distrito + medio.informa +
      r.edad + profesion + est.civil,
    data = train,
    method = 'class',
    importance = TRUE
  )

# plot(rf)
truestat <- as.integer(test$curso)
testres <- (as.integer(predict(rf, test, type = 'class')) == truestat)
print('Random forest')
print(as.matrix(table(testres,  truestat)))
print(roc(testres, truestat)$auc)
plot(roc(testres, truestat), col = 'red', main = 'RandomForest')

# install.packages('party')
library(party)
fit <-
  cforest(
    curso ~ sexo + distrito + medio.informa +
      r.edad + profesion + est.civil,
    data = train
  )
truestat <- as.integer(test$curso)
testres <- (as.integer(predict(fit, test, OOB = TRUE, type = 'response')) == truestat)

print('Party')
print(as.matrix(table(testres,  truestat)))
print(roc(testres, truestat)$auc)
plot(roc(testres, truestat), col = 'red', main = 'Party')
