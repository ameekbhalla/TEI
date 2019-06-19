library(tidyr)
library(purrr)
library(dplyr)
library(ggplot2)

library(mosaic)
library(statisticalModeling)
library(dbl)
library(readxl)
library(XLConnect)
library(openxlsx)
library(formatR)

setwd("C:/Users/Ameek Bhalla/Dropbox/Ameek/Analysis")
setwd("C:/Users/Bhalla/Dropbox/Ameek/Analysis")


# F0 ---------------------------

F0 <- loadWorkbook("F0.xlsx") #XLConnect command, requires Java
# F0L3 <- readWorkbook("F0.xlsx", 1) #alternative openxlsx command
#L6.1 <- readWorkbook("F0+F1, Eighth Lot Technical Replicate 1.xlsx", 4)


F0L3 <- bind_rows(lapply(F0[2][ , 1], function(x) {
  filter(F0[1], label == x) %>% mutate(genotype = F0[2][which(F0[2][ , 1] == x), 2], 
                                       media    = F0[2][which(F0[2][ , 1] == x), 3],
                                       sex      = F0[2][which(F0[2][ , 1] == x), 4])
}
))

createSheet(F0, "F0L3") #XLConnect command, requires Java
#addWorksheet(F0, "F0L3") #openxlsx command
writeWorksheet(F0, F0L3, "F0L3") #XLConnect command, requires Java
#writeDataTable(F0, "F0L3", F0L3) #openxlsx command
saveWorkbook(F0) #XLConnect command, requires Java
#saveWorkbook(F0) #openxlsx command

# F1 ---------------------------

F1 <- loadWorkbook("F1.xlsx")

F1L1 <- bind_rows(lapply(F1[2][ , 1], function(x) {
  filter(F1[1], label == x) %>% mutate(genotype = F1[2][which(F1[2][ , 1] == x), 2], 
                                       media    = F1[2][which(F1[2][ , 1] == x), 3],
                                       sex      = F1[2][which(F1[2][ , 1] == x), 4])
}
))

createSheet(F1, "F1L1")
writeWorksheet(F1, F1L1, "F1L1")
saveWorkbook(F1)


F1L3 <- bind_rows(lapply(F1[6][ , 1], function(x) {
  filter(F1[5], label == x) %>% mutate(genotype = F1[6][which(F1[6][ , 1] == x), 2], 
                                       media    = F1[6][which(F1[6][ , 1] == x), 3],
                                       sex      = F1[6][which(F1[6][ , 1] == x), 4])
}
))

createSheet(F1, "F1L3")
writeWorksheet(F1, F1L3, "F1L3")
saveWorkbook(F1)

# F2 ---------------------------
F2 <- loadWorkbook("F2.xlsx")

F2L1 <- bind_rows(lapply(F2[2][ , 1], function(x) {
  filter(F2[1], label == x) %>% mutate(genotype = F2[2][which(F2[2][ , 1] == x), 2], 
                                       media    = F2[2][which(F2[2][ , 1] == x), 3],
                                       sex      = F2[2][which(F2[2][ , 1] == x), 4])
}
))

createSheet(F2, "F2L1")
writeWorksheet(F2, F2L1, "F2L1")
saveWorkbook(F2)


F2L2 <- bind_rows(lapply(F2[6][ , 1], function(x) {
  filter(F2[5], label == x) %>% mutate(genotype = F2[6][which(F2[6][ , 1] == x), 2], 
                                       media    = F2[6][which(F2[6][ , 1] == x), 3],
                                       sex      = F2[6][which(F2[6][ , 1] == x), 4])
}
))

createSheet(F2, "F2L2")
writeWorksheet(F2, F2L2, "F2L2")
saveWorkbook(F2)

# F3 ---------------------------
F3 <- loadWorkbook("F3.xlsx")

F3L1 <- bind_rows(lapply(F3[2][ , 1], function(x) {
  filter(F3[1], label == x) %>% mutate(genotype = F3[2][which(F3[2][ , 1] == x), 2], 
                                       media    = F3[2][which(F3[2][ , 1] == x), 3],
                                       sex      = F3[2][which(F3[2][ , 1] == x), 4])
}
))

createSheet(F3, "F3L1")
writeWorksheet(F3, F3L1, "F3L1")
saveWorkbook(F3)


F3L2 <- bind_rows(lapply(F3[6][ , 1], function(x) {
  filter(F3[5], label == x) %>% mutate(genotype = F3[6][which(F3[6][ , 1] == x), 2], 
                                       media    = F3[6][which(F3[6][ , 1] == x), 3],
                                       sex      = F3[6][which(F3[6][ , 1] == x), 4])
}
))

createSheet(F3, "F3L2")
writeWorksheet(F3, F3L2, "F3L2")
saveWorkbook(F3)

# All------------------------------
column_types <- c("text", "text", "numeric", "numeric", "numeric", "numeric", "text", "text", "text", "text", "text", "date", "date", "date", "text")
F0L3 <- read_excel('F0.xlsx', 'F0L3', col_types = column_types)
F1L1 <- read_excel('F1.xlsx', 'F1L1', col_types = column_types)
F1L3 <- read_excel('F1.xlsx', 'F1L3', col_types = column_types)
F2L1 <- read_excel('F2.xlsx', 'F2L1', col_types = column_types)
F2L2 <- read_excel('F2.xlsx', 'F2L2', col_types = column_types)
F3L1 <- read_excel('F3.xlsx', 'F3L1', col_types = column_types)
F3L2 <- read_excel('F3.xlsx', 'F3L2', col_types = column_types)

samples <- bind_rows(F0L3, F1L1, F1L3, F2L1, F2L2, F3L1, F3L2)

table(samples$plate)
samples$plate <- gsub(pattern='1.1000000000000001', replacement='1.1', x=samples$plate)
samples$plate <- gsub(pattern='2.2000000000000002', replacement='2.2', x=samples$plate)
table(samples$genotype)
samples$genotype <- gsub(pattern='.1118', replacement='control', x=samples$genotype)
samples$genotype <- gsub(pattern='\\+/\\+; \\+/\\+', replacement='test', x=samples$genotype)
table(samples$media)
samples$media <- gsub(pattern='ls', replacement='LS', x=samples$media)
samples$media <- gsub(pattern='hs', replacement='HS', x=samples$media)
table(samples$sex)
samples$sex <- gsub(pattern='females', replacement='female', x=samples$sex)
samples$sex <- gsub(pattern='males', replacement='male', x=samples$sex)

samples$tag_corrected <- samples$tag - samples$tag_blank
samples$bca_corrected <- samples$bca - samples$bca_blank

samples$treatment_ended <- as.Date(samples$treatment_ended, "%Y%m%d")
samples$homogenization <- as.Date(samples$homogenization, "%Y%m%d")
samples$assay_date <- as.Date(samples$assay_date, "%Y%m%d")
samples$assay_time <- unite(samples, col = assayed_at, assay_date, assay_time, sep = " ")$assayed_at
samples$assay_time <- as.POSIXct(samples$assay_time, "%Y-%m-%d %I:%M:%S %p")
samples$assay_date <- NULL

# Standards------------------------------
column_types_stds <- c("text", "text", "numeric", "numeric", "text", "numeric", "date", "text")
F0L3_std <- filter(read_excel('F0.xlsx', 'F0L3 Standards', col_types = column_types_stds), remark != 1)
F1L1_std <- filter(read_excel('F1.xlsx', 'F1L1 Standards', col_types = column_types_stds), remark != 1)
F1L3_std <- filter(read_excel('F1.xlsx', 'F1L3 Standards', col_types = column_types_stds), remark != 1)
F2L1_std <- filter(read_excel('F2.xlsx', 'F2L1 Standards', col_types = column_types_stds), remark != 1)
F2L2_std <- filter(read_excel('F2.xlsx', 'F2L2 Standards', col_types = column_types_stds), remark != 1)
F3L1_std <- filter(read_excel('F3.xlsx', 'F3L1 Standards', col_types = column_types_stds), remark != 1)
F3L2_std <- filter(read_excel('F3.xlsx', 'F3L2 Standards', col_types = column_types_stds), remark != 1)

stds <- bind_rows(F0L3_std, F1L1_std, F1L3_std, F2L1_std, F2L2_std, F3L1_std, F3L2_std)
stds$plate <- gsub(pattern='1.1000000000000001', replacement='1.1', x=stds$plate)
stds$plate <- gsub(pattern='2.2000000000000002', replacement='2.2', x=stds$plate)
stds$label <- gsub(pattern='reagent Tag', replacement='reagent TAG', x=stds$label)

stds$time <- unite(stds, col = assayed_at, date, time, sep = " ")$assayed_at
stds$time <- as.POSIXct(stds$time, "%Y-%m-%d %I:%M:%S %p")
stds$date <- NULL

#Changed the labels to simple "TAG" and "BCA" to allow group_by later
stds$label <- gsub(pattern='.*TAG.*', replacement='TAG', x=stds$label)
stds$label <- gsub(pattern='.*BCA.*', replacement='BCA', x=stds$label)
stds <- rename(stds, assay_type = label)
  
#reorganized the stds data so that each plate has a unique row for either assay, 
#and all other variables for that plate are nested in its row
stds_reorganized <- stds %>%
  group_by(generation, assay_type, plate, time) %>%
  nest()

parabolic_model <- function(df) {
  lm(absorbance ~ concentration + I(concentration^2), data = df)
} 

linear_model <- function(df) {
  lm(absorbance ~ concentration, data = df)
} 

std_models <- stds_reorganized %>%
  mutate(
    parabolic_mod = map(data, parabolic_model),
    linear_mod    = map(data, linear_model))

#unique std times
std_times <- unique(std_models$time)

nearest_time <- function(x) { 
  std_times[which.min(abs(std_times-x))]}

samples <- samples %>%
   mutate(
     closest_assay =  map_dbl(samples$assay_time, nearest_time))

#root for ax^2+bx+c = 0  
quadratic_function <- function(a, b, c, d = 0){
    pos_root <- ((-b) + sqrt((b^2) - 4*a*(c-d))) / (2*a)
    neg_root <- ((-b) - sqrt((b^2) - 4*a*(c-d))) / (2*a)
    return(pos_root)
  }

linear_function <- function(y,m,c){
  x <- ((y-c)/m)
  return(x)
}

para_model_tag <- bind_rows(map(1:nrow(samples), function(x){
  filter(std_models, generation == samples$generation[x], 
                     assay_type == 'TAG',
                     time == samples$closest_assay[x]) %>% select(parabolic_mod)
  })) 

para_model_bca <- bind_rows(map(1:nrow(samples), function(x){
  filter(std_models, generation == samples$generation[x], 
         assay_type == 'BCA',
         time == samples$closest_assay[x]) %>% select(parabolic_mod)
  }))

para_tag_calculated <- pmap_dbl(list(c =  map(para_model_tag$parabolic_mod, c("coefficients", '(Intercept)')),
                           d = samples$tag,
              b =  map(para_model_tag$parabolic_mod, c("coefficients", 'concentration')),
              a =  map(para_model_tag$parabolic_mod, c("coefficients", 'I(concentration^2)'))
), quadratic_function)
  

para_bca_calculated <- pmap_dbl(list(c =  map(para_model_bca$parabolic_mod, c("coefficients", '(Intercept)')),
                           d = samples$bca,
              b =  map(para_model_bca$parabolic_mod, c("coefficients", 'concentration')),
              a =  map(para_model_bca$parabolic_mod, c("coefficients", 'I(concentration^2)'))
), quadratic_function) 

lin_model_tag <- bind_rows(map(1:nrow(samples), function(x){
  filter(std_models, generation == samples$generation[x], 
         assay_type == 'TAG',
         time == samples$closest_assay[x]) %>% select(linear_mod)
})) 

lin_model_bca <- bind_rows(map(1:nrow(samples), function(x){
  filter(std_models, generation == samples$generation[x], 
         assay_type == 'BCA',
         time == samples$closest_assay[x]) %>% select(linear_mod)
}))

linear_tag_calculated <- pmap_dbl(list(c =  map(lin_model_tag$linear_mod, c("coefficients", '(Intercept)')),
                                  y =  samples$tag,
                                  m =  map(lin_model_tag$linear_mod, c("coefficients", 'concentration'))
), linear_function)


linear_bca_calculated <- pmap_dbl(list(c =  map(lin_model_bca$linear_mod, c("coefficients", '(Intercept)')),
                                  y =  samples$bca,
                                  m =  map(lin_model_bca$linear_mod, c("coefficients", 'concentration'))
), linear_function)


samples2 <- bind_cols(samples, lin_model_tag, lin_model_bca)
samples2 <- rename(samples2, tag_linear_model = linear_mod)
samples2 <- rename(samples2, bca_linear_model = linear_mod1)

samples2$tag_calculated <- linear_tag_calculated
samples2$bca_calculated <- linear_bca_calculated

closest_assay <- as.POSIXlt.POSIXct(samples2$closest_assay, "GMT")
two_earliest_assays <- samples2 %>% 
  distinct(closest_assay) %>% 
  top_n(n = -2) %>% 
  unlist() %>%
  as.POSIXlt.POSIXct(tz = "GMT")

#conclusion: only the one assay was perfomed before 14th Feb
#(NOTE: to convert mg/dl to ng/ml, multiply TAG sample values with 5000 for old kit & 10000 for new kit 
# new kit used since Feb 14
#eg: as in file named "LS media Transgenerational result March")

samples3 <- samples2 %>% 
  mutate(tag_ng_ml = 
           if_else(samples2$closest_assay == min(samples2$closest_assay), 
                   samples2$tag_calculated*5000, 
                   samples2$tag_calculated*10000)
         )

samples3$closest_assay <- as.POSIXlt.POSIXct(samples2$closest_assay, "GMT")
                            
samples3$normalized_tag <- (samples3$tag_ng_ml/samples3$bca_calculated)

All <- loadWorkbook("All.xlsx")
createSheet(All, "samples3")
writeWorksheet(All, samples3, "samples3")
saveWorkbook(All)

                            
#____________________________________
final %>% gather(key = model_type, value = tag_final, 2:3) %>% 
  ggplot(aes(x = log(tag_final), fill = as.factor(model_type))) +
  geom_density(alpha = .3)


group_stats <- samples3 %>%
  select(generation, tag_blank, tag, bca_blank, bca, genotype, media, sex, remark, 
         tag_corrected, bca_corrected, tag_calculated, bca_calculated, tag_ng_ml, normalized_tag) %>%
  group_by(generation, genotype, media, sex) %>% 
  nest()



group_stats2 <- filter(group_stats, sex == 'male') 
map(group_stats2$data, 'tag_ng_ml') %>%  geom_boxplot() 
 
%>% summarise(average = mean(tag_ng_ml))
barplot(group_stats$average)

ggplot(samples3, aes(x = log(normalized_tag))) +
  geom_density()


#trip_fit <- stds %>% 
#  group_by(generation, label, plate) %>%
#  summarise(linear_mse = mean(residuals(lm(absorbance ~ concentration))^2), 
#            linear_r2 = rsquared(lm(absorbance ~ concentration)),
#            exp_mse = mean(residuals(lm(log(absorbance) ~ concentration))^2), 
#            exp_r2 = rsquared(lm(log(absorbance) ~ concentration)),
#            parab_mse = mean(residuals(lm(absorbance ~ concentration + I(concentration^2)))^2),
#            parab_r2 = rsquared(lm(absorbance ~ concentration+ I(concentration^2))))
#Checking that parabolic MSEs are lower than linear MSEs (trip_fit[4]-trip_fit[8]) > 0
#Checking that parabolic R2s are higher than linear R2s  (trip_fit[5]-trip_fit[9]) < 0
#Conclusion parabolic regression is superior for both TAG and BCA assays

#fmodel(std_model_F0L3) + geom_point(aes(x=concentration, y= F0L3_std[c(1:9), 4]))

#gf_point(absorbance ~ concentration, data = test)

#mutate_each(F0L3, funs(as.factor), plate:remark)
#new_tibble <- gather(new, my_key, my_value, x1:x12)

#Rename the levels of factor_vector
#levels(factor_vector) <- c("name1", "name2",...)																																												
