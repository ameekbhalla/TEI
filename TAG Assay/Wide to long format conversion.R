library(readxl)
library(dplyr)
library(tidyr)
library(XLConnect)
library(purrr)
library(ggplot2)
library(knitr)
library(lubridate)

 setwd("C:/Users/Ameek Bhalla/Dropbox/Ameek/Analysis")
# setwd("C:/Users/Bhalla/Dropbox/Ameek/Analysis")

##Change this to automatically extract sheet names from each file and then write corresponding sheets with a suffix

##===================================================================================

abc <- read_excel("F0+F1, Sixth Lot Technical Replicate 1.xlsx", sheet = 1, col_names = T) 
## extract 500nm col
a1 <- abc %>%
  slice(1:8) %>%
  gather(2:13,
         key = "col", value = "500nm")#returns a three column df
## extract 562nm col
a2 <- abc %>%
  slice(10:17) %>%
  gather(2:13,
         key = "col", value = "562nm")
## extract labels
a3 <- abc %>%
  slice(19:26) %>%
  gather(2:13, 
         key = "col", value = "label")
## extract labels
a4 <- abc %>%
  slice(28:35) %>%
  gather(2:13, 
         key = "col", value = "reagent")
## extract labels
a5 <- abc %>%
  slice(37:44) %>%
  gather(2:13, 
         key = "col", value = "remark")

plate1 <- bind_cols(a1, a2[,3], a3[,3], a4[,3], a5[,3]) #only extract the third column from the latter four df

#------------------------------------------------------------------------------------------------------------

abc <- read_excel("F0+F1, Sixth Lot Technical Replicate 1.xlsx", sheet = 2, col_names = T) 
## extract 500nm col
a1 <- abc %>%
  slice(1:8) %>%
  gather(2:13,
         key = "col", value = "500nm")#returns a three column df
## extract 562nm col
a2 <- abc %>%
  slice(10:17) %>%
  gather(2:13,
         key = "col", value = "562nm")
## extract labels
a3 <- abc %>%
  slice(19:26) %>%
  gather(2:13, 
         key = "col", value = "label")
## extract labels
a4 <- abc %>%
  slice(28:35) %>%
  gather(2:13, 
         key = "col", value = "reagent")
## extract labels
a5 <- abc %>%
  slice(37:44) %>%
  gather(2:13, 
         key = "col", value = "remark")

plate2 <- bind_cols(a1, a2[,3], a3[,3], a4[,3], a5[,3]) #only extract the third column from the latter four df

#------------------------------------------------------------------------------------------------------------

abc <- read_excel("F0+F1, Sixth Lot Technical Replicate 1.xlsx", sheet = 3, col_names = T) 
## extract 500nm col
a1 <- abc %>%
  slice(1:8) %>%
  gather(2:13,
         key = "col", value = "500nm")#returns a three column df
## extract 562nm col
a2 <- abc %>%
  slice(10:17) %>%
  gather(2:13,
         key = "col", value = "562nm")
## extract labels
a3 <- abc %>%
  slice(19:26) %>%
  gather(2:13, 
         key = "col", value = "label")
## extract labels
a4 <- abc %>%
  slice(28:35) %>%
  gather(2:13, 
         key = "col", value = "reagent")
## extract labels
a5 <- abc %>%
  slice(37:44) %>%
  gather(2:13, 
         key = "col", value = "remark")

plate3 <- bind_cols(a1, a2[,3], a3[,3], a4[,3], a5[,3]) #only extract the third column from the latter four df

L6 <- bind_rows("1" = plate1, "2" = plate2, "3" = plate3, .id = "plate")

##===================================================================================

abc <- read_excel("F0+F1, Eighth Lot Technical Replicate 1.xlsx", sheet = 1, col_names = T) 
## extract 500nm col
a1 <- abc %>%
  slice(1:8) %>%
  gather(2:13,
         key = "col", value = "500nm")#returns a three column df
## extract 562nm col
a2 <- abc %>%
  slice(10:17) %>%
  gather(2:13,
         key = "col", value = "562nm")
## extract labels
a3 <- abc %>%
  slice(19:26) %>%
  gather(2:13, 
         key = "col", value = "label")
## extract labels
a4 <- abc %>%
  slice(28:35) %>%
  gather(2:13, 
         key = "col", value = "reagent")
## extract labels
a5 <- abc %>%
  slice(37:44) %>%
  gather(2:13, 
         key = "col", value = "remark") 

plate4 <- bind_cols(a1, a2[,3], a3[,3], a4[,3], a5[,3]) #only extract the third column from the latter four df

#------------------------------------------------------------------------------------------------------------

abc <- read_excel("F0+F1, Eighth Lot Technical Replicate 1.xlsx", sheet = 2, col_names = T) 
## extract 500nm col
a1 <- abc %>%
  slice(1:8) %>%
  gather(2:13,
         key = "col", value = "500nm")#returns a three column df
## extract 562nm col
a2 <- abc %>%
  slice(10:17) %>%
  gather(2:13,
         key = "col", value = "562nm")
## extract labels
a3 <- abc %>%
  slice(19:26) %>%
  gather(2:13, 
         key = "col", value = "label")
## extract labels
a4 <- abc %>%
  slice(28:35) %>%
  gather(2:13, 
         key = "col", value = "reagent")
## extract labels
a5 <- abc %>%
  slice(37:44) %>%
  gather(2:13, 
         key = "col", value = "remark")

plate5 <- bind_cols(a1, a2[,3], a3[,3], a4[,3], a5[,3]) #only extract the third column from the latter four df

#------------------------------------------------------------------------------------------------------------

abc <- read_excel("F0+F1, Eighth Lot Technical Replicate 1.xlsx", sheet = 3, col_names = T) 
## extract 500nm col
a1 <- abc %>%
  slice(1:8) %>%
  gather(2:13,
         key = "col", value = "500nm")#returns a three column df
## extract 562nm col
a2 <- abc %>%
  slice(10:17) %>%
  gather(2:13,
         key = "col", value = "562nm")
## extract labels
a3 <- abc %>%
  slice(19:26) %>%
  gather(2:13, 
         key = "col", value = "label")
## extract labels
a4 <- abc %>%
  slice(28:35) %>%
  gather(2:13, 
         key = "col", value = "reagent")
## extract labels
a5 <- abc %>%
  slice(37:44) %>%
  gather(2:13, 
         key = "col", value = "remark")

plate6 <- bind_cols(a1, a2[,3], a3[,3], a4[,3], a5[,3]) #only extract the third column from the latter four df

#------------------------------------------------------------------------------------------------------------

abc <- read_excel("F0+F1, Eighth Lot Technical Replicate 1.xlsx", sheet = 4, col_names = T) 
## extract 500nm col
a1 <- abc %>%
  slice(1:8) %>%
  gather(2:13,
         key = "col", value = "500nm")#returns a three column df
## extract 562nm col
a2 <- abc %>%
  slice(10:17) %>%
  gather(2:13,
         key = "col", value = "562nm")
## extract labels
a3 <- abc %>%
  slice(19:26) %>%
  gather(2:13, 
         key = "col", value = "label")
## extract labels
a4 <- abc %>%
  slice(28:35) %>%
  gather(2:13, 
         key = "col", value = "reagent")
## extract labels
a5 <- abc %>%
  slice(37:44) %>%
  gather(2:13, 
         key = "col", value = "remark")

plate7 <- bind_cols(a1, a2[,3], a3[,3], a4[,3], a5[,3]) #only extract the third column from the latter four df

#When the argument .id is supplied, a new column is created with the name which is the value of the .id argument. 
#The column stores the names of the original DF from which each row was extracted. The original DFs must have names

L8 <- bind_rows("4" = plate4, "5" = plate5, "6" = plate6, "7" = plate7, .id = "plate")

##===================================================================================

abc <- read_excel("Tenth lot_males_Ameek LS & HS.xlsx", sheet = 1, col_names = T) 
## extract 500nm col
a1 <- abc %>%
  slice(1:8) %>%
  gather(2:13,
         key = "col", value = "500nm")#returns a three column df
## extract 562nm col
a2 <- abc %>%
  slice(10:17) %>%
  gather(2:13,
         key = "col", value = "562nm")
## extract labels
a3 <- abc %>%
  slice(19:26) %>%
  gather(2:13, 
         key = "col", value = "label")
## extract labels
a4 <- abc %>%
  slice(28:35) %>%
  gather(2:13, 
         key = "col", value = "reagent")
## extract labels
a5 <- abc %>%
  slice(37:44) %>%
  gather(2:13, 
         key = "col", value = "remark")

plateA <- bind_cols(a1, a2[,3], a3[,3], a4[,3], a5[,3]) #only extract the third column from the latter four df

#------------------------------------------------------------------------------------------------------------

abc <- read_excel("Tenth lot_males_Ameek LS & HS.xlsx", sheet = 3, col_names = T) 
## extract 500nm col
a1 <- abc %>%
  slice(1:8) %>%
  gather(2:13,
         key = "col", value = "500nm")#returns a three column df
## extract 562nm col
a2 <- abc %>%
  slice(10:17) %>%
  gather(2:13,
         key = "col", value = "562nm")
## extract labels
a3 <- abc %>%
  slice(19:26) %>%
  gather(2:13, 
         key = "col", value = "label")
## extract labels
a4 <- abc %>%
  slice(28:35) %>%
  gather(2:13, 
         key = "col", value = "reagent")
## extract labels
a5 <- abc %>%
  slice(37:44) %>%
  gather(2:13, 
         key = "col", value = "remark")

plateB <- bind_cols(a1, a2[,3], a3[,3], a4[,3], a5[,3]) #only extract the third column from the latter four df

#------------------------------------------------------------------------------------------------------------

abc <- read_excel("Tenth lot_males_Ameek LS & HS.xlsx", sheet = 4, col_names = T) 
## extract 500nm col
a1 <- abc %>%
  slice(1:8) %>%
  gather(2:13,
         key = "col", value = "500nm")#returns a three column df
## extract 562nm col
a2 <- abc %>%
  slice(10:17) %>%
  gather(2:13,
         key = "col", value = "562nm")
## extract labels
a3 <- abc %>%
  slice(19:26) %>%
  gather(2:13, 
         key = "col", value = "label")
## extract labels
a4 <- abc %>%
  slice(28:35) %>%
  gather(2:13, 
         key = "col", value = "reagent")
## extract labels
a5 <- abc %>%
  slice(37:44) %>%
  gather(2:13, 
         key = "col", value = "remark")

plateC <- bind_cols(a1, a2[,3], a3[,3], a4[,3], a5[,3]) #only extract the third column from the latter four df

#------------------------------------------------------------------------------------------------------------

abc <- read_excel("Tenth lot_males_Ameek LS & HS.xlsx", sheet = 5, col_names = T) 
## extract 500nm col
a1 <- abc %>%
  slice(1:8) %>%
  gather(2:13,
         key = "col", value = "500nm")#returns a three column df
## extract 562nm col
a2 <- abc %>%
  slice(10:17) %>%
  gather(2:13,
         key = "col", value = "562nm")
## extract labels
a3 <- abc %>%
  slice(19:26) %>%
  gather(2:13, 
         key = "col", value = "label")
## extract labels
a4 <- abc %>%
  slice(28:35) %>%
  gather(2:13, 
         key = "col", value = "reagent")
## extract labels
a5 <- abc %>%
  slice(37:44) %>%
  gather(2:13, 
         key = "col", value = "remark")

plateD <- bind_cols(a1, a2[,3], a3[,3], a4[,3], a5[,3]) #only extract the third column from the latter four df

#------------------------------------------------------------------------------------------------------------

abc <- read_excel("Tenth lot_females_Ameek LS.xlsx", sheet = 1, col_names = T) 
## extract 500nm col
a1 <- abc %>%
  slice(1:8) %>%
  gather(2:13,
         key = "col", value = "500nm")#returns a three column df
## extract 562nm col
a2 <- abc %>%
  slice(10:17) %>%
  gather(2:13,
         key = "col", value = "562nm")
## extract labels
a3 <- abc %>%
  slice(19:26) %>%
  gather(2:13, 
         key = "col", value = "label")
## extract labels
a4 <- abc %>%
  slice(28:35) %>%
  gather(2:13, 
         key = "col", value = "reagent")
## extract labels
a5 <- abc %>%
  slice(37:44) %>%
  gather(2:13, 
         key = "col", value = "remark")

`1LS` <- bind_cols(a1, a2[,3], a3[,3], a4[,3], a5[,3]) #only extract the third column from the latter four df
#it is necessary to enclose names like 1LS and 1HS in `` otherwise you get an error

#-----------------------------------------------------------------------------------------------------------

abc <- read_excel("Tenth lot_females_Ameek LS.xlsx", sheet = 2, col_names = T) 
## extract 500nm col
a1 <- abc %>%
  slice(1:8) %>%
  gather(2:13,
         key = "col", value = "500nm")#returns a three column df
## extract 562nm col
a2 <- abc %>%
  slice(10:17) %>%
  gather(2:13,
         key = "col", value = "562nm")
## extract labels
a3 <- abc %>%
  slice(19:26) %>%
  gather(2:13, 
         key = "col", value = "label")
## extract labels
a4 <- abc %>%
  slice(28:35) %>%
  gather(2:13, 
         key = "col", value = "reagent")
## extract labels
a5 <- abc %>%
  slice(37:44) %>%
  gather(2:13, 
         key = "col", value = "remark")

`2LS` <- bind_cols(a1, a2[,3], a3[,3], a4[,3], a5[,3]) #only extract the third column from the latter four df

#------------------------------------------------------------------------------------------------------------

abc <- read_excel("Tenth lot_females_Ameek LS.xlsx", sheet = 3, col_names = T) 
## extract 500nm col
a1 <- abc %>%
  slice(1:8) %>%
  gather(2:13,
         key = "col", value = "500nm")#returns a three column df
## extract 562nm col
a2 <- abc %>%
  slice(10:17) %>%
  gather(2:13,
         key = "col", value = "562nm")
## extract labels
a3 <- abc %>%
  slice(19:26) %>%
  gather(2:13, 
         key = "col", value = "label")
## extract labels
a4 <- abc %>%
  slice(28:35) %>%
  gather(2:13, 
         key = "col", value = "reagent")
## extract labels
a5 <- abc %>%
  slice(37:44) %>%
  gather(2:13, 
         key = "col", value = "remark")

`3LS` <- bind_cols(a1, a2[,3], a3[,3], a4[,3], a5[,3]) #only extract the third column from the latter four df

#------------------------------------------------------------------------------------------------------------

abc <- read_excel("Tenth lot_females_Ameek LS.xlsx", sheet = 4, col_names = T) 
## extract 500nm col
a1 <- abc %>%
  slice(1:8) %>%
  gather(2:13,
         key = "col", value = "500nm")#returns a three column df
## extract 562nm col
a2 <- abc %>%
  slice(10:17) %>%
  gather(2:13,
         key = "col", value = "562nm")
## extract labels
a3 <- abc %>%
  slice(19:26) %>%
  gather(2:13, 
         key = "col", value = "label")
## extract labels
a4 <- abc %>%
  slice(28:35) %>%
  gather(2:13, 
         key = "col", value = "reagent")
## extract labels
a5 <- abc %>%
  slice(37:44) %>%
  gather(2:13, 
         key = "col", value = "remark")

`4LS` <- bind_cols(a1, a2[,3], a3[,3], a4[,3], a5[,3]) #only extract the third column from the latter four df

#------------------------------------------------------------------------------------------------------------

abc <- read_excel("Tenth lot_females_Ameek HS.xlsx", sheet = 1, col_names = T) 
## extract 500nm col
a1 <- abc %>%
  slice(1:8) %>%
  gather(2:13,
         key = "col", value = "500nm")#returns a three column df
## extract 562nm col
a2 <- abc %>%
  slice(10:17) %>%
  gather(2:13,
         key = "col", value = "562nm")
## extract labels
a3 <- abc %>%
  slice(19:26) %>%
  gather(2:13, 
         key = "col", value = "label")
## extract labels
a4 <- abc %>%
  slice(28:35) %>%
  gather(2:13, 
         key = "col", value = "reagent")
## extract labels
a5 <- abc %>%
  slice(37:44) %>%
  gather(2:13, 
         key = "col", value = "remark")

`1HS` <- bind_cols(a1, a2[,3], a3[,3], a4[,3], a5[,3]) #only extract the third column from the latter four df

#--------------------------------------------------------------------------------------------------------------

abc <- read_excel("Tenth lot_females_Ameek HS.xlsx", sheet = 2, col_names = T) 
## extract 500nm col
a1 <- abc %>%
  slice(1:8) %>%
  gather(2:13,
         key = "col", value = "500nm")#returns a three column df
## extract 562nm col
a2 <- abc %>%
  slice(10:17) %>%
  gather(2:13,
         key = "col", value = "562nm")
## extract labels
a3 <- abc %>%
  slice(19:26) %>%
  gather(2:13, 
         key = "col", value = "label")
## extract labels
a4 <- abc %>%
  slice(28:35) %>%
  gather(2:13, 
         key = "col", value = "reagent")
## extract labels
a5 <- abc %>%
  slice(37:44) %>%
  gather(2:13, 
         key = "col", value = "remark")

`2HS` <- bind_cols(a1, a2[,3], a3[,3], a4[,3], a5[,3]) #only extract the third column from the latter four df

#--------------------------------------------------------------------------------------------------------------

abc <- read_excel("Tenth lot_females_Ameek HS.xlsx", sheet = 3, col_names = T) 
## extract 500nm col
a1 <- abc %>%
  slice(1:8) %>%
  gather(2:13,
         key = "col", value = "500nm")#returns a three column df
## extract 562nm col
a2 <- abc %>%
  slice(10:17) %>%
  gather(2:13,
         key = "col", value = "562nm")
## extract labels
a3 <- abc %>%
  slice(19:26) %>%
  gather(2:13, 
         key = "col", value = "label")
## extract labels
a4 <- abc %>%
  slice(28:35) %>%
  gather(2:13, 
         key = "col", value = "reagent")
## extract labels
a5 <- abc %>%
  slice(37:44) %>%
  gather(2:13, 
         key = "col", value = "remark")

`3HS` <- bind_cols(a1, a2[,3], a3[,3], a4[,3], a5[,3]) #only extract the third column from the latter four df

#--------------------------------------------------------------------------------------------------------------

abc <- read_excel("Tenth lot_females_Ameek HS.xlsx", sheet = 4, col_names = T) 
## extract 500nm col
a1 <- abc %>%
  slice(1:8) %>%
  gather(2:13,
         key = "col", value = "500nm")#returns a three column df
## extract 562nm col
a2 <- abc %>%
  slice(10:17) %>%
  gather(2:13,
         key = "col", value = "562nm")
## extract labels
a3 <- abc %>%
  slice(19:26) %>%
  gather(2:13, 
         key = "col", value = "label")
## extract labels
a4 <- abc %>%
  slice(28:35) %>%
  gather(2:13, 
         key = "col", value = "reagent")
## extract labels
a5 <- abc %>%
  slice(37:44) %>%
  gather(2:13, 
         key = "col", value = "remark")

`4HS` <- bind_cols(a1, a2[,3], a3[,3], a4[,3], a5[,3]) #only extract the third column from the latter four df

L10 <- bind_rows("A" = plateA, "B" = plateB, "C" = plateC, "D" = plateD, 
                 "1LS" = `1LS`, "2LS" = `2HS`, "3LS" = `3LS`, "4LS" = `4LS`,
                 "1HS" = `1HS`, "2HS" = `2HS`, "3HS" = `3HS`, "4HS" = `4HS`,
                 .id = "plate")

##===================================================================================

names <- loadWorkbook("names.xlsx")
names["sixth"] <- mutate_at(names["sixth"], "label", as.character) #sheets must be subset by name & not postion in mutate()
names["eighth"] <- mutate_at(names["eighth"], "label", as.character)
names["tenth"] <- mutate_at(names["tenth"], "label", as.character)
L6 <- full_join(L6, names["sixth"], by = "label") #the by argument must have variables of same class
L8 <- full_join(L8, names["eighth"], by = "label") #hence "label" was transformed to a chr vector above
L10 <- full_join(L10, names["tenth"], by = "label") 

L6 <- L6 %>% mutate(
       treatment_ended = mdy("12/07/2017"), 
       homogenization = mdy("12/07/2017"), #should be the denaturtion date in all cases
       assay_date = mdy("04/18/2018"),
       assay_time = if_else(plate == "plate1", mdy_hms("4/18/2018 10:33:30 PM"),
                    if_else(plate == "plate2", mdy_hms("4/18/2018 10:26:18 PM"), mdy_hms("4/18/2018 10:30:06 PM")
                                    ))
       )
       
L8 <- L8 %>% mutate( 
       treatment_ended = mdy("02/22/2018"), 
       homogenization = mdy("02/23/2018"), 
       assay_date = mdy("04/18/2018"),
       assay_time = if_else(plate == "plate4", mdy_hms("4/18/2018 11:12:55 PM"),
                    if_else(plate == "plate5", mdy_hms("4/18/2018 11:09:39 PM"),
                    if_else(plate == "plate6", mdy_hms("4/18/2018 11:06:24 PM"), mdy_hms("4/18/2018 11:17:11 PM")
                               )))
       )

L10 <- L10 %>% mutate( 
       treatment_ended = mdy("06/29/2018"), 
       homogenization = if_else(sex == "female", mdy("06/29/2018"), mdy("07/12/2018")), #males denatured 2 weeks later
       assay_date = if_else(sex == "female", mdy("07/18/2018"), mdy("07/13/2018")),
       assay_time = if_else(plate == "1HS", mdy_hms("7/18/2018 8:44:39 PM"),
                    if_else(plate == "2HS", mdy_hms("7/18/2018 8:47:53 PM"),
                    if_else(plate == "3HS", mdy_hms("7/18/2018 8:51:00 PM"),
                    if_else(plate == "4HS", mdy_hms("7/18/2018 9:13:23 PM"),
                    if_else(plate == "1LS", mdy_hms("7/18/2018 9:26:01 PM"),
                    if_else(plate == "2LS", mdy_hms("7/18/2018 9:37:56 PM"),
                    if_else(plate == "3LS", mdy_hms("7/18/2018 9:22:31 PM"),
                    if_else(plate == "4LS", mdy_hms("7/18/2018 9:17:13 PM"),
                    if_else(plate == "A", mdy_hms("7/13/2018 9:15:15 PM"),
                    if_else(plate == "B", mdy_hms("7/13/2018 10:06:29 PM"),
                    if_else(plate == "C", mdy_hms("7/13/2018 9:46:36 PM"), mdy_hms("7/13/2018 9:51:55 PM")
                            )))))))))))
      )


# map(list(L6, L8, L10), colnames)

# F1_new <- bind_rows(list("L6" = L6, "L8" = L8, "L10" = L10), .id = "Replicate")


#add "treatment_ended"	"homogenization"	"assay_date"	"assay_time"
#in format "12/30/2016 00:00:00"	"01/02/2017 00:00:00"	"02/28/2017 00:00:00"	"8:56:43 PM"
#separate out standards data: "generation"	"label"	"concentration"	"absorbance"	"plate"	"remark"	"date"	"time"
#grepl tag, std etc
#calculate "tag_blank"	"tag"	"bca_blank"	"bca"
#exclude rows with NA or 1

#search for "#" and make a notes in google keep