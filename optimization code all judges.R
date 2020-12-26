# loading libraries
require(lpSolve)
require(xlsx)
require(prodlim)

# Read the data
library(readxl)
# Case_Time_DF <- read_excel("E:/daksh/Case Time Matrix.xlsx")
Case_Time_DF <- read_excel("Case Time Matrix.xlsx")
dim(Case_Time_DF)

# Change the Row names as First Column
library(tidyverse)
Case_Time_DF1 <- Case_Time_DF %>% remove_rownames %>% column_to_rownames(var="CaseNumber")

# Replace the Not filled values of time with -1.
Case_Time_DF2 <- Case_Time_DF1 %>% replace(., is.na(.),-1)

# Convert Dataframe into a matrix
Case_Time_Matrix = as.matrix.data.frame(Case_Time_DF2)

# Type Stage - Case Counts Table for Each Judge
JudgeNames <- unique(Case_Time_Matrix["CCH25"])
onejudgedata <- Case_Time_Matrix[Case_Time_Matrix["CCH25"]]


# Setup LP
nCases=nrow(Case_Time_Matrix)
nJudges=ncol(Case_Time_Matrix)

# For each Judge
# How many decision variables? - nCases
# How many constraints ? 
# 1. Constraint of Time for Each Judge = 1 ( Time must be less than equal to 360)
# 2. All decision variables should be binary.

f.obj <- as.numeric(Case_Time_Matrix)
f.con <- t(Case_Time_Matrix)
f.dir <- rep("<=",nJudges)
f.rhs <- rep(360,nJudges)
#
# Now run.
#
# Clear some space
rm(Case_Time_DF,Case_Time_DF1,Case_Time_DF2)

mod1 = lp ("max", f.obj, f.con, f.dir, f.rhs, all.bin = TRUE)
View(mod1$solution)

mod1_fnl_solution = matrix(mod1$solution,nrow=nrow(Case_Time_Matrix),ncol=ncol(Case_Time_Matrix),byrow=TRUE)
colnames(mod1_fnl_solution) <- colnames(Case_Time_Matrix)
rownames(mod1_fnl_solution) <- rownames(Case_Time_Matrix)


# Write the output
library(xlsx)
mod1_fnl_solution_df <- as.data.frame.matrix(mod1_fnl_solution)
write.xlsx(mod1_fnl_solution_df,"OptimizationSolutionCT.xlsx",sheetName = "JudgesName", append = TRUE)


