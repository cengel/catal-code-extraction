###### 
##### Before running this script use the AWK command that exports
##### indiviual CSV files from each VBA original source.
######

# This script generates one huge table with all the scripts and 
# - adds team and filename column
# - makes sure blank lines of code have no additional spaces

# TODO:
# - marks everything between private Sub and End Sub with subroutine name
# - marks lines with option compare scripts (only when the only thing in script)
# - marks comments according to
# -- commented code
# -- developer's comment
# -- use stopword file? use numbers in comment field?

# For analysis: 
# - remove blank lines with: d$code != ""
# - remove scripts that only are 2(?) lines

setwd("/Users/cengel/Anthro/Catal/CodeAnalysis/code_extraction/CodeWithCSV")

d <- NULL
for (csv in dir(pattern = "\\.csv$", full.names=TRUE, recursive=TRUE)){
  d <- rbind(d, cbind(read.table(csv, sep = "\t", quote="", stringsAsFactors = F), csv))
}
names(d)[1:3] <- c("line_no", "single_quots", "code")

# make empty fields consistent
# later, for analysis remove blank lines with d$code != ""
d$code <- gsub("[[:space:]]", "", d$code) 

d$team <- sapply(strsplit(as.character(d[,"csv"]), "/"), "[[", 2)
d$filename <- sapply(strsplit(as.character(d[,"csv"]), "/"), "[[", 3)
d$filename <- sapply(strsplit(as.character(d[,"filename"]), "[.]"), "[[", 1)

# write out
write.csv(d, "../CodeAll.csv", row.names = F)
