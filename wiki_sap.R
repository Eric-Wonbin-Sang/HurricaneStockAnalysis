
if("rvest" %in% rownames(installed.packages()) == FALSE) {
  install.packages("rvest",repos = "http://cran.us.r-project.org")
  print(" - library rvest installed")
}

library(rvest)


print("-------- Parsing Wikipedia page... --------")
url <- "https://en.wikipedia.org/wiki/List_of_S%26P_500_companies"
SP500 <- url %>%
  read_html() %>%
  html_nodes(xpath='//*[@id="mw-content-text"]/div/table[1]') %>%
  html_table(fill = TRUE)
SP500 <- SP500[[1]]
Tix <- SP500$`Symbol`
print(" - Tic vector created")
# print(Tix)
Sec <- SP500$`Security`
print(" - Sec vector created")
# print(Sec)
Gics <- SP500$`GICS Sector`
print(" - Gics vector created")
# print(Gics)
Gics_sub <- SP500$`GICS Sub Industry`
print(" - Gisc_sub vector created")
# print(Gics_sub)
Loc <- SP500$`Location`
print(" - Loc vector created")
# print(Loc)

fileConn <- file("wiki_sap.txt")
data <- c()
for (i in c(1:length(Tix))){
  data <- c(data, paste(Tix[i], Sec[i], Gics[i], Gics_sub[i], Loc[i], sep = "; "))
}
writeLines(data, fileConn)
close(fileConn)
print("-------- wiki_sap.txt saved! --------")