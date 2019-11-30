options(java.parameters = "- Xmx10240m")

stock_to_sheet <- function(ticker, sheet){
  stock_data <- data.frame(loadSymbols(ticker, auto.assign = F))
  addDataFrame(stock_data, sheet=sheet, row.names=TRUE)
  cat("\r", paste("\t - ", ticker, "recieved"))
  # print(paste(" - ", ticker, "recieved"))
}

if("quantmod" %in% rownames(installed.packages()) == FALSE) {
  install.packages("quantmod",repos = "http://cran.us.r-project.org")
  print(" - library quantmod installed")
}
if("xlsx" %in% rownames(installed.packages()) == FALSE) {
  install.packages("xlsx",repos = "http://cran.us.r-project.org")
  print(" - library xlsx installed")
}

library(quantmod)
library(xlsx)
print("-------- Reading tickers.txt... ---------")
filename <- "tickers.txt"
companies <- read.table(filename, sep = '')
companies <- as.vector(t(companies))

options("getSymbols.yahoo.warning"=FALSE)
options("getSymbols.warning4.0"=FALSE)

workbook = createWorkbook()
print("Appending stocks to workbook...")
for (ticker in companies) {
  sheet = createSheet(workbook, ticker)
  try(stock_to_sheet(ticker, sheet))
}

saveWorkbook(workbook, "Stock_Data.xlsx")
print("")
print("-------- Workbook Saved! ---------")
