This is a script to analyze sorted stock ticker data. To run script download and import module1.bas.
The only modules that need to be run are summarize_worksheet and summarize_all.
Script was broken into submodules to increase readability.

The modules are as follows:
  -summarize_worksheet: this works on the current selected worksheet
  -summarize_all: This loops through the worksheets in a given workbook and calls summarize_worksheet for each
  -format: formats cells and adds column labels and row labels where appropriate. Called by summarize_worksheet
  -Makerow: formats and creates a row summarizing data for each ticker symbol. Called by summarize_worksheet
  -greatest: Bonus module that finds greatest increase/decrease/total stockvolume. Called by summarize_worksheet

Some notes: If a stock had value 0 at the beginning of the year, it's percent change is recorded as the string "NA".
For conditional formatting, if the stock increased it was colored green, if it decreased green, and if there was no change it was colored blue.
