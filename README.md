# VBA-challenge 
This script analyzes multiple years of stock data, creating new summary tables to calculate different metrics. 

Included in this repo are the VBA script with the StockAnalysis subroutine, a sample data set and images of the results when the script was run on a full data set.

## Metrics calculated
1. Summary of stocks in list by ticker
2. The annual change in value of the stock (dollar and percentage)
3. The stock with the greatest increaase annually
4. The stock with the greatest decrease annually
5. The stock with the greatest total value

## Overview
The script contains a routine called StockAnalysis that when run,
- Creates an output table to summarize the data
- Loops through each row in the spreadsheet
- Creates a list of all stock tickers in the sheet
- Outputs and winners and losers
- Continues to the next sheet to start again

### Setup 
In order for the script to work, you need a worksheet with the following columns with headers:
- The ticker symbol
- Date
- Daily opening price
- Daily high price
- Daily low price
- Daily closing price
- Daily volume

<img src="https://github.com/kmcmurphy/VBA-challenge/blob/main/worksheet_setup.png" width = "50%" alt="Worksheet setup" />

After the script ran on the full dataset, here are a summary of the results from each worksheet (year):

### 2018 Results

<img src="https://github.com/kmcmurphy/VBA-challenge/blob/main/2018_results.png" width = "50%" alt="2018 Results" />

### 2019 Results

<img src="https://github.com/kmcmurphy/VBA-challenge/blob/main/2019_results.png" width = "50%" alt="2018 Results" />

### 2020 Results

<img src="https://github.com/kmcmurphy/VBA-challenge/blob/main/2020_results.png" width = "50%" alt="2018 Results" />
