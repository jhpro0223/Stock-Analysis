# Stock_Analysis
Performing analysis on Stock Data with VBA
# STOCK ANALYSIS WITH VBA + EXCEL
## OVERVIEW: VBA Stock Analysis Project

### Purpose

The purpose of the project is to edit and refactor the Stock Market Dataset with VBA code and determine whether refactoring the code successfully made the VBA script run faster. The goal is to make the code more efficient— less steps, using less memory, or improving the logic of the code to make it easier for users to read. 

### Analysis and Challenges
Analysis and Challenges of this Project:
- Prepare the data set `VBA_Challenge.vbs` file for the project.
- Convert `XLSM` file from `*.vbs` dataset as `VBA_Challenge.xlsm`.
- Refactor VBA code and measure performance and add code where indicated by the numbered comments in the starter code file.
> Use the starter code provided in this Project to refactor the VBA Script dataset and loop through the data one time and collect all of the information.
#### Challenge Data Background
> Steve loves would like to expand the dataset and include the entire stock market over the last few years. Although the code works well for a dozen stocks, it might not work as well for thousands of stocks. And if it does, it may take a long time to execute.
> challenge - edit, or refactor, the solution code to loop through all the data one time in order to collect the same information that has already been done in the previous module. Then, determine whether refactoring the code successfully made the VBA script run faster. 
> Refactoring, as a key part of the coding, doesnt refer to adding new functionality. It is about finding ways to make the code more efficient. 
## RESULTS: Refactor VBA Code and Measure Performance
 
### Deliverable Compare Stock Performance and Timestamp procedure below:

**1. The `tickerIndex` is set equal to zero before looping over the rows.**

![name-of-you-image](https://github.com/Dorislava/Stock_Analysis/blob/main/TickerIndex.PNG)

**2. Arrays are created for `tickers`, `tickerVolumes`, `tickerStartingPrices`, and `tickerEndingPrices`.**
![name-of-you-image](https://github.com/Dorislava/Stock_Analysis/blob/main/Arrays%20created.PNG)

**3. The `tickerIndex` is used to access the stock ticker index for the `tickers`, `tickerVolumes`, `tickerStartingPrices`, and `tickerEndingPrices` arrays.**

![name-of-you-image](https://github.com/Dorislava/Stock_Analysis/blob/main/TickerIndex%20%20to%20access%20the%20stock%20ticker.PNG)

**4. The script loops through stock data, reading and storing all of the following values from each row: `tickers`, `tickerVolumes`, `tickerStartingPrices`, and `tickerEndingPrices`.**

![name-of-you-image](https://github.com/Dorislava/Stock_Analysis/blob/main/Loops%20through%20Stock%20data.PNG)

**5. Code for formatting the cells in the spreadsheet is working.**

![name-of-you-image](https://github.com/Dorislava/Stock_Analysis/blob/main/Formatting.PNG)

**6. The outputs for the 2017 and 2018 stock analyses in the `VBA_Challenge.xlsm` workbook match the outputs from the AllStockAnalysis in the module**

***Dataset Examples***

![name-of-you-image](https://github.com/Dorislava/Stock_Analysis/blob/main/All%20Stocks%202017.PNG)
![name-of-you-image](https://github.com/Dorislava/Stock_Analysis/blob/main/All%20Stocks%202018.PNG)

**Final VBA Analysis 2017 and 2018** 

**7. The pop-up messages showing the elapsed run time for the script are saved as `VBA_Challenge_2017.png` and `VBA_Challenge_2018.png`**

***Time on VBA_Challenge_2017.PNG***

![name-of-you-image](https://github.com/Dorislava/Stock_Analysis/blob/main/VBA_Challenge_2017.PNG)

***Time on VBA_Challenge_2018.PNG***
![name-of-you-image](https://github.com/Dorislava/Stock_Analysis/blob/main/VBA_Challenge_2018.PNG)

> Running our fully 2017 and 2018 data stock analysis gave us an elapsed run time for each year.

## SUMMARY:
> - The code refactoring should be performed in small steps to improve the code.
> - A long procedure may contain the same line of code in several locations, opportunity - change the logic to eliminate the duplicate lines.
> - A complex unstructured code is better to split in several functions. 
A clean and well-organized code is always easy to understand, change and maintain. 
