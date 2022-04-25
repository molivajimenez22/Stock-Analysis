# Stock-Analysis

## Overview of Project
The purpose of this project is to review stock information and make an assessment on whether the stock is a profitable investment based on historical information for 2017 and 2018.  Additionally, by defactoring the code, we are inscreasing efficiency of the VBA. 

## Results
Below are examples of Code Defactoring used to increase efficiency:

### 1. The tickerIndex is set equal to zero before looping over the rows
![name-of-you-image](https://github.com/molivajimenez22/Stock-Analysis/blob/main/tickerIndex_as_Variable.png)

### 2. Arrays are created for tickers, tickerVolumes, tickerStartingPrices and tickerEndingPrices 
![name-of-you-image](https://github.com/molivajimenez22/Stock-Analysis/blob/main/array_Final.png)

### 3. The tickerIndex is used to access the stock ticker index for the tickers, tickerVolumes, tickerStartingPrices, and tickerEndingPrices arrays 
![name-of-you-image](https://github.com/molivajimenez22/Stock-Analysis/blob/main/access_stock_ticker.png)

### 4. The script loops through stock data, reading and storing all of the following values from each row: tickers, tickerVolumes, tickerStartingPrices, and tickerEndingPrices 
![name-of-you-image](https://github.com/molivajimenez22/Stock-Analysis/blob/main/loop_data.png)

### 5. Defactored Analysis
#### Defactored Analysis for 2017
![name-of-you-image](https://github.com/molivajimenez22/Stock-Analysis/blob/main/VBA_Callenge_20172.png)

#### Defactored Analysis for 2018
![name-of-you-image](https://github.com/molivajimenez22/Stock-Analysis/blob/main/VBA_Callenge_20182.png)

## Summary 

### 1. What are the advantages or disadvantages of refactoring code?
Some advantages of defactoring code include the ability to identify errors that contain nested conditionals and loops as well as allowing one the ability to see patterns that are not easily seen in the source.

Some disadvantages of defactoring code center around the impact on testing outcomes as well the added work of solving for duplicates by changing logic.

### 2. How do these pros and cons apply to refactoring the original VBA script?
Ultimately, defactoring code allows a user to improve and update the code without impacting the software's functionality. As such, paying close attention to the defactoring code from the beginning will allow the user to continue to maintain clean and well organized data.

#### Additional Resources
[VBA Challenge](https://github.com/molivajimenez22/Stock-Analysis/blob/main/VBA_Challenge.xlsm)

[VBA Challenge 2017](https://github.com/molivajimenez22/Stock-Analysis/blob/main/VBA_Challenge_2017.png)

[VBA Challegen 2018](https://github.com/molivajimenez22/Stock-Analysis/blob/main/VBA_Challenge_2018.png)
