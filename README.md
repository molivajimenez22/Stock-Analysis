# Stock-Analysis

### Overview of Project
The purpose of this project is to review stock information and make an assessment on whether the stock is a profitable investment based on historical information for 2017 and 2018.  Additionally, by defactoring the code, we are inscreasing efficiency of the VBA. 

### Results
The variable created was `tickerIndex`. This variable was then used to access the arrays including `ticker`, `tickerVolumes`, `tickerStartinGPrices` and `tickerEndingPrices`. 

`tickerIndex = 0`

![name-of-you-image](https://github.com/molivajimenez22/Stock-Analysis/blob/main/Resources/arrays.png)

**Accessing the stock ticker index for the tickers, tickerVolumes, tickerStartingPrices, and tickerEndingPrices arrays** 
![name-of-you-image](https://github.com/molivajimenez22/Stock-Analysis/blob/main/Resources/Accessing.png)

**Script loops through stock data, reading and storing all of the following values from each row: tickers, tickerVolumes, tickerStartingPrices, and tickerEndingPrices**

![name-of-you-image](https://github.com/molivajimenez22/Stock-Analysis/blob/main/Resources/loop.png)

### 2017 Analysis
<p align="center" width="100%">
    <img width="80%" src="https://github.com/molivajimenez22/Stock-Analysis/blob/main/Resources/2017Analysis.png">

#### 2018 Analysis
<p align="center" width="100%">
    <img width="80%" src="https://github.com/molivajimenez22/Stock-Analysis/blob/main/Resources/2018Analysis.png">

## Summary 

  **1. What are the advantages or disadvantages of refactoring code?**
Some advantages of defactoring code include the ability to identify errors that contain nested conditionals and loops as well as allowing one the ability to see patterns that are not easily seen in the source.

Some disadvantages of defactoring code center around the impact on testing outcomes as well the added work of solving for duplicates by changing logic.

  **2. How do these pros and cons apply to refactoring the original VBA script?**
Ultimately, defactoring code allows a user to improve and update the code without impacting the software's functionality. As such, paying close attention to the defactoring code from the beginning will allow the user to continue to maintain clean and well organized data.

#### Additional Resources
[VBA Challenge](https://github.com/molivajimenez22/Stock-Analysis/blob/main/VBA_Challenge.xlsm)
