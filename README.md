# An Analysis of Various Stocks with Excel
In this project, analysis was performed on various alternative energy stocks for Steve's parents.

## Overview of Project
Steve just graduated from college in finance and his parents are his first clients. A workbook was put together to analyze an entire dataset of stocks at the click of a button which was very helpful. To further his research, the initial dataset was expanded to include the entire stock market over the last few years. In this project, the initial workbook's code was refactored to run analysis on a wider data set more efficiently.

### Purpose
The purpose of this analysis is to refactor initial solution code to loop through all the data one time in order to collect the same information that was collected with the original code. Once refactored, a determination is made whether or not refactoring the code succesffully made the VBA scrip run faster. This written analysis explains the findings and results of stock performance from 2017 and 2018.

## Results and Summary
Analysis was done to include the entire stock market over the last few years

### Results
Using images and examples of your code, compare the stock performance between 2017 and 2018, as well as the execution times of the original script and the refactored script.

<p align="center">
  Alternative Energy Stock Performance in 2017 and 2018
 <br>
  <img src="https://github.com/smyoung88/stock-analysis/blob/main/Resources/All_Stocks_Performance_2017.png" title="2017 Stock Analysis">
  <img src="https://github.com/smyoung88/stock-analysis/blob/main/Resources/All_Stocks_Performance_2018.png" title="2018 Stock Analysis">
</p>

<br>

<p align="center">
  Script Runtime on 2017 Analysis for Original & Refactored Code
 <br>
  <img width="300" height="150" src="https://github.com/smyoung88/stock-analysis/blob/main/Resources/VBA_Challenge_2017_Original.png" title="Original Code">
  <img width="300" height="150" src="https://github.com/smyoung88/stock-analysis/blob/main/Resources/VBA_Challenge_2017.png" title="Refactored Code">
</p>

<br>

<p align="center">
  Script Runtime on 2018 Analysis for Original & Refactored Code
 <br>
  <img width="300" height="150" src="https://github.com/smyoung88/stock-analysis/blob/main/Resources/VBA_Challenge_2018_Original.png" title="Original Code">
  <img width="300" height="150" src="https://github.com/smyoung88/stock-analysis/blob/main/Resources/VBA_Challenge_2018.png" title="Refactored Code">
</p>

### Summary

**What are the advantages or disadvantages of refactoring code?**

**_Advantages_**
1. One major advantage of refactoring code is that it makes code more efficient by taking fewer steps to accomplish the same tasks which eliminates a lot of redundancy.
2. Refactoring uses less memory than original code (usually) and will be beneficial to companies paying for storage or computing power.
3. It improves the logic of the code making it easier for future users to read which makes everyone's life easier in the long run.

**_Disadvantages_**
1. For an unexperienced programmer, it may take longer to refactor code than someone his is more familiar and experienced. For smaller datasets, the additional time might not be beneficial initially but may pay out later on if the dataset expands.
2. Whitespace organization becomes so much more important with complex code. With more interdependancies, an unorganized or lazy programmer may get lost in the weeds with everyone going on if not organized well.

**How do these pros and cons apply to refactoring the original VBA script**
1. As seen from the results section above, the pro of refactoring the code made the analysis xxx% more efficient in the 2017 and xxx% in the 2018 dataset.
2. When a code becomes more efficient, it will inevitably use less memory which was the case for this refactoring project
3. Although only 2 years were analyzed with the new refactored code, if Steve's parents were interested in a larger lookback study, it would only have to be added as a new sheet and formatted the same as the other years for the code to generate analysis on that given year. This code allows the user to focus on what matters most: drawing conclusions from the analyzed data.
4. As an unexperienced programmer, I had to wrestle through refactoring the code and it took me longer to refactor the original code than it took to create it. In this case, it serves as a pro as well as a con because moving forward, skills were sharpened to accomplished bigger challenges in datasets in future projects.
5. When the refactored code was not working originally, I noticed I let the whitespace get a bit unorganized with proper indentation. Once cleaned up, it allowed obvious mistakes to jump out and be fixed. Organized Whitespace is crutial! 

