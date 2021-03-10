# An Analysis of Various Stocks with Excel
In this project, analysis was performed on various alternative energy stocks for Steve's parents.

## Overview of Project
Steve just graduated from college in finance and his parents are his first clients. A workbook was put together to analyze an entire dataset of stocks at the click of a button which was very helpful. To further his research, the initial dataset was expanded to include the entire stock market over the last few years. In this project, the initial workbook's code was refactored to run analysis on a wider data set more efficiently.

### Purpose
The purpose of this analysis is to refactor initial solution code to loop through all the data one time in order to collect the same information that was collected with the original code. Once refactored, a determination is made whether or not refactoring the code succesffully made the VBA scrip run faster. This written analysis explains the findings and results of stock performance from 2017 and 2018.

## Results and Summary
Analysis was done to include the entire stock market over the last few years

### Results

#### Stock Performance
The alternative energy stocks performed very differently from 2017 to 2018. A side-by-side analysis of each over the two years is displayed below:

<p align="center">
  <b>Alternative Energy Stock Performance in 2017 and 2018</b>
 <br>
  <img width="400" height="402" src="https://github.com/smyoung88/stock-analysis/blob/main/Resources/All_Stocks_Performance_2017.png" title="2017 Stock Analysis">
  <img width="400" height="402" src="https://github.com/smyoung88/stock-analysis/blob/main/Resources/All_Stocks_Performance_2018.png" title="2018 Stock Analysis">
</p>

As seen from the analysis, if Steve's parents only had access to the 2017 information prior to their decision to invest, DQ would make sense for a strong investment as it outperformed all other peers and had a return of 199.4%. In 2017, all stocks but one had a positive return and an average return of 67.3% whereas in 2018, only two stocks had a positive return with an average of -8.5%. Totaly daily volumes were very similar in both year with a total of 3.17 billion in 2017 and 3.31 in 2018.
<br>

#### Original vs Refactored Script Performance
Original code was written to cycle through all of the rows of data for each respective stock ticker and returning those values before cycling to the next stock ticker. The code was then refactored to pull respective stock ticker data for each ticker with only one pass through the data. The efficiency gains from running the new refactored code vs the original for each year is shown below. Mousing-over each picture will indicate whether it is original or refactored.

<p align="center">
  <b>Execution Time of Original & Refactored Script on 2017 Analysis</b>
 <br>
  <img width="300" height="150" src="https://github.com/smyoung88/stock-analysis/blob/main/Resources/VBA_Challenge_2017_Original.png" title="Original Script">
  <img width="300" height="150" src="https://github.com/smyoung88/stock-analysis/blob/main/Resources/VBA_Challenge_2017.png" title="Refactored Script">
</p>

<br>

<p align="center">
  <b>Execution Time of Original & Refactored Script on 2018 Analysis</b>
 <br>
  <img width="300" height="150" src="https://github.com/smyoung88/stock-analysis/blob/main/Resources/VBA_Challenge_2018_Original.png" title="Original Script">
  <img width="300" height="150" src="https://github.com/smyoung88/stock-analysis/blob/main/Resources/VBA_Challenge_2018.png" title="Refactored Script">
</p>

The refactored code decreased the execution times of the 2017 and 2018 analysis by 83% and 84% respectively. As mentioned above, the original code cycles through each data set completely for one stock ticker before it loops back to gather data for the next stocks which from an operations standpoint would be considered an extremely high amount of "dead walking". Dead walking is when you have to go back to an area you already completed to accomplish another task that could have been done at the same time. The "dead walking" script is as follows:

<p align="center">
  <b>Original Script</b>
 <br>
  <img width="500" height="450" src="https://github.com/smyoung88/stock-analysis/blob/main/Resources/Original_Script.PNG" title="Original Script">
</p>

Refactoring the script allowed arrays to be setup for the data of interest so that after the starting price and ending price for a given stock ticker was cycled through the code, it would begin gathering data for the next one. Since the stock tickers are in alphabetical order, the code would only require one pass through each row in the dataset to gather all desired data. The refactored code that made this possible is below:

<p align="center">
  <b>Refactored Script</b>
 <br>
  <img width="750" height="600" src="https://github.com/smyoung88/stock-analysis/blob/main/Resources/Refactored_Script.png" title="Refactored Script">
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
5. When the refactored code was not working originally, I noticed I let the whitespace get a bit unorganized with proper indentation. Once cleaned up, it allowed obvious mistakes to jump out and be fixed. Organized whitespace is crucial! 

