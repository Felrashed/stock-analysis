# **Stock Analysis with VBA: Visualizing Green Energy Returns**

![](https://github.com/Felrashed/stock-analysis/blob/main/Resources/VBA_Example_Gif.gif)

## **Project Background and Purpose**

Our client, Steve, is an equity analyst who requested an automated file for his parents to be able to visualize the performance of a green energy stock they have been tracking, [**$DQ**](https://www.google.com/finance/quote/DQ:NYSE). His parents are passionate about green energy and are interested in $DQ, but Steve is concerned about diversification. He has tasked us with creating an excel file that his parents can use to easily analyze, compare, and visualize the 2- year performance of an index of green energy stocks he has selected for their review. Steve believes that his parents will be better equipped to make investment decisions as a result. To achieve these goals, Steve wants us to engineer the file to run on VBA with the following functionality and features:

1. Calculate the returns for both 2017 and 2018
2. Identify the total daily volume of each stock
3. Simplify the process for performing the analysis
4. Visualize the data in a simple and easily discernible way 
5. Refactor the code to efficiently handle expanded datasets in the future

## **RESULTS**

#### **STOCK PERFORMANCE**

Steve and his parents believe that total daily volume is a good indicator for the potential performance of a stock. They think that more [**liquid**](https://www.investopedia.com/terms/v/volume.asp) stocks have greater return potential and are an indication of strong [**investor sentiment**](https://www.investopedia.com/terms/m/marketsentiment.asp). We used this metric as well as [**yearly return**](https://www.investopedia.com/terms/a/annual-return.asp) to conduct our analysis on the basket of green energy stocks Steve selected. 


**Fiscal Year 2017**             |  **Fiscal Year 2018**
:-----------------------------------------------------------------------------------------:|:-----------------------------------------------------------------------------------:
![](https://github.com/Felrashed/stock-analysis/blob/main/Resources/2017_Performance.png)  |  ![](https://github.com/Felrashed/stock-analysis/blob/main/Resources/2018_Performance.PNG)

While the results clearly indicate that green energy stocks had a great year for returns in FY2017, there needs to be a further inspection of the results relative to Steve's parents' goals. $DQ, the ticker of interest, had a blowout year returning 199.4% for FY2017. At first glance, this seems to be an excellent indicator of $DQ's potential, but when reviewed against the volume of its peers, $DQ is actually the least liquid. This suggests that while the green energy industry was performing well as a whole, $DQ's gains were possibly a result of its [**low liquidity**](https://www.investopedia.com/terms/s/shortsqueeze.asp) which can significantly affect pricing to the upside when pairing gains with its sector peers. This finding is further supported by FY2018 results where $DQ's volume significantly increases (implying an uptick in general trading activity) and its share price declined at a greater rate than it's peers. 

Over the course of FY2017 through FY2018, green energy stocks as a sector group faired very poorly. This possibly suggests 2017 was the height of a speculative boom in green energy investing. While volumes are shifting year over year for each ticker and performance for the industry is declining by measure of annualized returns, there are still emerging companies that are trading contrary to their peer group and instead are leading their sector in annualized returns. Most notably, ENPH and RUN were able to triple and double, respectively, their daily trading volume and maintain positive annualized returns through 2018, contrary to the rest of their peer group. Declines in volumes and returns elsewhere in the industry suggest a consolidation of investor capital into the two strongest performing assets in the green energy sector group, ENPH and RUN. With this in mind, Steve and his parents should consider their investment objectives as the primary factor in which company to invest in:
1. Are they seeking high returns and able to tolerate risk? If so, it may be worthwile to identify a stock that was oversold in 2018 (one with high volume and poor returns) that they feel may be undervalued. This would yield greater returns as the stock price reverts to its 200 day moving average and begins to reflect future performance potential. 
2. If they are seeking safer and more stable returns, they should back ENPH and RUN for their consistent performance and leadership within their sector. 




#### **VBA CODE PERFORMANCE**

**Original 2017**             |  **Refactored 2017**
:-------------------------:|:-------------------------:
![](https://github.com/Felrashed/stock-analysis/blob/main/Resources/Orginal_Clock_Speed_cropped_2017.PNG)  |  ![](https://github.com/Felrashed/stock-analysis/blob/main/Resources/Orginal_Clock_Speed_cropped_2018.png)

- Here we can see that the refactored 2017 code performed at a significantly faster speed than it's original counterpart

**Original 2018**             |  **Refactored 2018**
:-------------------------:|:-------------------------:
![](https://github.com/Felrashed/stock-analysis/blob/main/Resources/VBA_Challenge_2017.png.PNG)  |  ![](https://github.com/Felrashed/stock-analysis/blob/main/Resources/VBA_Challenge_2018.png.PNG)

- Similarly, the refactored 2018 code performance was significantly faster than it's original counterpart as well

Our refactored code was more effecient because we removed redundant iterations and conditionals and set variables that could be interpreted regardless of which ticker was being analyzed. By doing so, our machines require less memory and are able to more efficiently process the directions while still achieving the desired outcomes. This is important because when dealing with larger datasets, it's important to ensure your processes are able to run smoothly and efficiently enough to be worthwhile. 

**Original Looping**             |  **Refactored Looping**
:-------------------------:|:-------------------------:
![](https://github.com/Felrashed/stock-analysis/blob/main/Resources/OG_Loop_Example.PNG)  |  ![](https://github.com/Felrashed/stock-analysis/blob/main/Resources/LoopCode_Refactored_Example.PNG)



## **SUMMARY**

- Advantages of Refactoring Code:
    1. It can make a program run more efficiently, as it did in this case
    2. Can make it easier to identify and fix bugs, magic number traps, etc. 
    3. Can augment the way we approach problem solving for future projects, including at the onset of development
    4. Can improve readability and understanding for future users of your code product
    5. Can remove poor structure and motivate clean code that reduces potential for errors, misinterpretation, and confusion by future users
- Disadvantages of Refactoring Code:
    1. This can be a time consuming task and may not yield the desired results if the code is already lean enough, therefore it is situationally dependent 
    2. There is always the potential that through the refactoring process, one makes a mistake that results in further need to debug
    3. It's possible that you change the code too much, and overcomplicate a process or even stray away from its original functionality

- Pros & Cons: Applied To Our Refactored VBA Code:

    1. In this case, we were able to make the code more efficient and run significantly faster
    2. The process was time consuming, but there were no externam time constraints that prevented us from pursuing a refactor
    3. The refactoring process allowed us to make a more dynamic file, capable of allowing Steve to analyze thousands of stocks in the future more efficiently