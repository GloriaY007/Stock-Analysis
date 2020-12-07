# Stock-Analysis
Using VBA to write macros for stock analysis

## Overview of Project
The purpose of this project is for Steve to help his parents analyze the stock market, more specifically the performance of the DAQO (Ticker:DQ) stock, for all the years of available data. The aim is to make sure that the final worksheet expand the dataset to include the entire stock market over the last few years, produces data using a few buttons for the user, and run in an optimal fashion using refactoring.

## Results
The analysis is well described with screenshots and code (4 pt).

Based on the Macro Code that were built in the back ground (in a previous exercice), I was able to run a code in **0.63 seconds**. The macro provided us with the following information for the year 2017:
- Most stocks performed well during that year. 
- The **DQ** stock performed extremely well, out of the 12 stocks we observed, with **199.45%** of return, followed by the **SEDG** stock with **184.47%** return. 
- For that year, there was only one stock which dropped in return. The TERP stock had a **Total Daily Volume** of **$139,402,800** for the year and had a **(7.21)%** drop in return.

     | Ticker  |Total Daily volume | Return |
     | ------------- | ------------- | ------------- |
     | DQ  | $35,796,200  | 199.45%  |
     | SEDG  | $206,896,200 | 184.47%  |
     | TERP  | $139,402,800  | (7.21)% |

  *This table highlights the above mentioned stocks
  
  
  
  ![2017 Return](https://github.com/GloriaY007/Stock-Analysis/blob/main/2017%20Return.PNG)
  
  *This image captures the full list of data for the 12 selected stocks, for the Year 2017
  
  
  ![Displaying the Elapsed Time for 2017 - Green Stock Analysis](https://github.com/GloriaY007/Stock-Analysis/blob/main/Displaying%20the%20Elapsed%20Time%20for%202017%20-%20Green%20Stock%20Analysis.PNG)
  
  *This image captures the elapsed time to run the Macro for the Year 2017
  

The same Macro Code that were built in the back ground (in a previous exercice), were also ran, in **0.62 seconds**, to get information for the year 2018:
- Most stocks performed poorly in 2018.
- The only 2 stocks that performed well were the **ENPH** tickered stock with a **Total Daily Volume** of **$607,473,500** and	**81.92%**, as well as **RUN** with 	**$502,757,100** and **83.95%** return.
- The **DQ** stock saw a drop of **62.60%** in return
    
     | Ticker  |Total Daily volume | Return |
     | ------------- | ------------- | ------------- |
     | DQ  | $107,873,900  | (62.60)%  |
     | ENPH | $607,473,500| 81.92%  |
     | RUN  | $502,757,100  | 83.95% |
     
  *This table highlights the above mentioned stocks
  
  
    ![2018 Return](https://github.com/GloriaY007/Stock-Analysis/blob/main/2018%20Return.PNG)
  
  *This image captures the full list of data for the 12 selected stocks, for the Year 2018
  
  
    ![Displaying the Elapsed Time for 2018 - Green Stock Analysis](https://github.com/GloriaY007/Stock-Analysis/blob/main/Displaying%20the%20Elapsed%20Time%20for%202018%20-%20Green%20Stock%20Analysis.PNG)
  
    *This image captures the full list of data for the 12 selected stocks, for the Year 2018
    
  
  The other portion of this analysis of stocks for Steve parents, was to automate the worksheet in a way that for any year and each stock (ticker), someone was able to run the macro without much effort or changes and in the least amount of time.

## Summary
All source documents for this analysis can be found in [Git Hub](https://github.com/GloriaY007/Stock-Analysis.git)

Due to the below syntax error, this particular specification could not be run.

  ![Error Message - Green Stock AnalysisRefactoredStart.PNG](https://github.com/GloriaY007/Stock-Analysis/blob/main/Displaying%20the%20Elapsed%20Time%20Error%20-%20Green%20Stock%20AnalysisRefactoredStart.PNG)

However, we can still deduce some answers to our analysis. In theory, refactoring a code has for purpose to render a code more efficient, this also should reduce the run time of a code to make it faster. In this particular case, we were only able to time the results of a specific subroutine titled 'DQAnalysis2'. The clear advantage of refactoring code a code is to make the code neater and faster, however it is at the cost of a lot of **time**. 

For instance, in this analysis, teh original code that was created was completed and ran withouth too much issues, and because the author is one and the same, it was easy to work along and fix issues on the fly. On the other hand, refactoring a code that felt foreign required more attention, identification of the patterns, corrections, and ultimately, in my case, had errors that were harder to correct.

Overall, using this exercise, Steve should be able to dissuade his parents from ivesting in a DQ stock, but to **look at the ENPH stock as an investment opportunity!**
