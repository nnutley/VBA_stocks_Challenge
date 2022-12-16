# **VBA Stocks Challenge**

## *Overview of Project:*
    -Steve,a recent finance graduate, was interested in analyzing several green energy stocks to help his parents make an informed decision on which ones they should invest in. In order to help Steve, we used VBA to compare the outcomes of a number of different green energy stocks across two years.
    
### Purpose:
    -The purpose of this analysis was to determine the total daily volume and yearly return of each of the green energy stocks, in the years 2017 and 2018, to determine the change in demand and growth in value over that time period, for each individual stock. Additionally, we would like the analysis of these stocks to be run as quickly as possible, so the execution time of the script will also be examined.
    
## *Results*

### Stock Performance:
    -From the analysis of the green energy stocks, we can see that one stock, RUN, had the most promising performance. This is because between the years 2017 and 2018, this stock had an increase in return and an increase in volume. The increase in return indicates that the stock price was increasing, and while paired with an increase in volume, this also indicates that many people were likely buying this stock, showing that it is likely in an uptrend. There were four stocks that had expecially bad performances, DQ, HASI, SEDG, VSLR. This is beccause between the years 2017 and 2018, these stocks had a decrease in return and an increase in volume. The decrease in return indicates that the stock price was decreasing, and while paired with an increase in volume, this also indicates that many people were selling this stock, showing that is is likely in a downtrend.

### Execution of Code Comparison:
    -Overall, the refactored script's execution time was shorter than the original script. For the original script, for both the 2017 and 2018 data sets, the execution time was about 0.28 seconds. For the refactored script, the 2017 data set was about 0.06 seconds and the 2018 data set was about 0.07 seconds. 
    
    
 ![2017 Execution Time](https://github.com/nnutley/VBA_stocks_Challenge/blob/main/Resources/VBA_Challenge_2017.png)
    
    
    
 ![2018 Execution Time](https://github.com/nnutley/VBA_stocks_Challenge/blob/main/Resources/VBA_Challenge_2018.png)
    
     
    -One reason that the refactored script ran faster was because there only 3 references to a different worksheets rather than the 4 references in the original script. Eliminating the one reference saved time as the script did not have to analzye that worksheets that additional time[^1]. 
    
    -Another reason that the refactored script ran faster was because in this script we made a tickerIndex variable (tickerIndex=0) in the beginning of the script and used this in our conditionals rather than a string. For example in the original script, when looking for the total volume for the current ticker the conditional was: 
    
                If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
        
                    startingPrice = Cells(j, 6).Value
            
                End If
                
        However, in the refactored script, when looking for the total volume of the current ticker, the conditional was:
        
        
                If Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
        
                    tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
            
                End If
                
        Since evaluating strings takes longer[^1], eliminating multiple references to strings within our conditionals decreased the execution time of our refactored script.

## *Summary:*
    -From this challenge, we are able to see how refactoring a code can impact the overall performance of the code. One advantage of refactoring code is that you are often able to make your code run more quickly, which can be benefical in longer and more complicated codes. Another advantage of refactoring code is that you can decrease the repetitve scripts of code which more the code more efficient and shorter. A disadvantage of refactoring code is that the line of code can become more complex which may make it more difficult to debug a line that is not working as intended. These advantages and disadvantages are reflected in the pros and cons of the original and refactored VBA scripts from this challenge.
    
        -The pro of the original VBA script is that the individuals lines of code are less complex so it is easier to follow what the code was doing and debug it. The con of the original code is that it was longer and slower than the refactored script.
    
        -The pro of the refactored VBA script is that the changes mentioned under the results section allowed it to run faster than the original VBA script. The con of the refactored VBA script is that debug the mistakes in the code was more difficult since there were more arrays that the code was running through.
    
    
  [^1]: [Excel VBA Speed and Efficiency](https://www.soa.org/news-and-publications/newsletters/compact/2012/january/com-2012-iss42/excel-vba-speed-and-efficiency/)
