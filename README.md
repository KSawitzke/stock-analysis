# stock-analysis

# Stock Analysis


## Overview of Project
  The purpose of this project was to help our friend Steve create a more efficient macro for stock analysis that he could use long term. He was pleased with the workbook we collaborated on to help his parents choose between green energy stocks, and wanted to refactor that existing code to be usable for future applications. By refactoring this code, we aimed to create a broader workbook that can be used quickly and efficiently on potentially huge numbers of stocks with large underlying trading datasets.

## Analysis

  We were able to help Steve and by extention his parents make an educated decision on their green energy stocks, even using our initial macro before refactoring it. There were huge gains seen in 2017, followed by a large drop off in 2018 with returns falling for all but two of the 12 stocks chosen to analyze. Here we can see the results across the two years of stock data based on the results of our macro compiling the data into easy-to-read tables.

#### Table 1: 2017 Results
![2017](Resources/2017_Results.PNG)

#### Table 2: 2018 Results
![2018](Resources/2018_Results.PNG)


## VBA Results
  In terms of code performance and functionality, there were great improvements, regardless of the poor performance of the stocks in question. 
  
  Firstly, we added an important feature that allowed for Steve or other end users to input the desired year that they wanted to run analysis on. Initially, the VBA code was hardcoded to use "2018" as the year, and in order to change which year of data was looked at, one would have to go in and edit the code.
  '''
      yearValue = InputBox("What year would you like to run the analysis on?")

    startTime = Timer
    
    'Format the output sheet on All Stocks Analysis worksheet
    Worksheets("All Stocks Analysis").Activate
    
    Range("A1").Value = "All Stocks (" + yearValue + ")"
    
    .
    .
    .
    
        'Activate data worksheet
    Worksheets(yearValue).Activate
    '''
    
    Here instead, we created an input box which saves the user input as a variable, and uses that to determine which worksheet to activate before running the calculation portions of the code. Other formatting such as titles is also updated accordingly. Now, this technically looks at the spreadsheet names, but in the current format they are named after years, so to an end user they function the same.



The analysis is well described with screenshots and code (4 pt).
Summary
There is a detailed statement on the advantages and disadvantages of refactoring code in general (3 pt).
There is a detailed statement on the advantages and disadvantages of the original and refactored VBA script (3 pt).
