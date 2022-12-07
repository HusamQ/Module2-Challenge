# Stocks Analysis using VBA

## Overview of Project
To do analysis and research about best stock to invest in. This analysis aims to expand the stocks dataset to include the entire stock market over the last few years. Using Excel along with VBA code we will calculate and analyze each ticker and calculate the total daily volume for each stock, opening price and closing price and then automate the analysis to make it easier for the team to look and make a decision on their investment.

##Results 
# Stock analysis for 2017 measuring the execution time:
![alt text](https://github.com/HusamQ/Module2-Challenge/blob/72e526c5d9f310ae0b4894dc5be22a3f15c6c7b8/Resources/VBA_OriginalScript_2017.PNG)

#Same Stock analysis for 2017 after refactoring the code for a faster execution time:
![alt text](https://github.com/HusamQ/Module2-Challenge/blob/72e526c5d9f310ae0b4894dc5be22a3f15c6c7b8/Resources/VBA_Challenge_2017.PNG)

# Stock analysis for 2018 measuring the execution time:
![alt text](https://github.com/HusamQ/Module2-Challenge/blob/72e526c5d9f310ae0b4894dc5be22a3f15c6c7b8/Resources/VBA_OriginalScript_2018.PNG)

#Same Stock analysis for 2018 after refactoring the code for a faster execution time:
![alt text](https://github.com/HusamQ/Module2-Challenge/blob/72e526c5d9f310ae0b4894dc5be22a3f15c6c7b8/Resources/VBA_Challenge_2018.PNG)

##Summary:

We notice that after editing the algorithm to analyze the stocks for both 2017 and 2018, it has enhanced the execution time. Making faster and more efficient comparing to the module’s code. Considering this data sample is containing 3013 rows we can see the difference in the execution time which for 2017 it changed from 0.664 to 0,0938 seconds. For 2018 it changed from 0.659 to 0.101 seconds. This will definitely a major improvement in the runtime especially for a bigger dataset.
The differences between the two codes is refactored code is using arrays rather than a variable to store the output of the nested if statements. Output arrays: tickerVolumes, tickerStartingPrices, and tickerEndingPrices are very efficient way to store and access the data in continuous memory allocation pattern. Variables, on the other hand, can store a single data at a time. The difference would be tangible when we use loops and performed arithmetic operation and how many times we need to access the memory when variables are used. So refactoring the code was great idea and very helpful in terms of execution runtime. Also, we can use it on a different and bigger data samples.
The disadvantages of refactoring the code for this challenge are, time consuming, in terms of renaming variables and changing methods. Another disadvantage would be introducing bugs that will take long time to debug and reassess the code.

