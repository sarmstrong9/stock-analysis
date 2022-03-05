# Stock Analysis

## Overview of Project
The scope of this assignment was to build a stock analysis tool using VBA that could be user-friendly enough to be passed along to a non-developer to compare the performance of multiple stocks over different years.  Furthermore, the initial code was then refactored to run the analysis much more efficiently and quicker.  This would allow the code to perform quickly and be run with much larger datasets without issue.

##Results
The impact of refactoring the code was significant.  Even with the relatively small dataset that this code is using, the refactored code took roughly 10% of the time of the original code.  This was achieved by populating arrays of data and having a single "For" loop run through all the data, as opposed to having a nested "For" loop that would have to run through the entire dataset for each of the 12 stock tickers.

###The differences for the year 2017 analysis:
Original Code for 2017:
![Original Code for 2017](/Resources/VBA_Challenge_2017_Original_Code.png)

Refactored Code for 2017:
![Refactored Code for 2017](/Resources/VBA_Challenge_2017.png)

###The differences for the year 2018 analysis:
Original Code for 2018:
![Original Code for 2018](/Resources/VBA_Challenge_2018_Original_Code.png)

Refactored Code for 2018:
![Refactored Code for 2018](/Resources/VBA_Challenge_2018.png)





##Summary

###Advantages of Refactoring Code

###Pros/Cons of Refactoring the Original VBA Script