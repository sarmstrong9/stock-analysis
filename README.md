# Stock Analysis

## Overview of Project
The scope of this assignment was to build a stock analysis tool using VBA that could be user-friendly enough to be passed along to a non-developer to compare the performance of multiple stocks over different years.  Furthermore, the initial code was then refactored to run the analysis much more efficiently and quicker.  This would allow the code to perform quickly and be run with much larger datasets without issue.

## Results
The impact of refactoring the code was significant.  Even with the relatively small dataset that this code is using, the refactored code took roughly 15-20% of the time of the original code.  This was achieved by populating arrays of data and having a single "For" loop run through all the data, as opposed to having a nested "For" loop that would have to run through the entire dataset for each of the 12 stock tickers.

### The difference in times for the years are:

2017 - Original code took 1.15 secs, while the refactored code took 0.24 secs to complete

2018 - Original code took 1.48 secs, while the refactored code took 0.18 secs to complete

#### Images of these times are posted below:

##### The differences for the year 2017 analysis:

Original Code for 2017:

![Original Code for 2017](/Resources/VBA_Challenge_2017_Original_Code.png)

Refactored Code for 2017:

![Refactored Code for 2017](/Resources/VBA_Challenge_2017.png)

##### The differences for the year 2018 analysis:

Original Code for 2018:

![Original Code for 2018](/Resources/VBA_Challenge_2018_Original_Code.png)

Refactored Code for 2018:

![Refactored Code for 2018](/Resources/VBA_Challenge_2018.png)


## Summary

### Advantages/Disadvantages of Refactoring Code
It is clear from the results that refactoring code can have a large advantage on how quickly the code can run.  In this case, the refactored code runs in 15-20% of the time compared to the original code which indicates that this method is much faster and efficient in analyzing a dataset.  The benefits will only be more noticeable the larger the dataset grows as well.  This example demonstrates that spending time and resources to refactor code can have a huge benefit if the original code is not very efficient.

A disadvantage of refactoring code could be applying time and resources to try and improve code if the original code is already fairly efficient.  In this scenario, the gains made by refactoring might not offset the time it took to improve the code and those same resources could have been spent on more productive tasks.  Refactoring code needs to be viewed as a tradeoff of gains vs effort required and a decision be made if its worth it.  There might be 1000+ different ways to obtain the same results, however a cost/benefit analysis needs to be done to determine if any additional refactoring is worth it.

### Pros/Cons of Refactoring the Original VBA Script
The pros of refactoring the original VBA script in this task was to learn a different and new method of analyzing the same dataset while obtaining the same result, and also seeing the impact of how that method can affect the time it takes for the computer to run the analysis.  This was very eye opening just how much impact the different methods have on analysis times.

The cons of refactoring the original VBA script were potentially breaking a known-good script by introducing new errors/bugs that were not there originally.  Using different methods in coding can utilize a different approach to the logic used, so it is very important to ensure the new code is still accomplishing all the goals of the original code, just in a more efficient way.  Refactoring can also take a long time to work through, either through new bugs or obtaining different results that need to be worked through.  In my case, I incorrectly referenced the "Adj Close" column instead of the "Close" column in the refactored code.  This single mistake took me about 2 additional hours to work through as I could not figure out why my results were different, even though I was not getting any errors and my logic statements were verified multiple times.  Once I referenced the correct column my code worked correctly, but this was difficult to troubleshoot since there were no errors being presented.


## Multiple Scripts with Single Button 
This was not part of the assignment, but in the earlier portion of the exercise we made a "clear sheet" script which would clear all the data from our current sheet, an "analysis" script that would analyze the required data, and a "format" script to make the analysis much easier to read.  All three scripts are useful, but I was getting tired of running all three independently to get the end result of a new formatted output each time, so I researched how to combine multiple scripts and apply them to a single button.  The below picture shows my VBA syntax to combine multiple scripts together.  Once this combined script was created, I assigned it to my button and now the button does all three without any additional clicks.

Screenshot shown here:

![Single Button w/ Multiple Scripts](/Resources/VBA_Challenge_Button_with_Multiple_Scripts.png)







