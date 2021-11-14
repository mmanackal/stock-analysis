# Refactoring Green Stock VBA Code

## Project Overview
Steve, a client, needed 12 stocks anlayzed for his parents to determine profitability. The analysis was completed by using VBA code in excel to review total daily volume against annual return. Steve loved the analysis what was prepared and asked to expand the analysis to include the entire stock market over the last few years.
### Purpose
Expanding the analysis beyond the 12 stocks for Steve's parents required a more efficient coding that would loop through a huge volume of stocks. In order to do this, the code required refactoring. A timer was used to compare between the original analysis and the refactored analysis to determine which was faster/efficient.
### Results
To make the code more efficient and able to go through larger data sets, I had to create 3 new arrays and loops for the analysis. Below are screenshots of the original code and the the refactored code:
#### Original Code 
![Original_Code](https://github.com/mmanackal/stock-analysis/blob/main/Resources/Original_Code.png)
#### Refactored Codes
![Refactored Code](https://github.com/mmanackal/stock-analysis/blob/main/Resources/Refactored_Code.png)
Below are the screenshots showing the runtime for the original code and the refactored code. The refactored code ran significantly faster than the original code and rendered the same results. This indicates its running faster and more efficiently.

#### Original Runtimes
![Runtime Original](https://github.com/mmanackal/stock-analysis/blob/main/Resources/2017_Original.png)
![Runtime_Original](https://github.com/mmanackal/stock-analysis/blob/main/Resources/2018_Original.png)

#### Refactored Runtimes
![Runtime Refactored]

### Summary
Refactoring the code made it more efficient and usable for a larger data set. The challenge with refactoring is ensuring that it is done correctly because you are taking something that is working and changing it which could lead to undesirable results. There were error messages received in this process before the code ran correctly. 
