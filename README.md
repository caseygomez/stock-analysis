Green Stock Analysis

This project utilizes VBA to quickly perform data analysis on Green Stock data for 2017 and 2018. 


Looking at the Green Stock data from 2017 and 2018 there are two stocks in particular that performed well. Based on the data I suggest Steve invests in ENPH and RUN. ENPH had a return of 129.52% in 2017 and 81.9% in 2018, both years are high performing investments. RUN had a return of 5.55% in 2017 and 84.0% in 2018 which shows positive growth in the market. Other Green Stocks performed well in 2017 but did poorly in 2018 which is why I would not suggest them at this time. 

<img width="264" alt="Stock_Performance_2017" src="https://user-images.githubusercontent.com/110272205/191615564-913352c5-f300-46f2-87a2-50c8d08dcfad.png">

<img width="266" alt="Stock_Performance_2018" src="https://user-images.githubusercontent.com/110272205/191615573-d22bdb89-a0b8-418b-9ca4-bf88cc0b0a64.png">

Utilizing the refactored script I was able to cut run time from .597 seconds to .171 for 2017. Similarly, run time in 2018 was reduced from .589 to .167. The original run times were more than 3x slower than the refactored scripts. This would be a significant time saver for frequently used worksheets with large data sets. The refactored script utilizes arrays to store data to cut down on run time. 

<img width="257" alt="VBA_Challenge_Original_2017" src="https://user-images.githubusercontent.com/110272205/191616992-6a48d791-44df-4eb3-8f58-0523e79a7736.png">

![VBA_Challenge_2017](https://user-images.githubusercontent.com/110272205/191617000-78fd5d6d-4349-4371-b0ac-90f1d5c8021b.png)

<img width="253" alt="VBA_Challenge_Original_2018" src="https://user-images.githubusercontent.com/110272205/191617007-680a922a-1851-4251-88d1-91a8f39b6925.png">

![VBA_Challenge_2018](https://user-images.githubusercontent.com/110272205/191617029-2fae3dd8-76c7-492a-8134-d413d33cf26c.png)


Summary: There is an advantage to refactoring code, it will save run time when working with a large data set in a frequently used worksheet. The main disadvantage with refactoring is the amount of time it takes to update code, especially when the previous code was working well. With a small data set that does not have a "long" run time it might not be worth the effort to refactor. 

Looking at the screenshots from both the original and the refactored run times you can see that the run time was significantly faster with the refactored code. The con was that it took me several hours to get the data to print as directed. 
