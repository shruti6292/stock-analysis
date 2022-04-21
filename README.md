# stock-analysis

## Overview of Project: 

####  In our project, we are developing a product for Steve to help his parents in stock analysis. We have written a functional code which works perfectly fine for small amount of data. But if we want to expand our results for large amount of data in stocks analysis, it takes so much time to exceute. Hence the concept of refactoring comes in.
####  Code refactoring  involves editing and cleaning up previously written software code without changing the function of the code at all.The basic purpose of code refactoring is to make the code more efficient and maintainable. 

####  So initially the Code for All stocks prices before refactoring was assigning value to each ticker and going through the given sheet everytime and complete the loop.That was repetative search and was taking more time. Also the variables were "double". This is shown below:

![alt text](https://github.com/RGK73/stock-analysis/blob/main/Resources/data%20type%20double.png)
![alt text](https://github.com/RGK73/stock-analysis/blob/main/Resources/Repetative%20loop%20through%20each%20ticker%20value.png)

Hence the time to execute the code for the years 2017 and 2018 was more as shown in the following images

![alt text](https://github.com/RGK73/stock-analysis/blob/main/Resources/All%20Stocks%20Analysis%202017%20before%20refactoring.png)
![alt text](https://github.com/RGK73/stock-analysis/blob/main/Resources/All%20Stocks%20Analysis%202018%20before%20refactoring.png)
![alt text](https://github.com/RGK73/stock-analysis/blob/main/Resources/original%202017.png)
![alt text](https://github.com/RGK73/stock-analysis/blob/main/Resources/original%202018.png)

Then we refactored the code to make it faster. We assingned the array and defined the variables as "single", did the indentation to make the code more readable and lean.
This is shown below:

![alt text](https://github.com/RGK73/stock-analysis/blob/main/Resources/data%20type%20single.png)
![alt text](https://github.com/RGK73/stock-analysis/blob/main/Resources/loop%20through%20all%20tickers%20once%20value.png)

We can immediately see the difference in the time for the execution of the code,it is a substancial change.

![alt text](https://github.com/RGK73/stock-analysis/blob/main/Resources/VBA_Challenge_2017.png)
![alt text](https://github.com/RGK73/stock-analysis/blob/main/Resources/VBA_Challenge_2018.png)

From the time stamps screenshots below, we can clearly see the difference.

![alt text](https://github.com/RGK73/stock-analysis/blob/main/Resources/refactored%202017.png)
![alt text](https://github.com/RGK73/stock-analysis/blob/main/Resources/refactored%202018.png)

## Summary: So overall we can see the What are the advantages or disadvantages of refactoring code :

### Advantages

1).The basic purpose of code refactoring is to make the code more efficient and maintainable. This is key in reducing technical cost since it’s much better to clean up the code now than pay for costly errors later. 
2).Code refactoring improves readability, makes the QA and debugging process go much more smoothly. And while it doesn’t remove bugs, it can certainly help prevent them in the future.
3).Refactoring the current code before adding in new programming or functionality will not only improve the quality of the product itself, it will make it easier for future developers to build on the original code.
4).Breaking down the refactoring process into manageable chunks and performing timely testing before moving on to other updates always results in a higher quality application and a better overall development experience.

### Disadvantages

1).Earlier we stressed that refactoring should never affect the performance of an application and that it should only serve as a clean-up effort. There are times, however, when an application needs to be completely revamped from the start. In these cases, refactoring is not necessary, as it would be much more efficient to simply start from scratch.
2)Another situation in which it would be wise to skip refactoring is if you are trying to get a product to market within a set time frame. Refactoring can be like going down the proverbial rabbit hole: Once you start, it can become quite time-consuming.
3). Adding any additional coding or testing to an already tight timeline will lead to frustration and additional cost for your client.

## How do these pros and cons apply to refactoring the original VBA script?

While refactoring you just want to make the code more efficient—by taking fewer steps, using less memory, or improving the logic of the code to make it easier for future users to read it is not always feasiable. But refactoring is key in reducing technical cost since it’s much better to clean up the code now than pay for costly errors later. Code refactoring, which improves readability, makes the QA and debugging process go much more smoothly. And while it doesn’t remove bugs, it can certainly help prevent them in the future. It's like a good housekeeping so that we can find everything quickly when in need.
