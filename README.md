# Stock-Analysis in VBA
## Overview of Project

To refactor existing VBA code in order to loop through and collect analysis information of the provided stock market dataset for 2017 and 2018, to determine whether refactoring the code made the VBA script run faster, and to present a written analysis explaining the findings.

## Results

### Comparison of Stock Perfomrance Between 2017 and 2018

Analysis performed by the refactored script indicates an overall drop in stock performance from 2017 to 2018. 

In 2017, every stock except for one had an increase in return:

![image](https://user-images.githubusercontent.com/5934390/111081963-31511100-84dc-11eb-99f3-7835f04419d8.png)

Whereas in 2018, every stock except for two decreased in return:

![image](https://user-images.githubusercontent.com/5934390/111081979-4b8aef00-84dc-11eb-896e-1e5616b8c71e.png)

### Comparison of Execution Time Between Original and Refactored Scripts

The refactored script executes much more quickly:

#### Original Script:
![image](https://user-images.githubusercontent.com/5934390/111082448-bb9a7480-84de-11eb-9832-727e6a62b739.png) ![image](https://user-images.githubusercontent.com/5934390/111082454-c7863680-84de-11eb-8d33-8542bf65c695.png)

#### Refactored Script:
![image](https://user-images.githubusercontent.com/5934390/111082273-c7396b80-84dd-11eb-8f50-fcb69ae378c0.png) ![image](https://user-images.githubusercontent.com/5934390/111082283-d0c2d380-84dd-11eb-8a0b-fcb60265e575.png)

## Summary
#### What are the advantages and disadvantages of refactoring code?

##### Advantages
* Can help make code easier to understand
* Can result in the program executing faster
* Can help with debugging

##### Disadvantages
* Can be time consuming when on deadline
* Can create new bugs
* Can break working code that is difficult to understand and poorly structured

#### How do these pros and cons apply to refactoring the original VBA script?
Refactoring the original script resulted in code that is easier to understand and a faster execution time, as evidenced by the runtime screenshots and the logical and commented structure of the script. It also resulted in the con of more time being invested in development in order to achieve these pros.
