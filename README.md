# Stocks Analysis for Steve

## Overview of Project: 

#### Our good friend, Steve, just graduated with a finance degree and his parents have become his first client. Steve's parents are passionate about green energy any have invested in one green energy stock, DAQO New Energy Corp, despite theier minimal knowledge about this subject.  The purpose of this project was to analyze the stocks data for Steve and to help him determine if it makes more sense for him to encourage his parents to diversify their funds in other stocks besides DAQO. In this project I used VBA to help automate some tasks within excel to help Steve better analyze these stocks. The work that we did in VBA allows Steve to reuse it with any stock and reduces the risk of error. In addition to helping Steve by building out a way to analyze stocks within excel using VBA, I also took a final step by refactoring the code in order to run the VBA script more efficiently and quicker.

## Results:

![2017 All Stocks](https://github.com/matthubb17/stocks-analysis/blob/main/Resources/2017%20All%20Stocks.png?raw=true) ![2017 All Stocks](https://github.com/matthubb17/stocks-analysis/blob/main/Resources/2018%20All%20Stocks.png?raw=true)

##### You can see above that my stocks analysis for both 2017 and 2018 show the volatility of these green energy stocks. The imagas above illustrate 12 individual stocks performances for both 2017 and 2018. When looking at the data more closely there does not seem to be a direct relationship between the "Total Daily Volume" of each stock and the "Return" that we see for each year. This is something that Steve should notify his parents about as this seemed to be something that they believed was correlated.

![VBA Script Refactored Example](https://github.com/matthubb17/stocks-analysis/blob/main/Resources/VBA%20Script%20Refactored.png?raw=true)

##### I have proveded my refactored code above.

- Through this code I was able to use both loops and nested loops along with conditional statements in order to have VBA run this script.





## Summary:

- What are the advantages or disadvantages of refactoring code?
- How do these pros and cons apply to refactoring the original VBA script?
  - The advantage of refactoring the code was pretty clear. In working to clean up the code we were able to significantly speed up the time it takes to run the code for each year. As you can see below it took 0.085 seconds to run the 2017 cold and 0.082 seconds to run the 2018 code. This was much faster than the original code that we had as those run times were over 1 second.
  - In addition to speeding up the run time of the code, we were also able to make the code easier to understand and more condensed so that edits are easier to make in the future.
  - One con that I saw from refactoring the code is it is very time consuming. While the new refactored code is more efficient and easier to read, sometimes it may be worth considering "If it ain't broke, dont fix it". I ran into many issues and errors while refactoring the code during this project.

![2017](https://github.com/matthubb17/stocks-analysis/blob/main/Resources/2017%20Run%20Script.png?raw=true) ![2018](https://github.com/matthubb17/stocks-analysis/blob/main/Resources/2018%20Run%20Script.png?raw=true)


