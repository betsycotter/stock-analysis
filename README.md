# Stock Analysis
Stock analysis exercises and results. 
## Overview of Project
 
### Purpose
 
Steve, a good friend, needs help analyzing some data for his clients, who also happen to be his parents. Steve’s parents are looking to invest in green energy, specifically DAQO New Energy Corporation. However, Steve is looking to analyze that company’s stock activity as well as other green energy companies, so that he can provide an informed recommendation to his parents. Steve has provided a file with green energy stock data for 2017 and 2018 and has asked that we help him analyze this information. We felt that this analysis would be best achieved by using VBA in Excel. We have been given a VBA script that we need to refactor for this analysis. 
 
## Results
 
As previously mentioned, Steve’s parents were interested in DAQO (Ticker: DQ) for their investment. We used our refactored code to calculate the total daily volume and the annual return. The first calculation was to take the daily volume and add it to each subsequent daily volume for that particular ticker. See code below: 

 <img width="544" alt="Screen Shot 2022-11-02 at 4 22 16 PM" src="https://user-images.githubusercontent.com/116031639/199613911-06c71a22-9ec0-4735-94f3-e18c69cdcc5a.png">
 
Then, to calculate the annual return, we found the starting and ending prices for the year. We took those values and divided the ending price by the starting price and subtracted 1 from that number. This value has been formatted to be a percentage, as it is easier to read. See calculations codes below: 

<img width="677" alt="Screen Shot 2022-11-02 at 4 23 06 PM" src="https://user-images.githubusercontent.com/116031639/199613996-44208192-f96e-45f1-b97a-e7015dd30324.png">

Formatting code is as follows: 

<img width="458" alt="Screen Shot 2022-11-02 at 4 24 23 PM" src="https://user-images.githubusercontent.com/116031639/199614098-d964a669-0ca4-40cb-92ba-b0ceea32137b.png">
 
After reviewing the data for 2017 and 2018, it would not be the best recommendation for them to invest in DQ. In 2017, DQ had an almost 200% return, yet in 2018 their return was -62.6%. The volatility of this stock, especially with a downward trend is not reassuring for future success. The only two stocks that had positive returns in 2018 were ENPH and RUN. These two stocks also had positive returns in 2017, however ENPH has a lower return value in 2018 than 2017. See results below: 

2017 Results:

<img width="253" alt="2017 Results" src="https://user-images.githubusercontent.com/116031639/199614202-1d06f009-088c-48dc-87e5-5f17217e36d0.png">

2018 Results:

 <img width="254" alt="2018 Results" src="https://user-images.githubusercontent.com/116031639/199614215-0e94eac0-b3ec-4015-a4e1-0e6fe3d85f24.png">
 
## Summary
 
### What are the advantages or disadvantages of refactoring code? 
 
#### Advantages: 
1. Ideally, you can make the macro run faster, if you successfully refactor the code. 
2. DRY theory - by refactoring code, you can utilize code that has already been written and executed successfully. 
3. Time saving - if the code was written well originally, this can save time by copy/pasting the original code. 
4. Learning new code styles and functions - someone might use a function or code style that was previously unknown to you. With refactoring, you can read and understand the code from someone else. 
5. Learning about the data structures - if you are working at a company where you are unfamiliar with the data structures, you can see the data sources and common code patterns of more tenured employees. 
 
#### Disadvantages: 
1. It might be hard to understand - if someone doesn’t include comments or uses overly complex codes, it would take time to digest this code and understand each part. 
2. It could take longer to run - you could make the code more complicated and therefore take a longer amount of time to run the macro. 
3. New mistakes could occur - without a proper QA of the code, the refactor could add additional errors/mistakes. 
 
### How do these pros and cons apply to refactoring original VBA script? 
Overall, the main reasons to refactor code include time, learning, and resources. Time can be viewed as both a pro and a con, because ultimately refactoring could save time, yet the time it takes to refactor, could be a detriment to deploying a project quickly. Also, learning is a pro and a con because the person doing the refactor could learn good or bad codes from the original code. Learning also comes into play because if someone is new to VBA codes, doing a refactor might not teach them how to code better, just refactor someone else’s work. Finally resources can be a pro and a con, because similar to time, the code could improve memory space, but might also take other resources away. 

In this particular refactor, we were able to improve on the run time for the script. In the original file, the elapsed time was about 0.5 seconds, see below for the new results: 
 
2017: 

 <img width="263" alt="VBA_Challenge_2017" src="https://user-images.githubusercontent.com/116031639/199614425-2d0ff7bd-e6e1-440a-881e-6726047184d2.png">

2018: 

<img width="263" alt="VBA_Challenge_2018" src="https://user-images.githubusercontent.com/116031639/199614439-67465491-89d6-4c64-b310-bc1fe040be45.png">
