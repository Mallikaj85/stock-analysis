# stock-analysis

## VBA of Wall Street

### 1. Overview of Project: Purpose of this analysis
The Challenge involves the “refactoring” or editing the existing code to make it more efficient. Refactoring is intended to improve the design, structure, and/or implementation of the existing VBA code while preserving its functionality. (Source: Wikipedia)


We have already prepared the workbook for Steve, in which he can access and analyse entire data set by using a VBA button. Though code went through several iterations and have significant redundancy and additional text. The code is unstructured, thus, a nontechnical person like Steve cannot understand the process involved in Stock Market Data Set Analysis.



Thus, refactoring is necessary and have potential advantages such as: improved code readability, reduced complexity and improved performance. In current analysis we are refactoring the code to improve the source code's maintainability and create a simpler, cleaner, or more expressive internal architecture or object model to improve extensibility and code performs faster or use less memory.

### 2. Results

**Refactored Code for All Stocks Analysis**


![image](https://user-images.githubusercontent.com/86085614/125226038-4da06700-e29e-11eb-86a6-6ad27582751c.png)
![image](https://user-images.githubusercontent.com/86085614/125226047-5002c100-e29e-11eb-8f8f-478ac18ef9bf.png)
![image](https://user-images.githubusercontent.com/86085614/125226062-58f39280-e29e-11eb-89c4-bdf277d4fe85.png)


**Methods used for refactoring the code**
1.	Extract Method: The more lines found in a method, the harder it is to figure out what the method does. This is the main reason for this refactoring. 
2.	Extract Variable: The main reason for extracting variables is to make a complex expression more understandable, by dividing it into its intermediate parts. These could be: Condition of the if() operator or a part of the ?: operator in C-based languages, A long arithmetic expression without intermediate results, Long multipart lines.
3.	Adding Comments Efficiently: Adding Comments is required, as a Best Practices for Writing a Readable Code such as Commenting & Documentation, Consistent Indentation, Avoid Obvious Comments, Code Grouping, Consistent Naming Scheme, DRY (Don't Repeat Yourself) Principle, Avoid Deep Nesting, and Limit Line Length.

**Stock performance between 2017 and 2018**


![Stock Performance 2017](https://user-images.githubusercontent.com/86085614/125226083-64df5480-e29e-11eb-8e74-523ccac0fb40.png)
![Stock Performance 2018](https://user-images.githubusercontent.com/86085614/125226088-67da4500-e29e-11eb-8191-a740498c4efa.png)


After adding Buttons for analysis and clearing the worksheet.

![image](https://user-images.githubusercontent.com/86085614/125227281-9c4f0080-e2a0-11eb-85dc-b8b9616e2f54.png)

![image](https://user-images.githubusercontent.com/86085614/125227116-4712ef00-e2a0-11eb-87cf-61a5d85b4f0c.png)

![image](https://user-images.githubusercontent.com/86085614/125227144-55f9a180-e2a0-11eb-8fd0-bbf025c93899.png)





**Execution times of the original script and the refactored script**
![Execution Time before and after Refactoring the code](https://user-images.githubusercontent.com/86085614/125226024-46795900-e29e-11eb-8258-0add42289550.png)


### 3. Summary

**Advantages of refactoring code**

•	More readable code! By giving the new method a name that describes the method’s purpose: tickerStartingPrices (), tickerVolumes (), etc.

•	Less code duplication. Often the code that is found in a method can be reused in other places in your program. Thus, the duplicates can be replaced with calls to your new method.

•	Extract method particularly isolates independent parts of code, meaning that errors are less likely (such as if the wrong variable is modified).

•	Extracted variables good names that announce the variable’s purpose loud and clear. More readability, fewer long-winded comments.

**Disadvantages of refactoring code**

•	Code refactoring is time-consuming. It takes anywhere between 50-150 hours to update the technology stack for small projects. Big projects requiring major changes on the backend and the frontend may easily take over 500 hours.

•	If you need to make some big changes to the system and modify the system’s structure, it’s easier to build new software from scratch.



**How do these pros and cons apply to refactoring the original VBA script?**

For Refactoring of VBA Scripting, the main pros that can be applied are : 

•	cleaner code, 

•	less execution time for code to run, 

•	more easier to understand, 

•	standardization of code, 

•	fewer space and duplicated comments.

The cons that can be applied for Refactoring of VBA Scripting are: 

•	reinventing the code which take same or more amount of time of writing the original code, 

•	restructuring sometime lose the important comments that can be used for readers to understand the code easily

•	Refactoring loses the preliminary thought process and essence on which code has been initially developed

