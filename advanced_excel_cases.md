<h1 align="center">Peter Garay-Robles </h1>

<h3 align="center">A Data Engineer in SQL and Excel. </h3>

----


## Excel Data Pipeline & Analytics Framework

A Structured Suite of 10 Advanced Excel Functions Designed to Automate Data Cleaning, Deduplicate Records, and Summarize KPI's by Region Dynamically.


---

## 1. Excel FILTER Function

7 Use-Cases to Utilize Dynamic Segmentation Based Location, Multiple Conditions, Customized Results, Aggregation, and Approximate Matching.

- Filter by UK Region

---

![excel filter function](https://github.com/peter28data/excel-modeling/blob/7f5c1c2a28c8b0a66befcc427e74ac5bd302cf7b/images_excel_modeling/excel_filter_function.png)

- Result of FILTER Function

![excel modeling result](https://github.com/peter28data/excel-modeling/blob/7f5c1c2a28c8b0a66befcc427e74ac5bd302cf7b/images_excel_modeling/excel_filter_result.png)

- Filter by 2 Conditions

The Asterisk (*) enables the Segmentation Tool to Filter 2 Conditions such as United States Clients in the Technology Industry.

![filter and](https://github.com/peter28data/excel-modeling/blob/29ef3a7a27ac1c95c7bd5012d5aeda114aba0071/images_excel_modeling/v2/filter_excel_2conditions_and.png)

- Filter by 2 Conditions

The Addition (+) enables the Segmentation Tool to Filter for Clients that are Either from the United States Or in the Technology Industry.

![filter or](https://github.com/peter28data/excel-modeling/blob/29ef3a7a27ac1c95c7bd5012d5aeda114aba0071/images_excel_modeling/v2/filter_excel_2conditions_or.png)

- Filter with Headers

The VSTACK Function Combines the original headers with Each Record Detail to make the Results of the Filter Legible.

![filter vstack](https://github.com/peter28data/excel-modeling/blob/29ef3a7a27ac1c95c7bd5012d5aeda114aba0071/images_excel_modeling/v2/filter_excel_vstack_headers.png)

- Filter with Choosing Columns

To Concentrate on the Important Details of the Project, nested functions with the Filter tool enables the User to Choose important details such as Client Name, Industry, and Net Worth. Ignoring irrelevant data such as Indexing and Country.

![filter choosecols](https://github.com/peter28data/excel-modeling/blob/29ef3a7a27ac1c95c7bd5012d5aeda114aba0071/images_excel_modeling/v2/filter_excel_vstack_choosecols.png)

- Filter and Average Aggregation

To Isolate Clients from the United States and Aggregate Financial Data.

![filter average](https://github.com/peter28data/excel-modeling/blob/29ef3a7a27ac1c95c7bd5012d5aeda114aba0071/images_excel_modeling/v2/filter_excel_average_networth_americans.png)

- Filter Subset of Family Clients

A small group of Clients are from Families ("Bernaurd & Family"). An Approximate Match Tool such as SEARCH will provide a search bar for the User to Isolate Clients from Families. Additionally the ISNUMBER function will return True of False to combine with the Filter Function otherwise returning a Lookup error ("N/A") for clients without "Family".


![filter isnumber search](https://github.com/peter28data/excel-modeling/blob/29ef3a7a27ac1c95c7bd5012d5aeda114aba0071/images_excel_modeling/v2/filter_excel_isnumber_search.png)

---

## 2. Excel SUMIFS Function

Utilize to Implement Automated Financial Aggregation By Region and Unique Product.

- Find The Total Sales in the U.K. for Eclairs Product.

---

![Excel Sumifs](https://github.com/peter28data/excel-modeling/blob/2efbee4cca62b8a2435e08bc25195976565a8e2d/images_excel_modeling/v1/excel_sumifs_function.png)


- Dynamic Arrays in Tables

A Spill Error occurs if we select the entire column inside a table.

![wrong table dynamic array](https://github.com/peter28data/excel-modeling/blob/4249af4dbec8f8674fad188be6ea2a5974138e2d/images_excel_modeling/v3/wrong_table_dynamic_array.png)

![spill error table array](https://github.com/peter28data/excel-modeling/blob/4249af4dbec8f8674fad188be6ea2a5974138e2d/images_excel_modeling/v3/spill_error_table_array.png)

Tables have a fixed dynamic array in the Table, therefore, to overcome the Spill Error we select only one cell. 

![correct table dynamic array](https://github.com/peter28data/excel-modeling/blob/4249af4dbec8f8674fad188be6ea2a5974138e2d/images_excel_modeling/v3/correct_table_dynamic_array.png)

To use a Function on the Table
1. Convert Table to Dynamic Array with Filter
2. Use Hash Operator for Dynamic Insights

---

 ## 3. Excel COUNTIFS Function

Utilize to Segment Sales Record to Isolate Salesperson.

- How many of those sales came from the Sales Person "Oby Sorrel"?

![Excel COUNTIF Function](https://github.com/peter28data/excel-modeling/blob/2efbee4cca62b8a2435e08bc25195976565a8e2d/images_excel_modeling/v1/excel_countifs_function.png)

## 3a. Excel COUNT Function

Returns Frequency of Records Only if Containing Numerical Data. This can Be Useful if There are cells with empty spaces or other characters (_, N/A, etc).

## 3b. Excel Nested Count Function if Unique

=COUNTA(UNIQUE(A2:A5))

## 3c. Excel COUNTA Function

Utilize to Return the Frequency of Array if not empty. Useful to compare how many cells are empty.

In this Example, We use the Unique Function to List the Distinct Subjects from Student Test Records. Then we use the COUNTA Function to Calculate the Frequency of Distinct Subjects.

However, if the Student Records are Updated and a New Subject is introduced, the Unique Function will display it as long as it is within the range But the Count Function will not Calculate the Frequency because it is outside the range, as shown below.

![counta without](https://github.com/peter28data/excel-modeling/blob/138127bb44dd9184f5803cc725516b9e005e8feb/images_excel_modeling/v2/counta_without.png)

- Dynamic Arrays with Hash Operator

The Hash Operator (#) Dynamically adds New Data to Arrays whereas The New Data would be Excluded if Added after Cell Range was Implemented. 

![counta with hash](https://github.com/peter28data/excel-modeling/blob/138127bb44dd9184f5803cc725516b9e005e8feb/images_excel_modeling/v2/counta_with_hashoperator.png)

- Hash Operator with Dynamic Aggregation

Although the Function is Dynamic the Cell Range is Limited. Adding a Hash Operator Enables the Function To Have a Dynamic Range.
![dynamic array filter table](https://github.com/peter28data/excel-modeling/blob/4249af4dbec8f8674fad188be6ea2a5974138e2d/images_excel_modeling/v3/dynamic_array_filter_table.png)

![hash average aggregation](https://github.com/peter28data/excel-modeling/blob/4249af4dbec8f8674fad188be6ea2a5974138e2d/images_excel_modeling/v3/hash_average_aggregation.png)

![counta wrong not needs index column](https://github.com/peter28data/excel-modeling/blob/4249af4dbec8f8674fad188be6ea2a5974138e2d/images_excel_modeling/v3/wrong_hash_counta.png)


- Hash Operator with Nested Indexing

The Frequency Function was counting all cells returned by the Filter Function. To isolate for The Client Names we nest the Index Function. 

![correct hash counta index col](https://github.com/peter28data/excel-modeling/blob/4249af4dbec8f8674fad188be6ea2a5974138e2d/images_excel_modeling/v3/correct_hash_counta_indexcol1.png
)


- Hash Operator with Dynamic Dropdown

Dropdown Menus can also be complemented with Hash Operators to Dynamically add new Records if the Dataset is Updated.

![hash dynamic dropdown](https://github.com/peter28data/excel-modeling/blob/4249af4dbec8f8674fad188be6ea2a5974138e2d/images_excel_modeling/v3/excel_dynamic_dropdown.png)


---

## 4. Excel XLOOKUP Function

Utilize for Dynamic Indexing and Retrieval.

- Return Matching Product Identifier Based on Product Details

![excel xlookup function](https://github.com/peter28data/excel-modeling/blob/2efbee4cca62b8a2435e08bc25195976565a8e2d/images_excel_modeling/v1/excel_xlookup_function_v1.png)

---

## 5. Excel UNIQUE Function

Utilize to Eliminate Redundant Transaction IDs to isolate distinct Customer Records. The system Performs Automated Data Deduplication.

-Remove Duplicates to Find Unique Customers or Products.

![Excel Unique Function](https://github.com/peter28data/excel-modeling/blob/2efbee4cca62b8a2435e08bc25195976565a8e2d/images_excel_modeling/v1/excel_unique_function.png)

---

## 6. Excel SORT Function

Utilize to Dynamically Sort Records in Ascending Alphabetical Order by Client Surname

-Arrange Data by Category.

![Excel SORT Function](https://github.com/peter28data/excel-modeling/blob/2efbee4cca62b8a2435e08bc25195976565a8e2d/images_excel_modeling/v1/excel_sort_function.png)

![Excel SORT Result](https://github.com/peter28data/excel-modeling/blob/1cc0eef2594347b6267b812eaeed5728c2dd59a0/images_excel_modeling/v1/excel_sort_result.png)

---

## 7. Excel PROPER Function

Utilize to Standardize Capitalization for Client Lists.

![Excel Proper Function](https://github.com/peter28data/excel-modeling/blob/2efbee4cca62b8a2435e08bc25195976565a8e2d/images_excel_modeling/v1/excel_proper_function.png)

![Excel Proper Result](https://github.com/peter28data/excel-modeling/blob/2efbee4cca62b8a2435e08bc25195976565a8e2d/images_excel_modeling/v1/excel_proper_result.png)

---

## 8. Excel LET Function

Implement Automated Summary reporting the Total Financial Performance of the UK Region and Average Order Size.

![Excel LET Function](https://github.com/peter28data/excel-modeling/blob/2efbee4cca62b8a2435e08bc25195976565a8e2d/images_excel_modeling/v1/excel_let_function.png)

---

## 9. Excel IFS Function

To Evaluate Box Volumes Shipped, The IFS function is utilized for Multi-Conditional Classification.

Good for Labeling Performance of Customer or Employee Records.

![Excel IFS Function](https://github.com/peter28data/excel-modeling/blob/2efbee4cca62b8a2435e08bc25195976565a8e2d/images_excel_modeling/v1/excel_ifs_function.png)

---

# 10. Excel Descriptive Statistics

Display Statistical Parameters for Amount Sold and Boxes Shipped such as Measures of Central Tendency (Mean & Median), Spread (Standard Deviation, Variance, Min & Max Range) To Measure Volatility of Sales Sizes and Shipping Volumes. 

![Excel Descriptive Statistics](https://github.com/peter28data/excel-modeling/blob/2efbee4cca62b8a2435e08bc25195976565a8e2d/images_excel_modeling/v1/excel_descriptive_statistics.png)



---

11. Excel Dot Operator 

To calculate the Profit, We select a Cell Range with Extra Empty Cells Selected if we want to Add New Figures in the future. The Profit column produced 0 for the rows where there is no current Data. 

=C3:C12-D3:D12
=C3:.C12-D3:.D12

- Adding a Dot Operator Before the end of the Range for both cell ranges selected. The Zeros disappear.

![excel dot operator](https://github.com/peter28data/excel-modeling/blob/821637ccda34485d1d43154fafdb6820a2718584/images_excel_modeling/v4/excel_dot_operator.png)

![dot operator vstack](https://github.com/peter28data/excel-modeling/blob/821637ccda34485d1d43154fafdb6820a2718584/images_excel_modeling/v4/dot_operator_vstack.png)

---

12. Excel Apostrophe Operator

Adding Indexing such as 0015 will be automatically populated as 15. 

With an Apostrophe '0015 The index will be populated as 0015.

- This can be Useful to show Formulas such as '=SUM(D3:D15)

- Another Use-Case is when Ranges want to be displayed such as '1-5 Otherwise it will populate a different input.

---

13. Excel Space Operator

Only sums the intersection of cell range.

![excel space operator](https://github.com/peter28data/excel-modeling/blob/821637ccda34485d1d43154fafdb6820a2718584/images_excel_modeling/v4/excel_space_operator.png)

To Find the Sum of both cell ranges a comma in between will produce the sum of both cell ranges.

---

14. Excel Asterisk Operator

Find the Number of Gmail accounts from Client List.

![wrong_asterisk_operator](https://github.com/peter28data/excel-modeling/blob/821637ccda34485d1d43154fafdb6820a2718584/images_excel_modeling/v4/wrong_asterisk_operator.png)

We Return 0 results due to the count if function reading the "@gmail.com" as an exact match and emails are preceded with unique client identifiers.

![correct asterisk operator](https://github.com/peter28data/excel-modeling/blob/821637ccda34485d1d43154fafdb6820a2718584/images_excel_modeling/v4/correct_asterisk_operator.png)

The Asterisk (*) Operator ignores the preceding characters for the Exact Match to Function correctly.

![excel_apostrophe_before_after](https://github.com/peter28data/excel-modeling/blob/821637ccda34485d1d43154fafdb6820a2718584/images_excel_modeling/v4/excel_apostrophe_before_after.png)

Unlimited Number of Characters before and after "Anna" are returned when "*anna*" is inputed. 

---

15. Excel Question Mark Operator

For Specified Number of Characters to Ignore when Search Matching.

1 Question Mark = 1 Character Value

![excel_question_mark](https://github.com/peter28data/excel-modeling/blob/821637ccda34485d1d43154fafdb6820a2718584/images_excel_modeling/v4/excel_question_mark.png)

---

16. Excel Tilda Operator

If Searching for cells with the Question mark and the Input, putting a Tilda will enable the countif function to Return 2 for this range. Otherwise it will Return 10 as The Question Mark is an Operator for a single Character Search.

![excel_tilda_operator](https://github.com/peter28data/excel-modeling/blob/f6204639eb5f89110b8150355a6feb84ad0c07bb/images_excel_modeling/v4/excel_tilda_operator.png)

Some cells have "M?" and "?S" due to Human Error. To search for any cells containing the Question Mark either before the Size or After, an Asterisk Tilda and Question Mark followed by another Asterisk will Dynamically work with the countif function.

![tilda_asterisk_question_mark](https://github.com/peter28data/excel-modeling/blob/f6204639eb5f89110b8150355a6feb84ad0c07bb/images_excel_modeling/v4/tilda_asterisk_question_mark.png)

---

17. Excel Dollar Sign Operator (Most Important)

Implementing a Dollar Sign Operator will freeze a component of a cell range such as the Column or Row.

![dollarsign_operator](https://github.com/peter28data/excel-modeling/blob/f6204639eb5f89110b8150355a6feb84ad0c07bb/images_excel_modeling/v4/dollarsign_operator.png)

---

18. Excel Ampersand (Most Versatile)

Replaces a Formula such as Concatenate to combine a cell with the first name with an adjacent cell containing the last name.

The concatenate function has a critical limitation when spaces cannot be included between the first and last name for a new cell.

![ampersand_operator](https://github.com/peter28data/excel-modeling/blob/f6204639eb5f89110b8150355a6feb84ad0c07bb/images_excel_modeling/v4/ampersand_operator.png)

If we search the Employee name "Max" to Return the revenue using the xlookup function it will return an error. This is because Max is recorded with a last name so it is not an Exact Match. This can be fixed with an Ampersand and Asterisk Operator as shown below while also selecting 2 for the "Wildcard" option in xlookup.

![ampersand_asterisk_wildcard](https://github.com/peter28data/excel-modeling/blob/f6204639eb5f89110b8150355a6feb84ad0c07bb/images_excel_modeling/v4/ampersand_asterisk_wildcard.png)

To Make the xlookup function more robust for Use-Cases where we search for the last name we add the same Ampersand and Asterisk Operators as shown below.

![ampersand_asterisk_start_end](https://github.com/peter28data/excel-modeling/blob/f0e9847416545ecb6afc7a2cc77db57007b27cb7/images_excel_modeling/v4/ampersand_asterisk_start_end.png)

---

# 19. Double Dash Operator

Remove dollar sign manually recorded with revenue transactions. The delimiter in this case is the ($).

=TEXTAFTER

![textafter data cleaning](https://github.com/peter28data/excel-modeling/blob/78a63dd889cd9095119eacd2aa36bdfc169bdcd0/images_excel_modeling/v5/textafter_data_cleaning.png)

The dollar sign is removed but the Data Type is still Text Data. The effect is that Numeric Functions such as =SUM and =MAX return a zero value. 

![wrong textafter sum max](https://github.com/peter28data/excel-modeling/blob/78a63dd889cd9095119eacd2aa36bdfc169bdcd0/images_excel_modeling/v5/wrong_textafter_sum_max.png)

The Double Dash Operator Turns the Text Data to Numeric Data by applying a double negative sign to produce a positive value, as shown below.

![correct_textafter_sum_max](https://github.com/peter28data/excel-modeling/blob/78a63dd889cd9095119eacd2aa36bdfc169bdcd0/images_excel_modeling/v5/correct_textafter_sum_max.png)


## Double Dash Investigating Potential Schedule Conflicts

Another Use Case for the Double Dash Operator is for converting True or False Results to Binary Code. This can be useful when we use the WEEKDAY function to convert calendar dates to 1-7 to represent days of the week. The Goal is to Take a Frequency count of recorded calendar dates that fall on 6 or 7 since they represent the Weekend as Saturday or Sunday. 

This is Useful for scheduling to avoid setting meetings for the weekend. 

1. Use WEEKDAY Function to convert Calendar Dates to Numbered Days of the Week

![weekday_number](https://github.com/peter28data/excel-modeling/blob/78a63dd889cd9095119eacd2aa36bdfc169bdcd0/images_excel_modeling/v5/weekday_number.png)

2. We isolate weekend dates with this function =WEEKDAY(B3:B12,2)>5 To Return True or False

![weekday_true_false](https://github.com/peter28data/excel-modeling/blob/78a63dd889cd9095119eacd2aa36bdfc169bdcd0/images_excel_modeling/v5/weekday_true_false.png)

3. Implement Double Dash Operator to Convert to Binary =--WEEKDAY(B3:B12,2)>5

![weekday_binary](https://github.com/peter28data/excel-modeling/blob/78a63dd889cd9095119eacd2aa36bdfc169bdcd0/images_excel_modeling/v5/weekday_binary.png)

4. Implement the SUMPRODUCT Function to Return the total number of Dates

![sumproduct_doubledash_weekday](https://github.com/peter28data/excel-modeling/blob/78a63dd889cd9095119eacd2aa36bdfc169bdcd0/images_excel_modeling/v5/sumproduct_doubledash_weekday.png)

## Phone Number Validation

Some manual recording may have more or less than 9 digits. Find the Proportion of mistakes made from recording.

1. Implement =LEN Function to calculate the number of characters in a text string.

![len_phone_number](https://github.com/peter28data/excel-modeling/blob/78a63dd889cd9095119eacd2aa36bdfc169bdcd0/images_excel_modeling/v5/len_phone_number.png)

2. Modify Function to Return True or False if Cell is not equal to 9 characters with =LEN(B3:B17)<>9

3. Implement Double Dash Operator to turn True or False into Binary Code

4. Nest SUMPRODUCT Function to Return Total number of Phone Numbers Recorded that are not 9 digits =SUMPRODUCT(--(LEN(B3:B17)<>9))

![sumproduct_len_not_9](https://github.com/peter28data/excel-modeling/blob/78a63dd889cd9095119eacd2aa36bdfc169bdcd0/images_excel_modeling/v5/sumproduct_len_not_9.png)

5. Calculate Proportion by Dividing by Total Number of Phone Numbers Recorded using COUNT Function.


## HR Multiple Criteria Segmentation and Proportion

HR has conducted Training on Interns and received reports from their managers. The Report is in True or False Markers for each Employee. 

1. Using Logic Statement such as =(C3:C9=TRUE),(D3:D9=TRUE) This will return a True or False if The employee has completed training and a practice doc on that training to evaluate the new hire performance. 

2. Implement Double Dash Operator to convert TRUE OR FALSE evaluations to binary code.

3. Implement SUMPRODUCT to summarize the total number of employees that passed training and practice doc.

![sumproduct_doubledash_multiple_criteria](https://github.com/peter28data/excel-modeling/blob/78a63dd889cd9095119eacd2aa36bdfc169bdcd0/images_excel_modeling/v5/sumproduct_doubledash_multiple_criteria.png)

4. Find Proportion by dividing by Total number of Employees for the Report.

## Budget Projection for Overtime Costs

An HR manager requires assistance calculating overtime costs for the past the past weekend for the sales division

1. Implement a logical statement to confirm employee is from sales division =(C3:C9="Sales")

2. Supplement Statement with criteria such as True for working the weekend =(C3:C9="Sales"),(D3:D9=TRUE)

3. Convert Result from TRUE OR FALSE to Binary Code with Double Dash Operator =--(C3:C9="Sales"),--(D3:D9=TRUE)

4. Implement SUMPRODUCT

![otcost_sumproduct_doubledash_multiplecriteria](https://github.com/peter28data/excel-modeling/blob/78a63dd889cd9095119eacd2aa36bdfc169bdcd0/images_excel_modeling/v5/otcost_sumproduct_doubledash_multiplecriteria.png)

5. Add Arrays to Calculate the Overtime Hourly Rate and Number of Hours worked on the Weekend

![otcost_sumproduct_doubledash_multiplecriteria_multiplication](https://github.com/peter28data/excel-modeling/blob/78a63dd889cd9095119eacd2aa36bdfc169bdcd0/images_excel_modeling/v5/otcost_sumproduct_doubledash_multiplecriteria_multiplication.png)

![complete_otcost_sumproduct](https://github.com/peter28data/excel-modeling/blob/78a63dd889cd9095119eacd2aa36bdfc169bdcd0/images_excel_modeling/v5/complete_otcost_sumproduct.png)

---

## 🤝 Done!  Thank you for Reading
For Project in SQL, click below:


1. SQL Portfolio Link: https://github.com/peter28data/SQL

2) Tableau Portfolio Link: https://github.com/peter28data/Tableau

3) Python Portfolio Link: https://github.com/peter28data/Python

4) Power BI Portfolio Link: https://github.com/peter28data/powerbi

5) Data Modeling Link: https://github.com/peter28data/data-modeling

6) Data Warehouse Link: https://github.com/peter28data/data-warehouse


---

<p><img align="center" src="https://github-readme-streak-stats.herokuapp.com/?user=peter28data&" alt="peter28data" /></p>


<h3 align="left">Languages and Tools:</h3>
<p align="left"> <a href="https://git-scm.com/" target="_blank" rel="noreferrer"> <img src="https://www.vectorlogo.zone/logos/git-scm/git-scm-icon.svg" alt="git" width="40" height="40"/> </a> <a href="https://pandas.pydata.org/" target="_blank" rel="noreferrer"> <img src="https://raw.githubusercontent.com/devicons/devicon/2ae2a900d2f041da66e950e4d48052658d850630/icons/pandas/pandas-original.svg" alt="pandas" width="40" height="40"/> </a> <a href="https://www.postgresql.org" target="_blank" rel="noreferrer"> <img src="https://raw.githubusercontent.com/devicons/devicon/master/icons/postgresql/postgresql-original-wordmark.svg" alt="postgresql" width="40" height="40"/> </a> <a href="https://www.python.org" target="_blank" rel="noreferrer"> <img src="https://raw.githubusercontent.com/devicons/devicon/master/icons/python/python-original.svg" alt="python" width="40" height="40"/> </a> <a href="https://scikit-learn.org/" target="_blank" rel="noreferrer"> <img src="https://upload.wikimedia.org/wikipedia/commons/0/05/Scikit_learn_logo_small.svg" alt="scikit_learn" width="40" height="40"/> </a> <a href="https://seaborn.pydata.org/" target="_blank" rel="noreferrer"> <img src="https://seaborn.pydata.org/_images/logo-mark-lightbg.svg" alt="seaborn" width="40" height="40"/> </a> </p>


---

## 📫 How to Connect With Me:

🔗 Email: peter.garayrobles@gmail.com 

---
