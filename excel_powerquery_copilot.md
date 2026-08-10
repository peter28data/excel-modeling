<h1 align="center">Peter Garay-Robles </h1>

<h3 align="center">A Data Engineer in SQL and Excel. </h3>

---

## Excel

Goal: Investigate the toal Player count for each Baseball player from 2025 MLB.

Challenge: Players can qualify for multiple positions in a given season and need to be counted towards each of them.

1. Find the Unique List of Positions available from dataset. Since positions are combined in the same cell, we will combine the entirety of the column in one cell with TEXTJOIN.

2. Split the contents of the Cell across Rows with the Delimiter "/" with TEXTSPLIT.

![textsplit_textjoin](https://github.com/peter28data/excel-modeling/blob/231893ce0f9f30aff00bc31526a7bd643dd52f9c/images_excel_modeling/v6/textsplit_textjoin.png)

3. Implement the UNIQUE Function to Return Distinct Baseball positions from spreadsheet.

![unique_textsplit_textjoin](https://github.com/peter28data/excel-modeling/blob/231893ce0f9f30aff00bc31526a7bd643dd52f9c/images_excel_modeling/v6/unique_textsplit_textjoin.png)

4. Utilize COUNTIFS to initially see how many players qualified for this position.

![wrong_countifs](https://github.com/peter28data/excel-modeling/blob/231893ce0f9f30aff00bc31526a7bd643dd52f9c/images_excel_modeling/v6/wrong_countifs.png)

Senior Analyst Perspective: The OF position displays 193 player count but does not account for players who qualify for OF and other positions. To Fix this Issue we implement Wildcard characters into the Criteria component of COUNTIFS. 

![correct_countifs](https://github.com/peter28data/excel-modeling/blob/231893ce0f9f30aff00bc31526a7bd643dd52f9c/images_excel_modeling/v6/correct_countifs.png)

Note: OF position displays 294, a 52% increase over our initial count after applying robust updates to our COUNTIFS Formula. 

We can verify this data by using the Dropdown on the Position from the Table. Inputing "OF" will click any unique inputs. The bottom left of Excel will confirms 294 records.

![dropdown_search_click](https://github.com/peter28data/excel-modeling/blob/9f0f5274e98fded42ce3c14bb0c4edb1db677815/images_excel_modeling/v7/dropdown_search_click.png)

5. Return a sorted list by nesting the UNIQUE and COUNTIF function inside a LET function. Let "position" and "players" be variables. Use HSTACK to position results horizontally. This is to return the same dynamic answers into one array for the purpose of sorting the output. If we attempted to do this previously, the sort function would not work with both results.

![let_unique_countifs_sort_hstack](https://github.com/peter28data/excel-modeling/blob/9f0f5274e98fded42ce3c14bb0c4edb1db677815/images_excel_modeling/v7/let_unique_countifs_sort_hstack.png)

---

## Python in Excel

1. Collect the raw dataset into a python dataframe and open a python editor in excel.

![python_dataframe_excel](https://github.com/peter28data/excel-modeling/blob/9f0f5274e98fded42ce3c14bb0c4edb1db677815/images_excel_modeling/v7/python_dataframe_excel.png)

2. Split values based on forward slash.

![python_split_strings](https://github.com/peter28data/excel-modeling/blob/9f0f5274e98fded42ce3c14bb0c4edb1db677815/images_excel_modeling/v7/python_split_strings.png)

3. Create individual rows for each players position. Duplicated names will appear for players who qualify for multiple positions.

![python_exploded](https://github.com/peter28data/excel-modeling/blob/9f0f5274e98fded42ce3c14bb0c4edb1db677815/images_excel_modeling/v7/python_exploded.png)

4. Utilize Python's groupby attribute to display components of Positions by Name followed by a frequency count using count attribute. Rename the column to "Player Count". Sort the dataframe using sort_values attribute in descending order. 

![python_groupby_count_sortvalues](https://github.com/peter28data/excel-modeling/blob/9f0f5274e98fded42ce3c14bb0c4edb1db677815/images_excel_modeling/v7/python_groupby_count_sortvalues.png)

---

## Copilot in Excel

1. Click on Copilot in the upper right tabs of Excel. Click on Agent Mode

![copilot_agentmode](https://github.com/peter28data/excel-modeling/blob/9f0f5274e98fded42ce3c14bb0c4edb1db677815/images_excel_modeling/v7/copilot_agentmode.png)

2. Input the Prompt: Organize data by unique player position and return a sorted listed of how many players qualified for each position.

![copilot_hardcoded](https://github.com/peter28data/excel-modeling/blob/9f0f5274e98fded42ce3c14bb0c4edb1db677815/images_excel_modeling/v7/copilot_hardcoded.png)

---

## Power Query in Excel

1. Use the dropdown menu on the Position column

![powerquery_position_dropdown](https://github.com/peter28data/excel-modeling/blob/9f0f5274e98fded42ce3c14bb0c4edb1db677815/images_excel_modeling/v7/powerquery_position_dropdown.png)

2. Select split column by delimiter

![powerquery_splitcolumn_delimiter](https://github.com/peter28data/excel-modeling/blob/9f0f5274e98fded42ce3c14bb0c4edb1db677815/images_excel_modeling/v7/powerquery_splitcolumn_delimiter.png)

3. Split columns by rows

![powerquery_split_rows](https://github.com/peter28data/excel-modeling/blob/9f0f5274e98fded42ce3c14bb0c4edb1db677815/images_excel_modeling/v7/powerquery_split_rows.png)

4. Utilize dropdown on new Position column to select Groupby

![powerquery_dropdown_groupby](https://github.com/peter28data/excel-modeling/blob/9f0f5274e98fded42ce3c14bb0c4edb1db677815/images_excel_modeling/v7/powerquery_dropdown_groupby.png)

5. Groupby Position and Select Count Rows

![powerquery_groupby_countrows](https://github.com/peter28data/excel-modeling/blob/9f0f5274e98fded42ce3c14bb0c4edb1db677815/images_excel_modeling/v7/powerquery_groupby_countrows.png)


6. Sort the Counted Rows by Descending order with the dropdown to confirm the 3rd most qualified position in MLB Baseball for that given year.

![powerquery_sorted](https://github.com/peter28data/excel-modeling/blob/9f0f5274e98fded42ce3c14bb0c4edb1db677815/images_excel_modeling/v7/powerquery_sorted.png)


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
