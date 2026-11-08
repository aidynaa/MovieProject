# MovieProject
## Table of Content
[Problem Statement](#problem-statement)
[Data Source](#data-source)
[Tools](#tools)
[Data Cleaning and Preparation](#data-cleaning-and-preparation)
[]
### Problem Statement
Netflix wants to better understand which movie they should produce next, including the most suitable actors and directors. We have a dataset containing movie budgets, box office performance, actors, directors, and genres. 
- The best-performing actors
- The top movie based on box office metrics
- Directors performance
- Genre trends
- Seasonal patterns in movie performance
### Data Source
Movie Data: The primary dataset used for this analysis is the "Movie Data Homework.xlsx" file, containing detailed information about each movie's performance (budgets, box office performance, actors, directors, and genres). Download the original data source here: [Movie Dataset Excel file](https://github.com/user-attachments/files/34394232/Movie.dataset.xlsx)
### Tools
1. Power Query - I used Power Query for Data Cleaning
2. Excel - I used Excel for Data Analysis
3. Pivot Table - for Creating the Dashboard and Visualizations
### Data Cleaning and Preparation
- Data loading and inspection.
- Handling errors and missing values.
- Data cleaning and formatting. The Excel file after the data cleaning & preparation process can be downloaded here: [Ready to Use: Movie Dashboard.xlsx](https://github.com/user-attachments/files/34394430/Movie.Dashboard.xlsx)
### Dashboard
<img width="1710" height="1238" alt="image" src="https://github.com/user-attachments/assets/89bd8957-88ed-4ff2-8719-c23cbcfde154" />

### Challenges
```
= Table.Group(#"Sorted Rows", {"animal_id"}, 
{{"IndexColumn", each Table.AddIndexColumn(_,"Index",1,1,Int64.Type)}})
```
### Recommendation
Top 5 genres are Action, Comedy,
