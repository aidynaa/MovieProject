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
### M Code
```
let
  Source = Excel.Workbook(File.Contents("C:\Users\aidyn\Downloads\Movie dataset Dashboard - Homework.xlsx"), null, true),
  #"Navigation 1" = Source{[Item = "Movie Data", Kind = "Sheet"]}[Data],
  #"Promoted headers" = Table.PromoteHeaders(#"Navigation 1", [PromoteAllScalars = true]),
  #"Changed column type" = Table.TransformColumnTypes(#"Promoted headers", {{"Movie Title", type text}, {"Release Date", type date}, {"Wikipedia URL", type text}, {"Genre_First_ID", Int64.Type}, {"Genre_Second_ID", Int64.Type}, {"Director_First_ID", Int64.Type}, {"Cast_First_ID", Int64.Type}, {"Cast_Second_ID", Int64.Type}, {"Cast_Third_ID", Int64.Type}, {"Cast_Fourth_ID", Int64.Type}, {"Cast_Fifth_ID", Int64.Type}, {"Budget ($)", Int64.Type}, {"Box Office Revenue ($)", type number}}),
  #"Merged queries" = Table.NestedJoin(#"Changed column type", {"Genre_First_ID"}, Genres, {"ID"}, "Genres", JoinKind.LeftOuter),
  #"Expanded Genres" = Table.ExpandTableColumn(#"Merged queries", "Genres", {"Genre"}, {"Genre.1"}),
  #"Reordered columns" = Table.ReorderColumns(#"Expanded Genres", {"Movie Title", "Release Date", "Wikipedia URL", "Genre_First_ID", "Genre", "Genre_Second_ID", "Director_First_ID", "Cast_First_ID", "Cast_Second_ID", "Cast_Third_ID", "Cast_Fourth_ID", "Cast_Fifth_ID", "Budget ($)", "Box Office Revenue ($)"}),
    #"Added Custom" = Table.AddColumn(#"Reordered columns", "ROI", each ([#"Box Office Revenue ($)"]-[#"Budget ($)"])/[#"Budget ($)"]),
    #"Changed Type" = Table.TransformColumnTypes(#"Added Custom",{{"ROI", Percentage.Type}}),
  #"Merged queries 1" = Table.NestedJoin(#"Changed Type", {"Genre_Second_ID"}, Genres, {"ID"}, "Genres", JoinKind.LeftOuter),
  #"Reordered columns 1" = Table.ReorderColumns(#"Merged queries 1", {"Movie Title", "Release Date", "Wikipedia URL", "Genre_First_ID", "Genre", "Genre_Second_ID", "Genre 2nd", "Director_First_ID", "Cast_First_ID", "Cast_Second_ID", "Cast_Third_ID", "Cast_Fourth_ID", "Cast_Fifth_ID", "Budget ($)", "Box Office Revenue ($)"}),
  #"Merged queries 2" = Table.NestedJoin(#"Reordered columns 1", {"Director_First_ID"}, Directors, {"ID"}, "Directors", JoinKind.LeftOuter),
  #"Reordered columns 2" = Table.ReorderColumns(#"Merged queries 2", {"Movie Title", "Release Date", "Wikipedia URL", "Genre_First_ID", "Genre", "Genre_Second_ID", "Genre 2nd", "Director_First_ID", "Director", "Cast_First_ID", "Cast_Second_ID", "Cast_Third_ID", "Cast_Fourth_ID", "Cast_Fifth_ID", "Budget ($)", "Box Office Revenue ($)"}),
  #"Merged queries 3" = Table.NestedJoin(#"Reordered columns 2", {"Cast_First_ID"}, Actors, {"ID"}, "Actors", JoinKind.LeftOuter),
  #"Reordered columns 3" = Table.ReorderColumns(#"Merged queries 3", {"Movie Title", "Release Date", "Wikipedia URL", "Genre_First_ID", "Genre", "Genre_Second_ID", "Genre 2nd", "Director_First_ID", "Director", "Cast_First_ID", "Actor", "Cast_Second_ID", "Cast_Third_ID", "Cast_Fourth_ID", "Cast_Fifth_ID", "Budget ($)", "Box Office Revenue ($)"}),
  #"Expanded Actors 1" = Table.ExpandTableColumn(#"Reordered columns 3", "Actors", {"Actor"}, {"Actor.1"}),
    #"Added Custom1" = Table.AddColumn(#"Expanded Actors 1", "ROI.1", each ([#"Box Office Revenue ($)"]-[#"Budget ($)"])/[#"Budget ($)"]),
  #"Reordered columns 4" = Table.ReorderColumns(#"Added Custom1", {"Movie Title", "Release Date", "Wikipedia URL", "Genre_First_ID", "Genre", "Genre_Second_ID", "Genre 2nd", "Director_First_ID", "Director", "Cast_First_ID", "Actor", "Cast_Second_ID", "Actor.1", "Cast_Third_ID", "Cast_Fourth_ID", "Cast_Fifth_ID", "Budget ($)", "Box Office Revenue ($)"}),
  #"Merged queries 5" = Table.NestedJoin(#"Reordered columns 4", {"Cast_Third_ID"}, Actors, {"ID"}, "Actors", JoinKind.LeftOuter),
  #"Expanded Actors 2" = Table.ExpandTableColumn(#"Merged queries 5", "Actors", {"Actor"}, {"Actor.2"}),
  #"Renamed columns 1" = Table.RenameColumns(#"Expanded Actors 2", {{"Actor.2", "Cast 3rd"}}),
  #"Reordered columns 5" = Table.ReorderColumns(#"Renamed columns 1", {"Movie Title", "Release Date", "Wikipedia URL", "Genre_First_ID", "Genre", "Genre_Second_ID", "Genre 2nd", "Director_First_ID", "Director", "Cast_First_ID", "Actor", "Cast_Second_ID", "Actor.1", "Cast_Third_ID", "Cast 3rd", "Cast_Fourth_ID", "Cast_Fifth_ID", "Budget ($)", "Box Office Revenue ($)"}),
  #"Renamed columns 2" = Table.RenameColumns(#"Reordered columns 5", {{"Actor.1", "Cast 2nd"}, {"Actor", "Cast 1st"}}),
  #"Merged queries 6" = Table.NestedJoin(#"Renamed columns 2", {"Cast_Fourth_ID"}, Actors, {"ID"}, "Actors", JoinKind.LeftOuter),
  #"Expanded Actors 3" = Table.ExpandTableColumn(#"Merged queries 6", "Actors", {"Actor"}, {"Actor"}),
  #"Renamed columns 3" = Table.RenameColumns(#"Expanded Actors 3", {{"Actor", "Cast 4th"}}),
  #"Reordered columns 6" = Table.ReorderColumns(#"Renamed columns 3", {"Movie Title", "Release Date", "Wikipedia URL", "Genre_First_ID", "Genre", "Genre_Second_ID", "Genre 2nd", "Director_First_ID", "Director", "Cast_First_ID", "Cast 1st", "Cast_Second_ID", "Cast 2nd", "Cast_Third_ID", "Cast 3rd", "Cast_Fourth_ID", "Cast 4th", "Cast_Fifth_ID", "Budget ($)", "Box Office Revenue ($)"}),
  #"Merged queries 7" = Table.NestedJoin(#"Reordered columns 6", {"Cast_Fifth_ID"}, Actors, {"ID"}, "Actors", JoinKind.LeftOuter),
  #"Expanded Actors 4" = Table.ExpandTableColumn(#"Merged queries 7", "Actors", {"Actor"}, {"Actor"}),
  #"Renamed columns 4" = Table.RenameColumns(#"Expanded Actors 4", {{"Actor", "Cast 5th"}}),
  #"Reordered columns 7" = Table.ReorderColumns(#"Renamed columns 4", {"Movie Title", "Release Date", "Wikipedia URL", "Genre_First_ID", "Genre", "Genre_Second_ID", "Genre 2nd", "Director_First_ID", "Director", "Cast_First_ID", "Cast 1st", "Cast_Second_ID", "Cast 2nd", "Cast_Third_ID", "Cast 3rd", "Cast_Fourth_ID", "Cast 4th", "Cast_Fifth_ID", "Cast 5th", "Budget ($)", "Box Office Revenue ($)"})
in
    #"Reordered columns 7"
```
