# Netflix Data Analysis Dashboard

## Project Overview

This project analyzes the Netflix content catalog using Microsoft Excel. The objective is to transform raw Netflix data into meaningful insights through data cleaning, missing value analysis, data transformation, Pivot Tables, Pivot Charts, KPIs, and an interactive dashboard.

The project follows a structured Excel-based data analysis workflow:

1. Raw Data
2. Missing Data Analysis
3. Data Transformation
4. Pivot Table Analysis
5. Dashboard Development

The final dashboard provides insights into the distribution of Movies and TV Shows, content release trends, rating categories, and content age.

---

# Project Objectives

The main objectives of this project are:

* Analyze the distribution of Movies and TV Shows.
* Study the number of titles released over different years.
* Analyze content based on audience rating categories.
* Categorize content based on age and duration.
* Identify missing values and check for duplicate records.
* Transform raw data into meaningful categories.
* Create Pivot Tables and Pivot Charts for analysis.
* Develop KPI cards for a quick overview.
* Build an interactive Excel dashboard using slicers.

---

# Tools and Technologies

* Microsoft Excel
* Excel Functions
* Data Cleaning
* Data Transformation
* Conditional Formatting
* Pivot Tables
* Pivot Charts
* Slicers
* Dashboard Development

---

# Project Workflow

```text
Raw Netflix Dataset
        ↓
Data Understanding
        ↓
Missing Value Analysis
        ↓
Duplicate Check
        ↓
Data Transformation
        ↓
Helper Columns
        ↓
Pivot Table Analysis
        ↓
Pivot Charts and KPIs
        ↓
Interactive Excel Dashboard
```

---

# File 01: Raw Data

The first sheet contains the original Netflix dataset used for this project.

Each row represents one Netflix title.

## Original Dataset Columns

| Column         | Description                                                 |
| -------------- | ----------------------------------------------------------- |
| `show_id`      | Unique identifier for each Netflix title                    |
| `type`         | Type of content: Movie or TV Show                           |
| `title`        | Name of the Movie or TV Show                                |
| `director`     | Name of the director                                        |
| `cast`         | Cast members associated with the title                      |
| `country`      | Country or countries where the content was produced         |
| `date_added`   | Date when the title was added to Netflix                    |
| `release_year` | Original release year of the title                          |
| `rating`       | Audience rating or maturity classification                  |
| `duration`     | Movie duration in minutes or number of seasons for TV Shows |
| `listed_in`    | Genre or categories associated with the title               |
| `description`  | Brief description of the content                            |

The raw dataset was used as the starting point for the analysis. No additional helper columns or transformed categories were present at this stage.

---

# File 02: Missing Data Analysis

The second sheet focuses on data quality analysis.

The dataset was checked for duplicate records and missing values before proceeding with further analysis.

## Duplicate Check

The dataset was checked for duplicate records.

**Result: No duplicate records were found.**

## Missing Value Summary

The following missing values were identified:

| Column     | Missing Count |
| ---------- | ------------: |
| Director   |         2,634 |
| Cast       |           825 |
| Country    |           831 |
| Date Added |            10 |
| Rating     |             4 |
| Duration   |             3 |

The highest number of missing values was found in the `director` column, followed by `country` and `cast`.

## Excel Features Used

### COUNTBLANK()

The `COUNTBLANK()` function was used to count blank cells in each relevant column.

### Conditional Formatting

Conditional Formatting was used to highlight blank cells and make missing values easier to identify before analysis.

## Missing Data Handling

Not all columns with missing values were required for the final dashboard analysis.

* Missing values in `director`, `cast`, and `country` were not the primary focus of the dashboard.
* Missing ratings were grouped into the **No Rating** category during data transformation.
* The small number of missing values in `date_added` and `duration` were considered during analysis.

---

# File 03: Data Transformation

After analyzing the raw data and missing values, additional helper columns were created to simplify the analysis.

The following columns were added:

* Content Age
* Movie Length
* Rating Category
* Movie Age

---

## 1. Content Age

The `release_year` column was used to classify titles into three categories.

| Release Year   | Content Age |
| -------------- | ----------- |
| 2019 and later | New         |
| 2010 to 2018   | Recent      |
| Before 2010    | Old         |

This transformation was used to compare the age of Netflix content across different rating categories.

---

## 2. Movie Length

The `duration` column contains different formats for Movies and TV Shows.

Examples:

* `90 min`
* `120 min`
* `2 Seasons`

The duration values were categorized as follows:

| Duration Condition    | Result              |
| --------------------- | ------------------- |
| Contains Season       | Series/Multi-Series |
| Less than 60 minutes  | Short               |
| 60 to 120 minutes     | Medium              |
| More than 120 minutes | Long                |

This transformation made it easier to group Netflix content according to duration.

---

## 3. Rating Category

The original rating values were grouped into broader audience categories.

| Original Rating       | Rating Category |
| --------------------- | --------------- |
| TV-MA, R, NC-17       | Adults          |
| TV-14, PG-13          | Teens           |
| TV-PG, PG, TV-G, G    | Family          |
| TV-Y, TV-Y7, TV-Y7-FV | Kids            |
| NR, UR, No Rating     | No Rating       |

This transformation simplified the rating analysis and made the dashboard easier to interpret.

---

## 4. Movie Age

A `Movie Age` column was created to calculate how old each title is based on its release year.

The Excel formula used was:

```excel
=YEAR(TODAY())-[@[release_year]]
```

This calculates the difference between the current year and the title's release year.

For example:

| Release Year | Movie Age |
| ------------ | --------: |
| 2020         |         6 |
| 2015         |        11 |
| 2010         |        16 |

The calculation is dynamic because it uses the `TODAY()` function and updates automatically each year.

---

# Data Transformation Summary

| Helper Column   | Source Column  | Purpose                                        |
| --------------- | -------------- | ---------------------------------------------- |
| Content Age     | `release_year` | Categorize titles as New, Recent, or Old       |
| Movie Length    | `duration`     | Categorize content by duration                 |
| Rating Category | `rating`       | Group ratings into broader audience categories |
| Movie Age       | `release_year` | Calculate the age of each title                |

---

# Excel Functions Used

The following functions and features were used during the project:

* `COUNTBLANK()`
* `IF()` / Nested `IF()`
* `SEARCH()`
* `YEAR()`
* `TODAY()`
* Conditional Formatting
* Pivot Tables
* Pivot Charts
* Slicers

---

# Pivot Table Analysis

Pivot Tables were created to summarize the Netflix dataset and perform the main analysis.

| Analysis                        | Rows / Categories               | Values         | Purpose                                                        |
| ------------------------------- | ------------------------------- | -------------- | -------------------------------------------------------------- |
| Movies vs TV Shows              | Type                            | Count of Type  | Compare the distribution of Movies and TV Shows                |
| Content by Release Year         | Release Year                    | Count of Title | Analyze content release trends                                 |
| Category Split with Content Age | Rating Category and Content Age | Count of Title | Compare New, Recent, and Old content across rating categories  |
| Rating Category Distribution    | Rating Category                 | Count of Title | Analyze the distribution of content across audience categories |

---

# Dashboard Overview

The final dashboard was created using Microsoft Excel.

It combines:

* KPI Cards
* Pivot Charts
* Interactive Slicers

The dashboard provides a summary of the Netflix dataset and allows users to interact with the visualizations using filters.

---

# Dashboard KPIs

The dashboard includes the following KPI cards:

| KPI            | Description                                                      |
| -------------- | ---------------------------------------------------------------- |
| Total Titles   | Displays the total number of titles in the dataset               |
| Top Category   | Highlights the category with the highest number of titles        |
| Recent Content | Provides an overview of recently released content                |
| Top Rating     | Identifies the rating category with the highest number of titles |

---

# Dashboard Visualizations

## 1. Movies vs TV Shows

A Doughnut Chart was used to compare the distribution of Movies and TV Shows.

The analysis shows that:

* Movies account for approximately 70% of the dataset.
* TV Shows account for approximately 30% of the dataset.

This indicates that Movies make up the majority of titles in the Netflix dataset.

---

## 2. Content by Release Year

A Line Chart was used to visualize the number of titles released across different years.

The chart shows significant growth in the number of titles released in recent years, particularly after 2018.

---

## 3. Category Split with Content Age

A Clustered Column Chart compares rating categories with Content Age.

The Content Age categories are:

* New
* Recent
* Old

The rating categories include:

* Adults
* Teens
* Family
* Kids
* No Rating

This visualization helps analyze how the age of content varies across different audience categories.

---

## 4. Rating Category Distribution

A Pie Chart was used to visualize the distribution of Netflix content across audience rating categories.

The analysis includes:

* Adults
* Teens
* Family
* Kids
* No Rating

Adult-rated content represents the largest share of the dataset.

---

# Interactive Features

The dashboard includes slicers that allow users to dynamically filter the analysis.

## Release Year Slicer

Users can select one or multiple release years to explore content from specific periods.

## Content Age Slicer

Users can filter the dashboard based on:

* New
* Recent
* Old

The connected Pivot Tables and Pivot Charts update dynamically based on the selected filters.

---

# Key Insights

1. Movies account for approximately 70% of the Netflix dataset, while TV Shows account for approximately 30%.

2. The number of titles increased significantly in recent years, with strong growth after 2018.

3. Adult-rated content represents the largest audience category in the dataset.

4. Teen and Family content also make up a significant portion of the Netflix catalog.

5. Kids and No Rating categories represent a smaller share of the dataset.

6. Recent content is prominent across major rating categories.

---

# Dataset Limitations

This project has the following limitations:

1. The dataset is a static snapshot and may not represent Netflix's latest catalog.

2. Missing values are present in columns such as Director, Cast, Country, Date Added, Rating, and Duration.

3. Some titles contain multiple countries in a single cell, making detailed country-level analysis more difficult.

4. Some titles contain multiple genres in a single cell, which limits detailed genre analysis without additional transformation.

5. The analysis represents the overall dataset and is not specific to a particular country.

6. The Movie Age calculation is dynamic because it uses the current year.

7. Excel is suitable for small to medium-sized datasets but may have limitations when working with very large datasets.

---

# Future Scope

Possible future improvements include:

1. Build a more advanced dashboard using Power BI or Tableau.
2. Perform additional analysis using SQL.
3. Use Python and Pandas for advanced data cleaning and exploratory analysis.
4. Separate multiple countries and genres into individual records.
5. Add country-wise and genre-wise analysis.
6. Analyze directors and cast members.
7. Use an updated dataset for more recent insights.
8. Build a recommendation system using Machine Learning.


---

# Conclusion

This project demonstrates a complete data analysis workflow using Microsoft Excel.

The project involved:

* Understanding the raw dataset
* Checking for duplicate records
* Performing missing value analysis
* Using `COUNTBLANK()` and Conditional Formatting
* Transforming raw data into meaningful categories
* Creating helper columns
* Calculating Movie Age
* Creating Content Age categories
* Categorizing Movie Length
* Grouping ratings into audience categories
* Creating Pivot Tables
* Building Pivot Charts
* Developing KPI cards
* Adding interactive slicers
* Designing an interactive Excel dashboard

The raw Netflix dataset was transformed into an interactive dashboard that provides insights into content type distribution, release trends, audience categories, and content age.

This project demonstrates practical skills in **Data Cleaning, Data Transformation, Excel Functions, Pivot Tables, Data Visualization, and Dashboard Development**.
