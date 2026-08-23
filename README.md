# Netflix Data Analysis Dashboard

## Project Overview

This project analyzes the Netflix content catalog using **Microsoft Excel**. The objective is to explore the distribution of Movies and TV Shows, analyze content release trends, understand audience rating categories, examine content age and movie length, and build an interactive dashboard.

The analysis was performed using **Excel functions, data cleaning, data transformation, Pivot Tables, Pivot Charts, KPI cards, and slicers**.

---

# Project Objectives

The main objectives of this project are:

* Analyze the distribution of Movies and TV Shows.
* Study Netflix content release trends over the years.
* Analyze content based on audience rating categories.
* Compare New, Recent, and Old content across rating categories.
* Analyze content distribution based on movie length.
* Compare content type across different movie length categories.
* Compare content age across different movie length categories.
* Build an interactive Excel dashboard.

---

# Dataset Information:

The dataset contains information about Netflix Movies and TV Shows.

The original dataset contains the following columns:

| Column       | Description                                |
| ------------ | ------------------------------------------ |
| show_id      | Unique identifier for each title           |
| type         | Type of content: Movie or TV Show          |
| title        | Name of the Movie or TV Show               |
| director     | Director of the content                    |
| cast         | Cast members                               |
| country      | Country where the content was produced     |
| date_added   | Date when the content was added to Netflix |
| release_year | Original release year                      |
| rating       | Audience rating                            |
| duration     | Duration of the Movie or TV Show           |
| listed_in    | Genre or category                          |
| description  | Brief description of the content           |

---

# Tools and Technologies

* Microsoft Excel
* Pivot Tables
* Pivot Charts
* Excel Functions
* Conditional Formatting
* Slicers
* Data Cleaning
* Data Transformation

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
Helper Columns Creation
        ├── Content Age
        ├── Movie Length
        ├── Rating Category
        └── Movie Age
        ↓
Pivot Table Analysis
        ├── Movies vs TV Shows
        ├── Content by Release Year
        ├── Category Split with Content Age
        ├── Rating Category Distribution
        ├── Content Distribution by Movie Length
        ├── Content Type by Movie Length
        └── Content Age by Movie Length
        ↓
Data Visualization
        ↓
KPI Development
        ↓
Slicer Integration
        ↓
Interactive Netflix Dashboard
```

---

# Data Cleaning and Missing Value Analysis

The dataset was first analyzed to identify missing values and duplicate records.

The following missing values were identified:

| Column     | Missing Count |
| ---------- | ------------: |
| Director   |          2634 |
| Cast       |           825 |
| Country    |           831 |
| Date Added |            10 |
| Rating     |             4 |
| Duration   |             3 |

No duplicate records were found in the dataset.

## Excel Features Used

### COUNTBLANK()

The `COUNTBLANK()` function was used to count missing values in selected columns.

### Conditional Formatting

Conditional Formatting was used to highlight missing or blank values, making it easier to identify data quality issues before analysis.

---

# Data Transformation

Additional helper columns were created to make the dataset easier to analyze.

## 1. Content Age

The `release_year` column was used to categorize content into three groups:

| Condition            | Category |
| -------------------- | -------- |
| Release Year >= 2019 | New      |
| Release Year >= 2010 | Recent   |
| Otherwise            | Old      |

---

## 2. Movie Length

The `duration` column was categorized into different content length groups.

| Duration        | Category            |
| --------------- | ------------------- |
| Contains season | Series/Multi-Series |
| <= 120 min      | Medium              |
| > 120 min       | Long                |
| < 60 min        | Short               |

These categories were used to analyze the distribution of Netflix content based on duration and format.

---

## 3. Rating Category

The original rating values were grouped into broader audience categories.

| Rating                | Rating Category |
| --------------------- | --------------- |
| TV-MA, R, NC-17       | Adults          |
| TV-14, PG-13          | Teens           |
| TV-PG, PG, TV-G, G    | Family          |
| TV-Y, TV-Y7, TV-Y7-FV | Kids            |
| NR, UR, No Rating     | No Rating       |

This transformation simplified the analysis of audience categories.

---

## 4. Movie Age

A Movie Age column was created to calculate how old a title is based on its release year.

```excel
=YEAR(TODAY())-[@release_year]
```

This column was used as an additional transformation for analyzing the age of Netflix content.

---

# Pivot Table Analysis

Pivot Tables were created to summarize the Netflix dataset and perform different types of analysis.

| Analysis                             | Rows            | Columns     | Values         | Purpose                                                                       |
| ------------------------------------ | --------------- | ----------- | -------------- | ----------------------------------------------------------------------------- |
| Movies vs TV Shows                   | Type            | —           | Count of Title | Compare the distribution of Movies and TV Shows                               |
| Content by Release Year              | Release Year    | —           | Count of Title | Analyze content release trends over the years                                 |
| Category Split with Content Age      | Rating Category | Content Age | Count of Title | Compare New, Recent, and Old content across rating categories                 |
| Rating Category Distribution         | Rating Category | —           | Count of Title | Analyze content distribution across audience categories                       |
| Content Distribution by Movie Length | Movie Length    | —           | Count of Title | Analyze the overall distribution of content based on movie length             |
| Content Type by Movie Length         | Movie Length    | Type        | Count of Title | Compare Movies and TV Shows across different movie length categories          |
| Content Age by Movie Length          | Movie Length    | Content Age | Count of Title | Compare New, Recent, and Old content across different movie length categories |

---

# Dashboard KPIs

The dashboard includes the following key performance indicators:

| KPI          | Description                                        |
| ------------ | -------------------------------------------------- |
| Total Titles | Total number of Movies and TV Shows in the dataset |
| Top Category | Rating category with the highest number of titles  |
| Start Date   | Earliest release year available in the dataset     |
| End Date     | Latest release year available in the dataset       |

---

# Dashboard Visualizations

## 1. Movies vs TV Shows

A Doughnut Chart is used to compare the distribution of Movies and TV Shows in the Netflix catalog.

This visualization shows that Movies represent the majority of titles, while TV Shows account for a smaller portion.

---

## 2. Content by Release Year

A Line Chart is used to analyze the number of titles released across different years.

This visualization helps identify trends and growth in Netflix's content catalog over time.

---

## 3. Category Split with Content Age

A Clustered Column Chart compares different rating categories with Content Age.

The content is divided into:

* New
* Recent
* Old

This visualization shows how content age varies across Adults, Teens, Family, Kids, and No Rating categories.

---

## 4. Rating Category Distribution

A Pie Chart displays the distribution of Netflix content across different audience categories.

The categories include:

* Adults
* Teens
* Family
* Kids
* No Rating

This visualization helps identify the largest audience category in the Netflix catalog.

---

## 5. Content Distribution by Movie Length

A chart is used to analyze the total number of titles across different movie length categories.

The categories include:

* Long
* Medium
* Series/Multi-Series
* Short

This visualization helps identify the most common content length category in the dataset.

---

## 6. Content Type by Movie Length

A Pivot Chart compares Movies and TV Shows across different movie length categories.

The analysis includes:

* Long
* Medium
* Series/Multi-Series
* Short

This visualization helps identify how content type varies based on movie length.

---

## 7. Content Age by Movie Length

A 100% Stacked Column Chart compares the percentage distribution of content age across different movie length categories.

The content age categories include:

* New
* Recent
* Old

The movie length categories include:

* Long
* Medium
* Series/Multi-Series
* Short

This visualization helps understand how the age distribution of Netflix content varies based on its duration or format.

---

# Interactive Features

The dashboard includes interactive slicers that allow users to dynamically filter the analysis.

## Release Year Slicer

Allows users to filter the dashboard based on one or multiple release years.

## Content Age Slicer

Allows users to filter the dashboard based on:

* New
* Recent
* Old

The connected Pivot Tables and Pivot Charts update dynamically based on the selected filters.

---

# Key Insights

1. **Movies dominate the catalog**, accounting for approximately 70% of titles.

2. **Content releases grew rapidly after 2018.**

3. **Adult-rated content has the largest share**, followed by Teen and Family content.

4. **Recent content is more common than older content** across most rating categories.

5. **Medium-length content is the most common**, while short content represents a smaller portion of the catalog.

6. **Content distribution varies by type and age**, with Movies and TV Shows showing different patterns across movie length categories.

---

# Dashboard Design

The dashboard uses a Netflix-inspired color theme based on:

* Black background
* Netflix Red for highlights and primary data
* White and light gray for contrast
* Consistent borders, chart titles, and KPI cards

The dashboard follows a structured layout containing:

* 4 KPI Cards
* 7 Pivot Charts
* Key Insights Panel
* Release Year Slicer
* Content Age Slicer

---

# Dataset Limitations

1. The dataset is a static snapshot and may not represent Netflix's latest catalog.

2. Some records contain multiple countries and genres in a single cell, limiting detailed country-wise and genre-wise analysis.

3. Missing values are present in columns such as Director, Cast, Country, Date Added, Rating, and Duration.

4. The dataset represents content from multiple countries and provides an overall view of the Netflix catalog.

5. The analysis was performed using Microsoft Excel, which is suitable for small to medium-sized datasets but has limitations when working with large-scale data.

---

# Future Scope

Possible future improvements include:

1. Build a more advanced interactive dashboard using Power BI or Tableau.

2. Perform detailed data analysis using SQL.

3. Use Python and Pandas for advanced data cleaning and exploratory data analysis.

4. Separate multiple countries and genres for more detailed analysis.

5. Integrate updated Netflix data for more recent insights.

6. Develop a Machine Learning recommendation system based on user preferences and content genres.

---

# Project Structure

```text
Netflix-Data-Analysis/
│
├── README.md
│
├── data/
│   └── netflix_titles.csv
│
├── excel/
│   └── Netflix_Data_Analysis.xlsx
│
└── images/
    ├── dashboard.png
```

---

# Conclusion

This project demonstrates how Microsoft Excel can be used for end-to-end data analysis, including data understanding, missing value analysis, data transformation, Pivot Table analysis, data visualization, and interactive dashboard development.

The raw Netflix dataset was transformed into meaningful insights using helper columns, Pivot Tables, Pivot Charts, KPI cards, and slicers. The final dashboard provides an interactive view of Netflix content based on content type, release year, audience rating, content age, and movie length.

This project serves as a foundation for performing more advanced analysis using **SQL, Python, Power BI, Tableau, and Machine Learning**.
