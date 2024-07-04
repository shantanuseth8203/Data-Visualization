Data Visualization Project : AdventureWorks Sales Analysis Dashboard

Description:

This project is a data visualization dashboard built using Dash, Plotly, and Dash Bootstrap Components. The dashboard provides an interactive analysis of sales data from the AdventureWorks dataset. The application showcases various visualizations such as sales trends over time, monthly sales statistics, and distribution of bike sales prices by color and type.

Features:

- Sales Over Time: A line chart displaying the total sales over time with additional lines indicating the minimum, maximum, and average sales. A linear regression line is also plotted to show the trend.
- Monthly Sales, Costs, and Profit: A grouped bar chart presenting the monthly aggregated sales, costs, and profit.
- Bike Sales Price Distribution by Color: A violin plot illustrating the distribution of bike sales prices by color.
- Sales Price Distribution by Bike Type: A density heatmap showing the distribution of sales prices across different bike types (Mountain Bikes, Road Bikes, and Touring Bikes).

Data Preparation:

The data is sourced from the AdventureWorks Excel file containing three sheets: Products, Customers, and Sales. The data preparation steps include:

- Reading data from each sheet into a pandas DataFrame.
- Cleaning the data by removing blanks and duplicates.
- Converting the Date column to datetime format.
- Grouping and summarizing the sales data by date and month.
- Calculating key statistics such as minimum, maximum, average sales, and standard deviation.
- Dashboard Layout

The dashboard layout includes:

- Title: A main title at the top of the dashboard.
- Statistic Cards: Cards displaying total sales, maximum sales, minimum sales, average sales, and standard deviation.
- Graphs:
    - Sales Over Time: An interactive line chart.
    - Monthly Sales, Costs, and Profit: A grouped bar chart.
    - Bike Sales Price Distribution by Color: A violin plot.
    - Sales Price Distribution by Bike Type: A density heatmap.


# Data-Visualization
