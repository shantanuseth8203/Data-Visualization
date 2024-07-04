import dash
from dash import dcc, html
from dash.dependencies import Input, Output
import pandas as pd
import plotly.graph_objs as go
import dash_bootstrap_components as dbc
import plotly.express as px
import numpy as np

# Specify file path
file_path = 'AdventureWorks.xlsx'

# Create a dataframe for each sheet 
df_products = pd.read_excel(file_path, sheet_name='Products')
df_customers = pd.read_excel(file_path, sheet_name='Customers')
df_sales = pd.read_excel(file_path, sheet_name='Sales')

# Remove blanks and duplicates
for df in [df_products, df_customers, df_sales]:
    df.dropna(inplace=True)
    df.drop_duplicates(inplace=True)

# Ensure Date column is datetime
df_sales['Date'] = pd.to_datetime(df_sales['Date'])

# Group Sales by Date
sales_by_date = df_sales.groupby('Date')['Sales'].sum().reset_index()

# Calculate min, max, and average sales
min_sales = sales_by_date['Sales'].min()
max_sales = sales_by_date['Sales'].max()
avg_sales = sales_by_date['Sales'].mean()
std_dev_sales = sales_by_date['Sales'].std()

# Group sales by month and sum numerical columns
df_scp_monthly = df_sales.groupby(df_sales['Date'].dt.to_period('M'))[['Sales', 'Costs']].sum()

# Calculate profit 
df_scp_monthly['Profit'] = df_scp_monthly['Sales'] - df_scp_monthly['Costs']

###############################################################################################

# Initialize the Dash app
app = dash.Dash(__name__, external_stylesheets=[dbc.themes.DARKLY])

# Function to create styled statistic cards
def create_card(card_id, title, value):
    card = dbc.Card(
        dbc.CardBody([
            html.H5(title, className="card-title", style={"text-align": "center", "font-size": "17px", "color": "#FFFFFF"}),
            html.H3(value, className="card-text", style={"text-align": "center", "font-size": "17px", "color": "#FFFFFF"}),
        ]),
        style={"background-color": "#343a40", "margin": "5px", "width": "160px", "height": "100px"}
    )
    return card

# Define the layout of the app
app.layout = html.Div([
    html.H1("Sales Analysis", style={"text-align": "center", "color": "#FFFFFF", "margin-top": "60px", "margin-bottom": "60px", "font-size": "36px"}),
    # Cards for sales statistics
    html.Div([
        dbc.Row([
            dbc.Col(html.Div([
                dbc.Row([
                    dbc.Col(create_card("total-sales", "Total Sales", f"{int(sales_by_date['Sales'].sum()):,}€"), width=2),
                    dbc.Col(create_card("max-sales", "Max Sales", f"{int(max_sales):,}€"), width=2),
                    dbc.Col(create_card("min-sales", "Min Sales", f"{int(min_sales):,}€"), width=2),
                    dbc.Col(create_card("avg-sales", "Average Sales", f"{int(avg_sales):,}€"), width=2),
                    dbc.Col(create_card("std-dev-sales", "Standard Deviation", f"{int(std_dev_sales):,}"), width=2),
                ], justify="center"),
            ]), width=10, style={"margin": "auto"}) 
        ], style={"margin-top": "30px"}),
    ]),
    
    dcc.Graph(id='sales-graph'),
    dcc.Graph(id='monthly-stats-graph'),
    dcc.Graph(id='violin-plot'),
    dcc.Graph(id='kde-plot'),
    dcc.Interval(
        id='interval-component',
        interval=60000,  # in milliseconds, update every minute
        n_intervals=0
    )
])

# Sales over Time
@app.callback(
    Output('sales-graph', 'figure'),
    [Input('interval-component', 'n_intervals')]
)
def update_sales_graph(n):
    global sales_by_date
    
    # Create a Plotly figure
    fig = go.Figure()

    # Add a line trace for sales
    fig.add_trace(go.Scatter(x=sales_by_date['Date'], y=sales_by_date['Sales'], mode='lines', name='Sales', line=dict(width=1.2)))
    
    # Add lines for minimum, maximum, and average values
    fig.add_trace(go.Scatter(x=sales_by_date['Date'], y=[min_sales]*len(sales_by_date), mode='lines', name='Min Sales', line=dict(color='red', width=1, dash='solid')))
    fig.add_trace(go.Scatter(x=sales_by_date['Date'], y=[max_sales]*len(sales_by_date), mode='lines', name='Max Sales', line=dict(color='green', width=1, dash='solid')))
    fig.add_trace(go.Scatter(x=sales_by_date['Date'], y=[avg_sales]*len(sales_by_date), mode='lines', name='Average Sales', line=dict(color='orange', width=1, dash='solid')))

    # Calculate linear regression line
    x_values = np.arange(len(sales_by_date['Date']))
    slope, intercept = np.polyfit(x_values, sales_by_date['Sales'], 1)
    regression_y = slope * x_values + intercept
    
    # Add linear regression line trace
    fig.add_trace(go.Scatter(x=sales_by_date['Date'], y=regression_y, mode='lines', name='Linear Regression', line=dict(color='skyblue', width=1,dash='solid')))

    # Set layout options
    fig.update_layout(title='Sales Over Time',
                      title_x=0.5,  # Center the title
                      title_y=0.9,  # Adjust the vertical position of the title
                      title_font=dict(size=20),  # Set the font size of the title
                      yaxis_title='Sales (€)',
                      hovermode='x',
                      title_pad=dict(t=10, r=0, b=10, l=0),  # Adjust padding of the title
                      plot_bgcolor='rgba(0,0,0,0)',  # Dark background color
                      paper_bgcolor='rgba(0,0,0,0)',  # Dark background color
                      font_color='#FFFFFF',  # Light text color
                      margin=dict(l=300, r=300, t=100, b=10),  # Adjust the side margins
                      xaxis=dict(showgrid=False),  # Remove gridlines from the x-axis
                      yaxis=dict(showgrid=False)   # Remove gridlines from the y-axis
                      )

    return fig

# Monthly Sales, Costs and Profit
@app.callback(
    Output('monthly-stats-graph', 'figure'),
    [Input('interval-component', 'n_intervals')]
)
def update_monthly_stats_graph(n):
    global df_scp_monthly
    
    # Create a Plotly figure for monthly stats
    fig = go.Figure()

    # Add bar traces for sales, costs, and profit
    fig.add_trace(go.Bar(x=df_scp_monthly.index.to_timestamp(), y=df_scp_monthly['Sales'], name='Sales', marker=dict(line=dict(color='rgba(0,0,0,0)', width=0))))
    fig.add_trace(go.Bar(x=df_scp_monthly.index.to_timestamp(), y=df_scp_monthly['Costs'], name='Costs', marker=dict(line=dict(color='rgba(0,0,0,0)', width=0))))
    fig.add_trace(go.Bar(x=df_scp_monthly.index.to_timestamp(), y=df_scp_monthly['Profit'], name='Profit', marker=dict(line=dict(color='rgba(0,0,0,0)', width=0))))

    # Set layout options
    fig.update_layout(title='Monthly Sales, Costs, and Profit',
                      title_x=0.5,  # Center the title
                      title_y=0.9,  # Adjust the vertical position of the title
                      title_font=dict(size=20),  # Set the font size of the title
                      yaxis_title='Amount (€)',
                      barmode='group',  # Group the bars
                      hovermode='x',
                      plot_bgcolor='rgba(0,0,0,0)',  # Dark background color
                      paper_bgcolor='rgba(0,0,0,0)',  # Dark background color
                      font_color='#FFFFFF',  # Light text color
                      margin=dict(l=300, r=300, t=100, b=10),  # Adjust the side margins
                      )

    return fig

# Bikes Price Distribution by Color
@app.callback(
    Output('violin-plot', 'figure'),
    [Input('interval-component', 'n_intervals')]
)
def update_violin_plot(n):
    # Merge df_sales with df_products to get category and color information
    merged_df = pd.merge(df_sales, df_products, on='ProductKey', how='inner')
    
    # Filter data for bike category
    bike_sales = merged_df[merged_df['Category'] == 'Bikes']
    
    # Define the order of colors and corresponding colors
    color_order = ['Red', 'Silver', 'Black', 'Yellow', 'Blue']
    color_palette = ['#ff0000', '#e3e1d8', '#a1a1a1', '#f1c232', '#2f34fa']

    # Create a dictionary to map colors to categorical values
    color_map = {color_order[i]: color_palette[i] for i in range(len(color_order))}

    # Create a violin plot of bike category sales by color using Plotly Express
    fig = px.violin(bike_sales, x='Color', y='Sales', 
                    title='Bike Sales Price Distribution by Color',
                    color='Color', 
                    color_discrete_map=color_map)

    # Set layout options
    fig.update_layout(title='Bike Sales Price Distribution by Color',
                      title_x=0.5,  # Center the title
                      title_y=0.9,  # Adjust the vertical position of the title
                      title_font=dict(size=20),  # Set the font size of the title
                      yaxis_title='Price (€)',
                      xaxis_title='',
                      xaxis=dict(showticklabels=False),
                      barmode='group',  # Group the bars
                      hovermode='x',
                      plot_bgcolor='rgba(0,0,0,0)',  # Dark background color
                      paper_bgcolor='rgba(0,0,0,0)',  # Dark background color
                      font_color='#FFFFFF',  # Light text color
                      margin=dict(l=300, r=300, t=100, b=10)  # Adjust the side margins
                      )
    
    return fig

# Sales Price Distribution by Bike Type
@app.callback(
    Output('kde-plot', 'figure'),
    [Input('interval-component', 'n_intervals')]
)
def update_kde_plot(n):
    # Merge df_sales with df_products to get category information
    merged_df = pd.merge(df_sales, df_products, on='ProductKey', how='inner')
    
    # Filter sales data for selected subcategories
    selected_subcategories = ['Mountain Bikes', 'Road Bikes', 'Touring Bikes']
    filtered_sales = merged_df[merged_df['SubCategory'].isin(selected_subcategories)]
    
    # Create a KDE plot using Plotly Express
    fig = px.density_heatmap(filtered_sales, x='Sales', y='SubCategory', title='Sales Price Distribution by Bike Type')

    # Set layout options
    fig.update_layout(title='Sales Price Distribution by Bike Type',
                      title_x=0.5,  # Center the title
                      title_y=0.9,  # Adjust the vertical position of the title
                      title_font=dict(size=20),  # Set the font size of the title
                      yaxis_title='',
                      xaxis_title='Price(€)',
                      barmode='group',  # Group the bars
                      hovermode='x',
                      plot_bgcolor='rgba(0,0,0,0)',  # Dark background color
                      paper_bgcolor='rgba(0,0,0,0)',  # Dark background color
                      font_color='#FFFFFF',  # Light text color
                      margin=dict(l=300, r=300, t=100, b=100)  # Adjust the side margins
                      )
    
    return fig


# Run the app
if __name__ == '__main__':
    app.run_server(debug=True)
