import pandas as pd
import numpy as np
import random
import plotly.express as px
import dash
from dash import dcc, html
from dash.dependencies import Input, Output

# Generate Sample Data
categories = ['Electronics', 'Clothing', 'Home & Kitchen', 'Sports', 'Books']
months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']
data = {
    'Category': [random.choice(categories) for _ in range(1000)],
    'Month': [random.choice(months) for _ in range(1000)],
    'Sales': np.random.randint(100, 5000, 1000),
    'Profit': np.random.randint(20, 1000, 1000)
}
df = pd.DataFrame(data)

# Save to Excel
excel_file = "sales_data.xlsx"
df.to_excel(excel_file, index=False)

# Read Excel Data
df = pd.read_excel(excel_file)

# Create Dash App
app = dash.Dash(__name__)
app.layout = html.Div([
    html.H1("Excel Data Visualization Dashboard", style={'textAlign': 'center'}),
    
    dcc.Graph(id='bar-chart'),
    dcc.Graph(id='line-graph'),
    dcc.Graph(id='scatter-plot'),
    dcc.Graph(id='pie-chart')
])

# Update Graphs
@app.callback(
    Output('bar-chart', 'figure'),
    Output('line-graph', 'figure'),
    Output('scatter-plot', 'figure'),
    Output('pie-chart', 'figure'),
    Input('bar-chart', 'id')
)
def update_graphs(_):
    bar_fig = px.bar(df.groupby('Category', as_index=False)['Sales'].sum(), x='Category', y='Sales', title='Total Sales by Category')
    line_fig = px.line(df.groupby('Month', as_index=False)['Sales'].sum(), x='Month', y='Sales', title='Monthly Sales Trend')
    scatter_fig = px.scatter(df, x='Sales', y='Profit', title='Sales vs Profit', trendline='ols')
    pie_fig = px.pie(df.groupby('Month', as_index=False)['Sales'].sum(), names='Month', values='Sales', title='Sales Distribution by Month')
    return bar_fig, line_fig, scatter_fig, pie_fig

# Run Server
if __name__ == '__main__':
    app.run(debug=True)

    