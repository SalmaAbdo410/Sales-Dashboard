import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import streamlit as st
import plotly.express as px
import plotly.graph_objects as go


# List your file names
files = ['2009.xlsx', '2010.xlsx', '2011.xlsx', '2012.xlsx']

# Read and append all files into one DataFrame
df = pd.concat([pd.read_excel(f) for f in files], ignore_index=True)

df['Order Date'] = pd.to_datetime(df['Order Date'], errors='coerce')
df['Ship Date'] = pd.to_datetime(df['Ship Date'], errors='coerce')

df = df.sort_values(by="Order Date", ascending=True)
df = df.sort_values(by="Ship Date", ascending=True)


df = df.drop_duplicates()

df['Province'] = df['Province'].replace({'Saskachewan': 'Saskatchewan'})
df['Region'] = df['Region'].replace({'Prarie': 'Prairie'})

cat_cols = [
    'Order Priority', 'Ship Mode', 'Customer Name', 'Province',
    'Region', 'Customer Segment', 'Product Category', 
    'Product Sub-Category', 'Product Name', 'Product Container'
]
for col in cat_cols:
    if col in df.columns:
        df[col] = df[col].astype('category')


df['Discount'] = df['Discount'].clip(lower=0, upper=1)

if 'Product Base Margin' in df.columns:
    df['Product Base Margin'] = pd.to_numeric(df['Product Base Margin'], errors='coerce')
    df['Product Base Margin'].fillna(df['Product Base Margin'].mean(), inplace=True)

df = df[df['Ship Date'] >= df['Order Date']]


df = df.sort_values(by='Order Date')

df = df.reset_index(drop=True)
df['Order Priority'] = df['Order Priority'].replace({'Not Specified': 'Medium'})

# Display cleaned data summary
print(df.info())
print(f"\nSample data:\n{df.head()}")
# Display the combined DataFrame
print(df)

df.to_excel('cleaned_data.xlsx', index=False)

st.set_page_config(layout="wide")
st.title("📊 Sales Dashboard (2009–2012)")
st.sidebar.header("📊 Sales Dashboard")
st.sidebar.image('photo.jpg', width=150)
st.sidebar.write("This dashboard provides a comprehensive view of sales data from 2009 to 2012.")
st.sidebar.subheader("Filters")
year_range = st.sidebar.slider(
    "Select Year Range",
    min_value=2009,
    max_value=2012,
    value=(2009, 2012) 
)

region_filter = st.sidebar.multiselect(
    "Select Region",
    options=df['Region'].unique(),
    default=df['Region'].unique())

category_filter = st.sidebar.multiselect(
    "Select Product Category",
    options=df['Product Category'].unique(),
    default=df['Product Category'].unique())


filtered_df = df[
    (df['Order Date'].dt.year >= year_range[0]) & 
    (df['Order Date'].dt.year <= year_range[1]) &
    (df['Region'].isin(region_filter)) &
    (df['Product Category'].isin(category_filter))
]

# Calculate KPIs
total_sales = filtered_df['Sales'].sum()
total_orders = filtered_df['Order ID'].nunique()
avg_profit_margin = filtered_df['Profit'].sum() / total_sales if total_sales > 0 else 0
avg_delivery_time = (filtered_df['Ship Date'] - filtered_df['Order Date']).dt.days.mean()

# Create 4 columns for KPIs
col1, col2, col3, col4 = st.columns(4)

with col1:
    st.metric(label="Total Sales", value=f"${total_sales:,.0f}")

with col2:
    st.metric(label="Total Orders", value=f"{total_orders:,}")

with col3:
    st.metric(label="Avg Profit Margin", value=f"{avg_profit_margin:.1%}")

with col4:
    st.metric(label="Avg Delivery Time", value=f"{avg_delivery_time:.1f} days")


tab1, tab2, tab3 = st.tabs(["Product Analysis", "Sales & Revenue", "Customer & Orders"])

with tab1:
    st.header("Product Analysis")

    # 1. Best-selling products by revenue
    top_products = filtered_df.groupby('Product Name')['Sales'].sum().sort_values(ascending=False).head(10).reset_index()
    fig1 = px.bar(
        top_products,
        x='Product Name',
        y='Sales',
        title='Top 10 Best-Selling Products by Revenue',
        labels={'Sales': 'Sales ($)', 'Product Name': 'Product'},
        color_discrete_sequence=['#c9b3e6']
    )
    st.plotly_chart(fig1, use_container_width=True)

    # 2. Sales by Product Category
    sales_by_category = filtered_df.groupby('Product Category')['Sales'].sum().reset_index()
    pastel_colors = ['#f8bbd0', '#c8e6c9', '#bbdefb']  
    fig2 = px.pie(
        sales_by_category,
        values='Sales',
        names='Product Category',
        title='Sales Distribution by Product Category',
        color_discrete_sequence=pastel_colors
    )
    st.plotly_chart(fig2, use_container_width=True)

    # 3. Profit margin by product category
    profit_margin = filtered_df.groupby('Product Category').apply(
        lambda x: x['Profit'].sum() / x['Sales'].sum() if x['Sales'].sum() > 0 else 0
    ).reset_index(name='Profit Margin')
    fig3 = px.bar(
        profit_margin,
        x='Product Category',
        y='Profit Margin',
        title='Profit Margin by Product Category',
        color_discrete_sequence=['#fff9c4']
    )
    st.plotly_chart(fig3, use_container_width=True)


with tab2:
    st.header("Sales & Revenue")

    # 1. Total sales over time (monthly)
    filtered_df['Order Month'] = filtered_df['Order Date'].dt.to_period('M').astype(str)
    monthly_sales = filtered_df.groupby('Order Month')['Sales'].sum().reset_index()
    fig4 = px.line(monthly_sales, x='Order Month', y='Sales', title='Monthly Sales Over Time', labels={'Order Month':'Month', 'Sales':'Sales ($)'})
    st.plotly_chart(fig4, use_container_width=True)

    # 2. Sales by Region
    sales_by_region = filtered_df.groupby('Region')['Sales'].sum().reset_index()
    fig5 = px.bar(
        sales_by_region,
        x='Region',
        y='Sales',
        title='Sales by Region',
        labels={'Sales':'Sales ($)', 'Region':'Region'},
        color_discrete_sequence=['#a9cce3']
    )
    st.plotly_chart(fig5, use_container_width=True)

    # 3. Discount impact on sales (scatter plot)
    fig6 = px.scatter(filtered_df, x='Discount', y='Sales', color='Region', title='Discount vs Sales')
    st.plotly_chart(fig6, use_container_width=True)


with tab3:
    st.header("Customer & Orders")

    # 1. Top customers by sales (Horizontal Bar Chart)
    top_customers = filtered_df.groupby('Customer Name')['Sales'].sum().sort_values(ascending=False).head(10).reset_index()
    fig7_hbar = px.bar(
        top_customers, x='Sales', y='Customer Name', orientation='h',
        title='Top 10 Customers by Sales (Horizontal Bar Chart)',
        labels={'Sales': 'Sales ($)', 'Customer Name': 'Customer'},
        color_discrete_sequence=['#fbb4ae']
    )
    st.plotly_chart(fig7_hbar, use_container_width=True)

    # 2. Average shipping time by region (Bar Chart)
    filtered_df['Shipping Time (Days)'] = (filtered_df['Ship Date'] - filtered_df['Order Date']).dt.days
    shipping_time = filtered_df.groupby('Region')['Shipping Time (Days)'].mean().reset_index()
    fig8 = px.bar(
        shipping_time,
        x='Region',
        y='Shipping Time (Days)',
        title='Average Shipping Time by Region',
        color_discrete_sequence=['#a8d5ba']
    )
    st.plotly_chart(fig8, use_container_width=True)

    # 3. Orders by priority
    orders_priority = filtered_df['Order Priority'].value_counts().reset_index()
    orders_priority.columns = ['Order Priority', 'Count']
    pastel_colors = ['#f8bbd0', '#c8e6c9', '#bbdefb', '#fff9c4']  # light pink, light green, light blue, light yellow
    fig9 = px.pie(
        orders_priority,
        values='Count',
        names='Order Priority',
        title='Order Priority Distribution',
        color_discrete_sequence=pastel_colors
    )
    st.plotly_chart(fig9, use_container_width=True)
    