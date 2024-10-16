import pandas as pd
import streamlit as st
import plotly.express as px
import datetime


st.set_page_config(page_title='Sales Analysis Dashboard')
st.header('Supermarket sales',divider="gray")

###--- LOAD DATA
excel_file = 'supermarket_sales.xlsx'
sheet_name = 'Sheet1'

df = pd.read_excel(excel_file,
                   sheet_name=sheet_name,
                   usecols='A:Q',
                   header=0)

df_product = pd.read_excel(excel_file,
                           sheet_name=sheet_name,
                           usecols='S:T',
                           header=0,
                           nrows= 6)
st.subheader("Data Set" )
st.dataframe(df)
st.markdown("City, Customer type, Product line, Payment method")
col1, col2, col3, col4, col5 = st.columns(5)
col1.dataframe(df['City'].unique().tolist())
col2.dataframe(df['Customer type'].unique().tolist())
col3.dataframe(df['Product line'].unique().tolist())
col4.dataframe(df['Payment'].unique().tolist())

pie_chart = px.pie(df_product,
                   title='Total No. of Product line',
                   values='Sum of Quantity',
                   names='Product line.1')

st.plotly_chart(pie_chart)
st.divider()

st.subheader("Daily Sales Analysis")
productLine = df['Product line'].unique().tolist()
date = df['Date'].unique().tolist()
date_as_datetime = pd.to_datetime(date).date
year = 2019
date_selection = st.date_input('Choose a day range',
                                (min(date_as_datetime), max(date_as_datetime)),
                                datetime.date(year, 1, 1),
                                datetime.date(year, 3, 30),
                                format="DD.MM.YYYY")



department_selection = st.multiselect('Product line',
                                      productLine,
                                      default=productLine)


## กรองข้อมูลตามที่เลือก
mask = ((df['Date'].dt.date >= date_selection[0]) & (df['Date'].dt.date <= date_selection[1])) & (df['Product line'].isin(department_selection))
numbe_of_result = df[mask].shape[0]
st.markdown(f'Available Results: {numbe_of_result}*')

# จัดกลุ่ม data
df_grouped = df[mask].groupby(by=['Product line']).count()[['Total']]
df_grouped = df_grouped.reset_index()
# สร้าง bar chart
bar_chart_Dairy = px.bar(df_grouped,
                   x='Product line',
                   y='Total',
                   text='Total',
                   color_discrete_sequence = ['#F63366']*len(df_grouped),
                   template= 'plotly_white')
st.plotly_chart(bar_chart_Dairy)

st.divider()

st.subheader("Branch Sales Analysis")

if mask.any():  # ตรวจสอบว่ามีข้อมูลที่กรองได้
    # สร้าง df_grouped ที่รวม gross income ตาม Branch และ Product line
    df_grouped = df[mask].groupby(['Branch', 'Product line'])['gross income'].sum().reset_index()

    # สร้าง pie chart สำหรับแต่ละ Branch
    for branch in df_grouped['Branch'].unique():
        branch_data = df_grouped[df_grouped['Branch'] == branch]
        total_gross_income = branch_data['gross income'].sum()  # คำนวณ gross income รวม

        # สร้าง pie chart 
        pie_chart = px.pie(
            branch_data,
            names='Product line',
            values='gross income',
            title=f'Gross Income of Product Lines in Branch {branch} (Total Gross Income: {total_gross_income.round(2)})',
            template='plotly_white'
        )

        st.plotly_chart(pie_chart)
else:
    st.markdown('No results found for the selected criteria.')

st.divider()

# --- Customer Satisfaction Analysis ---
st.subheader("Customer Satisfaction Analysis")
if 'Rating' in df.columns:  # ตรวจสอบว่ามีคอลัมน์ Rating หรือไม่
    # คำนวณค่าเฉลี่ยของคะแนนความพึงพอใจตามประเภทสินค้า
    satisfaction_grouped = df[mask].groupby(['Product line'])['Rating'].mean().reset_index()
    
    # ตัดทศนิยม 1 ตำแหน่ง
    satisfaction_grouped['Rating'] = satisfaction_grouped['Rating'].round(1)

    satisfaction_grouped.columns = ['Product line', 'Average Rating']  # เปลี่ยนชื่อคอลัมน์

    # สร้าง bar chart ของค่าเฉลี่ยคะแนนความพึงพอใจ
    bar_chart_satisfaction = px.bar(satisfaction_grouped, x='Product line', y='Average Rating', text='Average Rating',
                                     color_discrete_sequence=['#4287f5'] * len(satisfaction_grouped),
                                     template='plotly_white',
                                     title='Average Customer Satisfaction Ratings by Product Line')
    st.plotly_chart(bar_chart_satisfaction)
else:
    st.markdown('No customer satisfaction data available.')

