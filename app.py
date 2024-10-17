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
st.subheader("Data Set Overview" )
st.dataframe(df)

pie_chart = px.pie(df_product,
                   title='Total No. of Product line',
                   values='Sum of Quantity',
                   names='Product line.1')

st.plotly_chart(pie_chart)
st.divider()
page = st.selectbox('Choose a page', ['Period Sales Analysis','Dairy Sales Analysis'])

if page == 'Period Sales Analysis':
    st.subheader("Period Sales Analysis")
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
    try:
        mask = ((df['Date'].dt.date >= date_selection[0]) & (df['Date'].dt.date <= date_selection[1])) & (df['Product line'].isin(department_selection))
        numbe_of_result = df[mask].shape[0]
        st.markdown(f'Available Results: {numbe_of_result}*')
    

        # จัดกลุ่ม data
        df_grouped = df[mask].groupby(by=['Product line'])['Total'].sum().round(2).reset_index()
        df_total =  df[mask]['Total'].sum().round(2)
        # สร้าง bar chart
        bar_chart = px.bar(df_grouped,
                                x='Product line',
                                y='Total',
                                text='Total',
                                color='Product line',  # ใช้ color แยกสีตามประเภทสินค้า
                                template='plotly_white',
                                title=f'Total Sales : {df_total} USD')

        st.plotly_chart(bar_chart)

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

        st.subheader("Customer Satisfaction Analysis")
        # --- Customer Satisfaction Analysis ---
        customerType = df['Customer type'].unique().tolist()

        # ให้ผู้ใช้เลือกประเภทลูกค้า
        customerType_selection = st.multiselect('Customer Type', 
                                                customerType, 
                                                default=customerType)

        # สร้าง mask เพื่อกรองข้อมูลตามประเภทลูกค้า
        customer_mask = (df['Customer type'].isin(customerType_selection))
        customer_filtered_df = df[mask & customer_mask]

        if 'Rating' in customer_filtered_df.columns:  # ตรวจสอบว่ามีคอลัมน์ Rating หรือไม่
            # คำนวณค่าเฉลี่ยของคะแนนความพึงพอใจตามประเภทสินค้าและประเภทลูกค้า
            satisfaction_grouped = customer_filtered_df.groupby(['Product line', 'Customer type'])['Rating'].mean().reset_index()

            # ตัดทศนิยม 1 ตำแหน่ง
            satisfaction_grouped['Rating'] = satisfaction_grouped['Rating'].round(1)
            satisfaction_grouped.columns = ['Product line', 'Customer type', 'Average Rating']  # เปลี่ยนชื่อคอลัมน์

            # สร้าง bar chart ของค่าเฉลี่ยคะแนนความพึงพอใจ แยกสีตามประเภทลูกค้า
            bar_chart_satisfaction = px.bar(satisfaction_grouped, 
                                            x='Product line', 
                                            y='Average Rating', 
                                            color='Customer type',  # แยกสีตามประเภทลูกค้า
                                            text='Average Rating',
                                            barmode='group',  # แสดงแต่ละกลุ่มในแผนภูมิเดียวกัน
                                            color_discrete_sequence=px.colors.qualitative.Set1,  # กำหนดสีของกลุ่ม
                                            template='plotly_white',
                                            title=f'Average Customer Satisfaction Ratings for {", ".join(customerType_selection)}')
            st.plotly_chart(bar_chart_satisfaction)
        else:
            st.markdown('No customer satisfaction data available.')

        st.divider()
        
        sales_groupped = df[mask].groupby(['Date'])['Total'].sum().reset_index()
        sales_total = df[mask]['Total'].sum().round(2)
        line_chart_dairy = px.line(sales_groupped,
                            x = 'Date',
                            y= 'Total',
                            title=f'Dairly sale (Total Sale: {sales_total} USD)')
        
        st.plotly_chart(line_chart_dairy)

        st.divider()
        st.subheader("Summary")

        total_sales = df[mask]['Total'].sum().round(2)
        total_income = df[mask]['gross income'].sum().round(2)
        total_rating = df[mask]['Rating'].sum().round(2)

        col1, col2, col3 = st.columns(3)
        col1.metric("Total sales",f'{total_sales} $')
        col2.metric("gross income", f'{total_income} $')
        col3.metric("Rating", f'{total_rating}')

    except Exception as e:
        st.error("An error occurred: Please check your input.")
        
elif page == 'Dairy Sales Analysis':
    st.subheader("Dairy Sales Analysis")

    date_input = st.date_input("Select date", datetime.date(2019, 1, 1), datetime.date(2019, 1, 1), datetime.date(2019, 3, 30) , format="DD.MM.YYYY")
    
                                    
    date_mask = (df['Date'].dt.date == date_input)
    previous_day_mask = (df['Date'].dt.date == (date_input - datetime.timedelta(days=1)))
    

    numbe_of_result = df[date_mask].shape[0]
    st.markdown(f'Available Results: {numbe_of_result}*')

    sales_quantity = df[date_mask]['Quantity'].sum() 
    line_chart_dairy = px.line(df[date_mask].sort_values(by='Time') ,
                         x = 'Time',
                         y= 'Quantity',
                         title=f'Dairly sale (Total Sale: {sales_quantity})')
    
    st.plotly_chart(line_chart_dairy)

    st.divider()
    st.subheader("Sale Quantity by Product Line")
    df_grouped_by_product = df[date_mask].groupby(by=['Product line'])['Quantity'].sum().reset_index()
    
    # สร้าง bar chart
    bar_chart_Product = px.bar(df_grouped_by_product,
                            x='Product line',
                            y='Quantity',
                            text='Quantity',
                            color='Product line',  # ใช้ color แยกสีตามประเภทสินค้า
                            template='plotly_white',
                            title='Sale Quantity')

    st.plotly_chart(bar_chart_Product)
    st.divider()
    st.subheader("Group by")
    pie1, pie2 = st.columns(2)

    df_grouped_by_payment = df[date_mask].groupby(by=['Payment']).count()[['Total']].reset_index()

    pie_chart = px.pie(
        df_grouped_by_payment,
        names='Payment',
        values='Total',
        title=f'Payment type',
        template='plotly_white'
    )

    pie1.plotly_chart(pie_chart)

    df_grouped_by_payment = df[date_mask].groupby(by=['Customer type']).count()[['Total']].reset_index()

    pie_chart = px.pie(
        df_grouped_by_payment,
        names='Customer type',
        values='Total',
        title=f'Customer Type',
        template='plotly_white'
    )

    pie2.plotly_chart(pie_chart)

    total_sales = df[date_mask]['Total'].sum().round(2)
    total_income = df[date_mask]['gross income'].sum().round(2)
    total_rating = df[date_mask]['Rating'].sum().round(2)

    previous_sales = df[previous_day_mask]['Total'].sum().round(2)
    previous_income = df[previous_day_mask]['gross income'].sum().round(2)
    previous_rating = df[previous_day_mask]['Rating'].sum().round(2)

    total_change = (((total_sales - previous_sales) / previous_sales) * 100).round(2)
    rating_change = (((total_rating - previous_rating) / previous_rating) * 100).round(2)
    
    st.divider()
    st.subheader("Summary")
    col1, col2, col3 = st.columns(3)
    col1.metric("Total sales",f'{total_sales} USD', total_change)
    col2.metric("gross income", f'{total_income} USD', total_change)
    col3.metric("Rating", f'{total_rating}', rating_change)