#pip install pandas, openpyxl,streamlit, plotly-express
import pandas as pd
import streamlit as st
import plotly.express as px
import math


show_visualization = False


def visualize():
    global show_visualization
    show_visualization = True


title = "ExcelVisualizer"
st.set_page_config(
    page_title=title,
    page_icon=":bar_chart:",
    layout="wide"
)

# Sidebar customization
if not show_visualization:
    st.sidebar.header("To start, set the options and click the 'Visualize' button.")
    select_file = st.sidebar.text_input(
        label="Please insert the name of the file here:",
        placeholder="example.xlsx",
        help="Currently, only XML files supported which located next to the main.py file."
    )
    select_sheet = st.sidebar.text_input(
        label="Please enter the name of the sheet to visualize:",
        placeholder="Example",
        help="Currently, only one sheet can be visualized."
    )
    select_rows_to_skip = st.sidebar.number_input(
        label="Please provide the number of rows to skip:",
        min_value=0,
        help="The default value is 0. The count of rows starts from the top. "
             "The provided number of rows will not be processed during the visualization."
    )
    select_columns = st.sidebar.text_input(
        label="Please enter the columns to visualize:",
        placeholder="B:R",
        help="The default value will contain A:B."
    )
    select_number_of_rows_to_show = st.sidebar.number_input(
        label="Please provide the number of rows to show:",
        min_value=1,
        help="The default value is 1. The number of rows that will be used during the visualization. "
             "The count of rows starts from the top."
    )

    visualize_button = st.sidebar.button(label="Visualize")

    if visualize_button:
        visualize()


if show_visualization:
    # readingOptions

#    @st.cache_data  # store data in short term memory
    def get_data_from_excel():
        df = pd.read_excel(
            io=select_file,  # excel file name eg:io='supermarkt_sales.xlsx',
            engine='openpyxl',  #
            sheet_name=select_sheet,  # sheet name eg: sheet_name='Sales',
            skiprows=select_rows_to_skip,  # how many rows to skip eg: skiprows=3,
            usecols=select_columns,  # which columns to use eg: usecols='B:R',,
            nrows=select_number_of_rows_to_show,  # how many rows include on the selection eg: nrows=1000,
        )

        if "Time" in df.columns:
            df["hour"] = pd.to_datetime(df["Time"], format="%H:%M:%S").dt.hour

        return df


    df = get_data_from_excel()

    st.sidebar.header("Please Filter Here:")
    #custom filter now by city
    city = st.sidebar.multiselect(
        "Select the City:",
        options=df["City"].unique(),
        default=df["City"].unique()
    )

    customer_type = st.sidebar.multiselect(
        "Select the Customer Type:",
        options=df["Customer_type"].unique(),
        default=df["Customer_type"].unique()
    )

    gender = st.sidebar.multiselect(
        "Select the Gender:",
        options=df["Gender"].unique(),
        default=df["Gender"].unique()
    )

    df_selection = df.query(
        "City == @city & Customer_type == @customer_type & Gender == @gender"
    )
#print(df_selection)

#main_page
    st.title(":bar_chart: " + title)
    st.markdown("##")

    total_sales = int(df_selection["Total"].sum())
    average_rating = round(df_selection["Rating"].mean(), 1)
    star_rating = ":star:" * int(round(average_rating, 0)) if not math.isnan(average_rating) else ""
    average_sale_by_transaction = round(df_selection["Total"].mean(), 2)

    left_column, middle_column, right_column = st.columns(3)
    with left_column:
        st.subheader("Total Sales:")
        st.subheader(f"US $ {total_sales:,}")
    with middle_column:
        st.subheader("Average Rating:")
        st.subheader(f"{average_rating} {star_rating}")
    with right_column:
        st.subheader("Average Sales Per Transaction:")
        st.subheader(f"US $ {average_sale_by_transaction}")

    st.markdown("""---""")

#bar_charts

    sales_by_product_line = (
        df_selection
        .groupby(by=["Product line"])["Total"]
        .sum()
        #.reset_index() to show as index
        #.sort_values("Total")
    )
    fig_product_sales = px.bar(
        sales_by_product_line,
        x="Total",
        y=sales_by_product_line.index,
        orientation="h", #horizontal_bar_chart
        title="<b>Sales by Product Line",
        color_discrete_sequence=["#0083B8"] * len(sales_by_product_line),
        template="plotly_white", #template_style
    )
    fig_product_sales.update_layout(
        plot_bgcolor="rgba(0,0,0,0)",
        xaxis=(dict(showgrid=False))
    )

# SALES BY HOUR [BAR CHART]
    sales_by_hour = df_selection.groupby(by=["hour"])[["Total"]].sum()
    fig_hourly_sales = px.bar(
        sales_by_hour,
        x=sales_by_hour.index,
        y="Total",
        title="<b>Sales by hour</b>",
        color_discrete_sequence=["#0083B8"] * len(sales_by_hour),
        template="plotly_white",
    )
    fig_hourly_sales.update_layout(
        xaxis=dict(tickmode="linear"),
        plot_bgcolor="rgba(0,0,0,0)",
        yaxis=(dict(showgrid=False)),
    )


    left_column, right_column = st.columns(2)
    left_column.plotly_chart(fig_hourly_sales, use_container_width=True)
    right_column.plotly_chart(fig_product_sales, use_container_width=True)


hide_st_style = """
            <style>
            #MainMenu {visibility: hidden;}
            footer {visibility: hidden;}
            header {visibility: hidden;}
            </style>
            """
st.markdown(hide_st_style, unsafe_allow_html=True)