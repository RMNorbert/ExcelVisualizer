#pip install pandas, openpyxl,streamlit, plotly-express
import pandas as pd
import streamlit as st
import plotly.express as px
import math
import time

if "show_visualization" not in st.session_state:
    st.session_state["show_visualization"] = False


def visualize():
    st.session_state["show_visualization"] = True
    global title
    title = st.session_state["file"]


def re_visualize():
    st.session_state["show_visualization"] = False
    #st.cache_data.clear


title = "ExcelVisualizer"
st.set_page_config(
    page_title=title,
    page_icon=":bar_chart:",
    layout="wide"
)


# Sidebar customization
if not st.session_state["show_visualization"]:
    header = st.sidebar.header("To start, set the options and click on the 'Visualize' button.")
    select_file = st.sidebar.text_input(
        label="Please insert the name of the file here:",
        placeholder="supermarkt_sales.xlsx",
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
    select_filters = st.sidebar.text_input(
        label="Please provide the name of the columns you would like to use for filtering:",
        placeholder="City,Gender",
        help="You can use any number of filter just separate them with: , "
    )
    visualize_button = st.sidebar.button(label="Visualize")

    if visualize_button:
        if "select_filters" not in st.session_state:
            st.session_state["file"] = select_file
            st.session_state["sheet"] = select_sheet
            st.session_state["skipped"] = select_rows_to_skip
            st.session_state["columns"] = select_columns
            st.session_state["shown"] = select_number_of_rows_to_show
            st.session_state["filters"] = select_filters
        visualize()


if st.session_state["show_visualization"]:
    time.sleep(0.5)
    # readingOptions
    array = st.session_state["filters"].split(',') if "," in st.session_state["filters"] else [
        st.session_state["filters"]]
    length = len(array)

    @st.cache_data  # store data in short term memory
    def get_data_from_excel():
        df = pd.read_excel(
            io=st.session_state["file"],  # excel file name eg:io='supermarkt_sales.xlsx',
            engine='openpyxl',  #
            sheet_name=st.session_state["sheet"],  # sheet name eg: sheet_name='Sales',
            skiprows=st.session_state["skipped"],  # how many rows to skip eg: skiprows=3,
            usecols=st.session_state["columns"],  # which columns to use eg: usecols='B:R',,
            nrows=st.session_state["shown"],  # how many rows include on the selection eg: nrows=1000,
        )

        if "Time" in df.columns:
            df["hour"] = pd.to_datetime(df["Time"], format="%H:%M:%S").dt.hour

        return df


    df = get_data_from_excel()
    header = st.sidebar.header("Please Filter Here:")
    #custom filter now by city
    filter1 = st.sidebar.multiselect(
                    "Select the " + array[0] + ":",
                    options=df[array[0]].unique(),
                    default=df[array[0]].unique()
                ) if not length == 1 else None
    if length >= 2:
        filter2 = st.sidebar.multiselect(
                "Select the " + array[1] + ":",
                options=df[array[1]].unique(),
                default=df[array[1]].unique()
            )
    if length >= 3:
        filter3 = st.sidebar.multiselect(
                "Select the " + array[2] + ":",
                options=df[array[2]].unique(),
                default=df[array[2]].unique()
            )

    query = ""

    for s in range(0, length, 1):
        query += array[s] + " == @"
        if s == 0:
            query += "filter1"
        if s == 1:
            query += "filter2"
        if s == 2:
            query += "filter3"
        if not s == length - 1:
            query += " & "
    # Generate the multiselect component

    df_selection = df.query(
        query
    )

    re_select = st.sidebar.button(label="Select another file")
    if re_select:
        re_visualize()

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