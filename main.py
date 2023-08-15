import pandas as pd
import streamlit as st
import plotly.express as px
import time

if 'show_visualization' not in st.session_state:
    st.session_state["show_visualization"] = False

if st.session_state.get('count') is None:
    st.session_state.count = 0


def setStates():
    st.session_state["file"] = select_file
    st.session_state["sheet"] = select_sheet
    st.session_state["skipped"] = select_rows_to_skip
    st.session_state["columns"] = select_columns
    st.session_state["shown"] = select_number_of_rows_to_show
    st.session_state["filters"] = select_filters
    st.session_state["total"] = select_total
    st.session_state["unit"] = select_unit
    st.session_state["group_by"] = select_group
    st.session_state.count += 1


def visualize():
    if "select_filters" not in st.session_state:
        setStates()
    global title
    title = st.session_state["file"]
    st.session_state["show_visualization"] = True
    st.session_state.count += 1


def return_to_options():
    st.session_state["show_visualization"] = False
    st.session_state.count = 0
    st.cache_data.clear()


title = "ExcelVisualizer"
st.set_page_config(
    page_title=title,
    page_icon=":bar_chart:",
    layout="wide"
    )


# Sidebar customization
if not st.session_state["show_visualization"]:
    if st.session_state.count == 0:
        header = st.header("To start, set the options and click on the 'Visualize' button.")
        select_file = st.text_input(
            label="Please insert the name of the file here:",
            placeholder="supermarkt_sales.xlsx",
            help="Currently, only files supported which located next to the main.py file.",
            value="supermarkt_sales.xlsx"
        )
        select_sheet = st.text_input(
            label="Please enter the name of the sheet to visualize:",
            placeholder="Sales",
            help="Currently, only one sheet can be visualized.",
            value="Sales"
        )
        select_rows_to_skip = st.number_input(
            label="Please provide the number of rows to skip:",
            min_value=0,
            help="The default value is 3. The count of rows starts from the top. "
                 "The provided number of rows will not be processed during the visualization.",
            value=3
        )
        select_columns = st.text_input(
            label="Please enter the columns to visualize:",
            placeholder="B:R",
            help="The default value will contain B:R.",
            value="B:R"
        )
        select_number_of_rows_to_show = st.number_input(
            label="Please provide the number of rows to show:",
            min_value=1,
            help="The default value is 1. The number of rows that will be used during the visualization. "
                 "The count of rows starts from the top.",
            value=100
        )
        select_filters = st.text_input(
            label="Please provide the name of the columns you would like to use for filtering:",
            placeholder="City,Gender",
            help="You can use any number of filter just separate them with: , ",
            value="City,Gender"
        )
        select_total = st.text_input(
            label="Please provide the name of the column to calculate the total value:",
            placeholder="Total",
            help="""You can one column only for the calculation. The base value will be the "Total" column """,
            value="Total"
        )
        select_unit = st.text_input(
            label="Please provide the base of the financial, thermochemical or other unit for the calculations:",
            placeholder="US $ or kcal, ect.",
            help="""You can give only one unit type for the calculations. The base value will be the "US $ """,
            value="US $"
        )
        select_group = st.text_input(
            label="Please provide the column to use to group by calculation:",
            placeholder="Product line",
            help="""You can give only one column name. The base value will be Product line""",
            value="Product line"
        )
        visualize_button = st.button(label="Visualize", on_click=visualize)


if st.session_state["show_visualization"]:
    if st.session_state.count != 0:
        time.sleep(0.5)
        # variables set for repeated use
        array = st.session_state["filters"].split(',') if "," in st.session_state["filters"] else [st.session_state["filters"]]
        length = len(array)
        total = st.session_state["total"]
        group_by = st.session_state["group_by"]

        # readingOptions
        # read engine types:
        # openpyxl supports newer Excel file formats
        # for older use xlrd,
        # odf for OpenDocument formats
        # pyxlsb for Binary excel files

        @st.cache_data  # store data in short term memory
        def get_data_from_excel():
            df = pd.read_excel(
                io=st.session_state["file"],  # excel file name eg:io='supermarkt_sales.xlsx',
                engine='openpyxl',  # openpyxl for newer Excel file formats
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
        # custom filter

        filter1 = st.sidebar.multiselect(
                        "Select the " + array[0] + ":",
                        options=df[array[0]].unique(),
                        default=df[array[0]].unique()
                    ) if length >= 1 else None

        filter2 = st.sidebar.multiselect(
                    "Select the " + array[1] + ":",
                    options=df[array[1]].unique(),
                    default=df[array[1]].unique()
                    ) if length >= 2 else None

        filter3 = st.sidebar.multiselect(
                    "Select the " + array[2] + ":",
                    options=df[array[2]].unique(),
                    default=df[array[2]].unique()
                    ) if length == 3 else None

        # Generate the multiselect component
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

        df_selection = df.query(
            query
        )

        re_select = st.button(label="Select another file", on_click=return_to_options)

    # main_page
        st.title(":bar_chart: " + title)
        st.markdown("##")
        total_sum = int(round(df_selection[total].sum()))
        reviewed = True if "Rating" in df_selection.columns.values else False
        if reviewed:
            try:
                average_rating = round(df_selection["Rating"].mean(), 1)
                star_rating = ":star:" * int(round(average_rating, 0))
            except KeyError:
                average_rating = ""
                star_rating = ""

        average_value = round(df_selection[total].mean(), 2)

        left_column, middle_column, right_column = st.columns(3)
        with left_column:
            f = st.session_state.get("unit")
            t = st.subheader("Sum of " + total)
            st.subheader(f"{f} {total_sum:,}")
        with middle_column:
            if reviewed:
                st.subheader("Average Rating:")
                st.subheader(f"{average_rating} {star_rating}")
        with right_column:
            st.subheader("Average Value:")
            st.subheader(f"{f} {average_value}")

        st.markdown("""---""")

# bar_charts

        values_by_group = (
            df_selection
            .groupby(by=[group_by])[total]
            .sum()
            #  .reset_index() to show as index
            #  .sort_values("Total")
        )
        fig_values = px.bar(
            values_by_group,
            x=total,
            y=values_by_group.index,
            orientation="h",  # horizontal_bar_chart
            title="<b>Value by " + group_by,
            color_discrete_sequence=["#0083B8"] * len(values_by_group),
            template="plotly_white",  # template_style
        )
        fig_values.update_layout(
            plot_bgcolor="rgba(0,0,0,0)",  # set  background color of the entire plot
            xaxis=(dict(showgrid=False))  # specifies that gridlines along the x-axis should be hidden
        )

    # VALUES BY HOUR [BAR CHART]
        left_column, right_column = st.columns(2)
        right_column.plotly_chart(fig_values, use_container_width=True)
        if "hour" in df.columns:
            values_by_hour = df_selection.groupby(by=["hour"])[[total]].sum()
            fig_hourly_values = px.bar(
                values_by_hour,
                x=values_by_hour.index,
                y=total,
                title="<b>Values by hour</b>",
                color_discrete_sequence=["#0083B8"] * len(values_by_hour),
                template="plotly_white",
            )
            fig_hourly_values.update_layout(
                xaxis=dict(tickmode="linear"),
                plot_bgcolor="rgba(0,0,0,0)",
                yaxis=(dict(showgrid=False)),
            )

            left_column.plotly_chart(fig_hourly_values, use_container_width=True)

hide_st_style = """
            <style>
            #MainMenu {visibility: hidden;}
            .main {margin-top: -85px !important;}
            .element-container > div {margin-top: -10px !important;}
            footer {visibility: hidden;}
            header {visibility: hidden;}
            </style>
            """
st.markdown(hide_st_style, unsafe_allow_html=True)
