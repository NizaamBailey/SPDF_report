import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px

def check_uploaded_file(uploaded_file):
    if uploaded_file is None:
        st.warning('Upload Excel File!')
        st.stop()

def read_excel_file(uploaded_file):
    df = pd.read_excel(uploaded_file, sheet_name=None)
    return df

def check_required_sheets(df):
    required_sheets = ['Customer Information', 'Incident Data', 'Availability Data']
    if not all(sheet in df for sheet in required_sheets):
        st.warning("Excel file doesn't contain the required sheets!")
        st.stop()

def extract_dataframes(df):
    df_Cust_Info = df['Customer Information']
    df_Incidents = df['Incident Data']
    df_Avail = df['Availability Data']
    df_adhoc = df['Ad hoc Data']
    return df_Cust_Info, df_Incidents, df_Avail, df_adhoc

def diplay_cust_info(df_Cust_Info):
    month = df_Cust_Info['Date'].iloc[0]
    customer_name = df_Cust_Info['Customer Name'].iloc[0]
    service_managers = df_Cust_Info['Service Manager']

    st.markdown(f"<h1 style='text-align: center;'>📊 Service Portfolio Overview</h1>", unsafe_allow_html=True)
    st.subheader(":date:" + " " + str(month))
    st.markdown("**Customer Name**: " + str(customer_name))
    st.markdown("**Service Manager:**" + " " + ", ".join(service_managers))
    st.markdown("""---""")

def calculate_availability_metrics(df_Avail):
    total_sites = int(df_Avail["Site"].count())
    missed_sla_count = df_Avail[df_Avail['Actual Site Availability'] < df_Avail['Target Availability (SLA)']]['Site'].count()
    above_target_count = df_Avail[df_Avail['Actual Site Availability'] >= df_Avail['Target Availability (SLA)']]['Site'].count()
    return total_sites, missed_sla_count, above_target_count

def display_kpis(total_sites, missed_sla_count, above_target_count):
    kpi_style = "font-size: 20px; font-weight: bold;"
    st.markdown("<h2 style='text-align: center;'> Top KPIs </h2>", unsafe_allow_html=True)
    st.markdown('#')

    col1, col2, col3 = st.columns(3)

    with col1:
        st.markdown("<p style='" + kpi_style + "'>No. of Sites that Met SLA</p>", unsafe_allow_html=True)
        st.markdown("<p style='" + kpi_style + "'>" + str(above_target_count) + "</p>", unsafe_allow_html=True)

    with col2:
        st.markdown("<p style='" + kpi_style + "'>No. of Sites That Missed SLA</p>", unsafe_allow_html=True)
        st.markdown("<p style='" + kpi_style + "'>" + str(missed_sla_count) + "</p>", unsafe_allow_html=True)

    with col3:
        st.markdown("<p style='" + kpi_style + "'>No. of Sites Reported On</p>", unsafe_allow_html=True)
        st.markdown("<p style='" + kpi_style + "'>" + str(total_sites) + "</p>", unsafe_allow_html=True)

def calculate_incident_metrics(df_Incidents):
    customer_caused = ['Power failure – Customer site', 'Client power issue', 'Customer Caused Interruptions', 'Power', 'Power Failure', 'Fault Is On Customer Side', 'Power disconnect – Cust Edge d', 'disconnect – Customer Edge dev', 'Customer Interruption', 'Customer Interruption – Mainte']
    vodacom_caused = ['Source site Power failure', 'Power failure – Vodacom site', 'Database outage', 'Power failure – Vodacom site', 'Power failure - Vodacom site']

    total_incidents = df_Incidents.shape[0]
    total_customer_caused = df_Incidents[df_Incidents['Root Cause'].str.lower().isin([x.lower() for x in customer_caused])].shape[0]
    total_vodacom_caused = df_Incidents[df_Incidents['Root Cause'].str.lower().isin([x.lower() for x in vodacom_caused])].shape[0]
    total_closed = df_Incidents["Status"].value_counts().get('Closed', 0)
    open_incidents = df_Incidents["Status"].value_counts().get('Open', 0)
    return total_incidents, total_customer_caused, total_vodacom_caused, total_closed, open_incidents

def display_incident_metrics(total_incidents, total_customer_caused, total_vodacom_caused, total_closed, open_incidents):
    overview_style = "font-size: 16px; font-weight: bold; margin-top: 20px;"

    col1, col2, col3, col4, col5 = st.columns(5)

    with col1:
        st.markdown("<p style='" + overview_style + "'>No. of Incidents Logged this Period</p>", unsafe_allow_html=True)
        st.markdown("<p style='" + overview_style + "'>" + str(total_incidents) + "</p>", unsafe_allow_html=True)

    with col2:
        st.markdown("<p style='" + overview_style + "'>No. of Customer Caused Incidents</p>", unsafe_allow_html=True)
        st.markdown("<p style='" + overview_style + "'>" + str(total_customer_caused) + "</p>", unsafe_allow_html=True)

    with col3:
        st.markdown("<p style='" + overview_style + "'>No. of Incidents Closed</p>", unsafe_allow_html=True)
        st.markdown("<p style='" + overview_style + "'>" + str(total_closed) + "</p>", unsafe_allow_html=True)

    with col4:
        st.markdown("<p style='" + overview_style + "'>No. of Incidents still Open</p>", unsafe_allow_html=True)
        st.markdown("<p style='" + overview_style + "'>" + str(open_incidents) + "</p>", unsafe_allow_html=True)

    with col5:
        st.markdown("<p style='" + overview_style + "'>No. of Vodacom Caused Incidents</p>", unsafe_allow_html=True)
        st.markdown("<p style='" + overview_style + "'>" + str(total_vodacom_caused) + "</p>", unsafe_allow_html=True)

def display_top_problematic_sites(df_Incidents):
    df_prob = df_Incidents.groupby(['Site Name'], as_index=False)['Root Cause'].count()
    df_prob = df_prob.sort_values(by='Root Cause', ascending=False)
    df_prob = df_prob.reset_index(drop=True)
    df_prob.index = df_prob.index + 1
    df_prob.columns = ['Site Name', 'Number of Incidents']

    st.subheader("Top 10 Problematic Sites:")
    st.dataframe(df_prob.head(10))

def display_site_availability_report(df_Cust_Info, df_Avail):
    st.subheader("Site Availability Report:" + " " + df_Cust_Info["Date"][0])

    columns_to_format_percentage = ['Target Availability (SLA)', 'Primary Gross Availability', 'Secondary Gross Availability', 'Non-Contractual Downtime', 'Actual Site Availability']
    df_Avail[columns_to_format_percentage] = df_Avail[columns_to_format_percentage].mul(100).round(2).astype(str) + '%'

    columns_to_format_NA = ['Comment', 'Primary Solution ID', 'Secondary Solution ID', 'Site']
    df_Avail[columns_to_format_NA] = df_Avail[columns_to_format_NA].fillna('').astype(str)

    df_Avail.index = np.arange(1, len(df_Avail) + 1)

    def highlight_less_than_target(val):
        target = df_Avail.loc[val.index, 'Target Availability (SLA)']
        val_float = val.str.rstrip('%').astype(float)
        target_float = target.str.rstrip('%').astype(float)
        return ['background-color: salmon' if x < y else '' for x, y in zip(val_float, target_float)]

    df_styled = df_Avail.style.apply(highlight_less_than_target, subset='Actual Site Availability', axis=0)

    st.table(df_styled.set_properties(**{'font-size': '12px'}))

def display_3_month_availability_trend(df_adhoc):
    st.subheader('Site Availability – 3 Months Trend')

    decimal_columns = df_adhoc.select_dtypes(include=[np.number]).columns
    df_adhoc[decimal_columns] = df_adhoc[decimal_columns].applymap('{:.2%}'.format)

    df_adhoc_styled = df_adhoc.style.set_properties(**{'font-size': '12px'})

    st.table(df_adhoc_styled)

def display_incident_report(df_Cust_Info, df_Incidents):
    st.subheader("Incident Report:" + " " + df_Cust_Info["Date"][0])

    columns_to_display = ["SR Number", "Site Name", "Root Cause", "SR Type", "Agent Priority", "Time Logged", "Time Resolved", "Resolution", "Source", "Actual Duration(DD:HH:MM:SS)", "Summary"]
    df_Incidents_selected = df_Incidents[columns_to_display]

    # Styling the DataFrame
    df_styled = df_Incidents_selected.style.set_properties(**{'font-size': '11px'}).set_table_styles([
        {'selector': 'table', 'props': [('display', 'none')]},
        {'selector': '.row_heading, .blank', 'props': [('display', 'none')]}
    ])

    # Convert the Styler object to HTML
    df_html = df_styled.render()

    # Display the HTML in Streamlit
    st.write(df_html, unsafe_allow_html=True)





def display_root_cause_count(df_Incidents):
    st.markdown('---')
    st.markdown("<h2 style='text-align: center;'> Analysis </h2>", unsafe_allow_html=True)
    
    root_cause_counts = df_Incidents['Root Cause'].value_counts().reset_index()
    root_cause_counts.columns = ['Root Cause', 'Count']

    root_cause_bar = px.bar(
        data_frame=root_cause_counts,
        x='Count',
        y='Root Cause',
        color='Root Cause',
        color_discrete_sequence=['rgb(230, 0, 0)', 'rgb(84,87,90)', 'rgb(235,151,0)', 'Black', 'rgb(156,42,160)', 'rgb(168,180,0)', 'rgb(254,203,0)', 'rgb(0,124,146)', 'rgb(0,176,202)', 'rgb(94,39,80)'],
        labels={'Count': 'Count', 'Root Cause': 'Root Cause'},
        title='Number of Incidents Logged by Root Cause',
        orientation='h'
    )

    root_cause_bar.update_layout(
        xaxis=dict(showgrid=False),
        yaxis=dict(showgrid=False),
        showlegend=False
    )

    st.plotly_chart(root_cause_bar, use_container_width=True)

    
def display_time_logged_histogram(df_Incidents):
    
    histogram = px.histogram(
        data_frame=df_Incidents,
        x='Time Logged',
        color='Root Cause',
        labels={'Time Logged': 'Time Logged', 'Root Cause': 'Root Cause'},
        title='Time Logged Histogram by Root Cause'
    )

    histogram.update_layout(
        xaxis=dict(title='Time Logged', showgrid=False),
        yaxis=dict(title='Count', showgrid=False),
        showlegend=True
    )

    st.plotly_chart(histogram, use_container_width=True)



def main():
    # Set page configuration
    st.set_page_config(
        page_title="Service Portfolio Overview",
        page_icon=":bar_chart:",
        layout="wide"
    )

    # File Uploader
    uploaded_file = st.sidebar.file_uploader(
        label='Upload your Excel file',
        type=['xlsx']
    )

    #Page Format
    check_uploaded_file(uploaded_file)
    df = read_excel_file(uploaded_file)
    check_required_sheets(df)
    df_Cust_Info, df_Incidents, df_Avail, df_adhoc = extract_dataframes(df)
    diplay_cust_info(df_Cust_Info)
    total_sites, missed_sla_count, above_target_count = calculate_availability_metrics(df_Avail)
    display_kpis(total_sites, missed_sla_count, above_target_count)
    total_incidents, total_customer_caused, total_vodacom_caused, total_closed, open_incidents = calculate_incident_metrics(df_Incidents)
    display_incident_metrics(total_incidents, total_customer_caused, total_vodacom_caused, total_closed, open_incidents)
    display_top_problematic_sites(df_Incidents)
    display_root_cause_count(df_Incidents)
    display_time_logged_histogram(df_Incidents)
    display_site_availability_report(df_Cust_Info, df_Avail)
    display_3_month_availability_trend(df_adhoc)
    display_incident_report(df_Cust_Info, df_Incidents)
    

if __name__ == "__main__":
    main()
