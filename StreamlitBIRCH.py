import streamlit as st
import pandas as pd
import numpy as np
import openpyxl
import plotly.express as px
import json
import os
import gspread
from google.oauth2.service_account import Credentials
from gspread_dataframe import get_as_dataframe

def get_data_from_excel(  name  ):

    scope = [   "https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"    ]

    credentials = Credentials.from_service_account_info( st.secrets["gcp_service_account"], scopes=scope  )

    #Authenticate and open the Google sheet
    gc = gspread.authorize( credentials )
    sheet = gc.open_by_url(  "https://docs.google.com/spreadsheets/d/1HyeMeiwmFHgwMTYt7vGYYABpiOB3Oq0WdQwY-rj1ATE"    )

    #Access a specific worksheet
    worksheet_name = str(name)
    worksheet = sheet.worksheet( worksheet_name )

    #Convert to dataframe
    data = worksheet.get_all_records()
    df = pd.DataFrame(data)

    return(df)

def get_metadata_from_excel(  name, url  ):

    scope = [   "https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"    ]

    credentials = Credentials.from_service_account_info( st.secrets["gcp_service_account"], scopes=scope  )

    #Authenticate and open the Google sheet
    gc = gspread.authorize( credentials )
    sheet = gc.open_by_url(  str(url)    )

    #Access a specific worksheet
    worksheet_name = str(name)
    worksheet = sheet.worksheet( worksheet_name )

    # Convert to DataFrame, evaluating formulas to retrieve hyperlinks
    df = get_as_dataframe(worksheet, evaluate_formulas=True)

    return(df)

    # Get the current directory path
    #__location__ = os.path.realpath(  os.path.join( os.getcwd(), os.path.dirname(__file__) ))
    #__location__= os.getcwd()
    #FileName='/BirchDashboardData.xlsx'
    #__filepath__=  __location__ + FileName
    #df = pd.read_excel( 
    #	                        'BirchDashboardData.xlsx',
    #	                        sheet_name=str(name)
    #                  )
    #return(df)
#Retrieve info from each provider
df_Budget_LMH = get_data_from_excel('Tracker', 'https://docs.google.com/spreadsheets/d/1AtleZY2uwjDi4AhG58aggfcK8xeVYvp8DN_UA3OV5eY' )   
df_Budget_LMH.insert(  1, 'Provider', 'LMH'  )
df_Budget_CHAI = get_data_from_excel('Tracker', 'https://docs.google.com/spreadsheets/d/1v5Zo3EQ8TlmR9fluEWAa9cW-1OnwZQpN2U-2JKKzaN8' )
df_Budget_CHAI.insert(  1, 'Provider', 'CHAI'  )
df_Budget_ICHESS = get_data_from_excel('Tracker', 'https://docs.google.com/spreadsheets/d/1obbDO0Z-W8etRzMocabYFTCgR0dzoSbBzv7y07PkBlM' )
df_Budget_ICHESS.insert(  1, 'Provider', 'ICHESS'  )
df_Budget_JHPIEGO = get_data_from_excel('Tracker', 'https://docs.google.com/spreadsheets/d/1Q0An0_RGG2IvzUQEB5HAZOA87DCkvnwk8bDZ6hvkEFU' )
df_Budget_JHPIEGO.insert(  1, 'Provider', 'JHPIEGO'  )
df_Budget = pd.concat([  df_Budget_LMH, df_Budget_CHAI, df_Budget_ICHESS, df_Budget_JHPIEGO   ])
#Retrieve invoices and everything else from the main sheet
df_ICCeiling = get_data_from_excel('IC Ceilings', "https://docs.google.com/spreadsheets/d/1HyeMeiwmFHgwMTYt7vGYYABpiOB3Oq0WdQwY-rj1ATE" )
df_Deliverables = get_data_from_excel('Deliverables', "https://docs.google.com/spreadsheets/d/1HyeMeiwmFHgwMTYt7vGYYABpiOB3Oq0WdQwY-rj1ATE")
df_Invoices = get_metadata_from_excel('Invoices', "https://docs.google.com/spreadsheets/d/1HyeMeiwmFHgwMTYt7vGYYABpiOB3Oq0WdQwY-rj1ATE")
df_POs = get_metadata_from_excel('POs', "https://docs.google.com/spreadsheets/d/1HyeMeiwmFHgwMTYt7vGYYABpiOB3Oq0WdQwY-rj1ATE")

# Function to convert URLs to clickable links
def convert_to_link(url):
    # Check if the cell is empty or not a string
    if url is None or not isinstance(url, str) or url.strip() == "":
        return ""  # Return an empty string if the cell is empty or just plain text
    
    # Return the HTML anchor tag for the URL
    return f'<a href="{url}" target="_blank">{url}</a>'


# Apply this to cells in the column(s) with hyperlinks
#df_Invoices["Associated PO Link"] = df_Invoices["Associated PO Link"].apply(convert_to_link)
#df_Invoices["Invoice Link"] = df_Invoices["Invoice Link"].apply(convert_to_link)
#df_Invoices["Invoice Draft Link"] = df_Invoices["Invoice Draft Link"].apply(convert_to_link)
#df_POs["Workplan and Budget Link"] = df_POs["Workplan and Budget Link"].apply(convert_to_link)
#df_POs["Workplan and Budget .xlsx Link"] = df_POs["Workplan and Budget .xlsx Link"].apply(convert_to_link)
#df_POs["SEAH Link"] = df_POs["SEAH Link"].apply(convert_to_link)
#df_POs["Due Diligence/Partner Review Link"] = df_POs["Due Diligence/Partner Review Link"].apply(convert_to_link)
#f_POs["POs drafted Link"] = df_POs["POs drafted Link"].apply(convert_to_link)
#f_POs["PO Signed Link"] = df_POs["PO Signed Link"].apply(convert_to_link)

#Set page header
st.set_page_config(
							page_title="BIRCH Project Overview",
							page_icon=":bar_chart:",
							layout="wide"
				  )
	
#st.dataframe(df_Budget)

#Sidebar
st.sidebar.header("Please Filter here:")
country = st.sidebar.multiselect(
		"Select the country/organization:",
		options=df_Budget['Country'].unique(),
		default=df_Budget['Country'].unique()
)

SourceOfFunds = st.sidebar.multiselect(
		"Select the source of funding:",
		options=df_Budget['FundingSource'].unique(),
		default=df_Budget['FundingSource'].unique()
)

Provider = st.sidebar.multiselect(
		"Select the provider:",
		options=df_Budget['Provider'].unique(),
		default=df_Budget['Provider'].unique()
)

df_selection = df_Budget.query(
		"Country == @country & Provider == @Provider & FundingSource == @SourceOfFunds"
)

df_ICCeiling_Selection = df_ICCeiling.query(
		"Country == @country" 
)

df_Invoices_Selection = df_Invoices.query(
		"OrganizationOrCountry == @country" 
)

df_POs_Selection = df_POs.query(
		"Country == @country" 
)

df_Deliverables_Selection = df_Deliverables.query(
		"Country == @country" 
)

#Mainpage
st.title(":bar_chart: BIRCH Project Overview")
st.markdown("##")

# hide streamlit style
hide_st_style = """
				<style>
				#MainMenu {visibility: hidden;}
				footer {visibility: hidden;}
				header {visibility: hidden;}
				</style>
				"""
st.markdown(hide_st_style, unsafe_allow_html=True)

#Top KPIs
TotalApprovedCeiling  = int(df_ICCeiling_Selection['IC Approved Ceiling'].sum())
TotalBudget = int(df_selection['Budget'].sum())
TotalSpent = int(df_Invoices_Selection['Pre-payment Amount'].sum())
Absorption=int(100*( df_Invoices_Selection['Pre-payment Amount'].sum()/df_selection['Budget'].sum() ))

left_column, left_middle_column, right_middle_column, right_column = st.columns(4)
with left_column:
	st.subheader("Total IC Approved Ceiling:")
	st.subheader(f"US${TotalApprovedCeiling:,}")
with left_middle_column:
	st.subheader("Total Budget:")
	st.subheader(f"US${TotalBudget:,}")
with right_middle_column:
	st.subheader("Total spent:")
	st.subheader(f"US${TotalSpent:,}")
with right_column:
	st.subheader("Absorption:")
	st.subheader(f"{Absorption}%")

st.markdown("---")

#Awards by intervention
Awards_by_Intervention = (
							df_selection.groupby( by=["Foundational Element"] ).sum()[['Budget']].sort_values(by="Budget")
						 )

fig_Awards_by_Intervention = px.bar(
										Awards_by_Intervention,
										x="Budget",
										y=Awards_by_Intervention.index,
										orientation="h",
										text_auto=True,
										title="<b>Budget by Foundational Element</b>",
										color_discrete_sequence=["#0083B8"] * len(Awards_by_Intervention),
										template="plotly_white",
										color_continuous_scale='viridis',
										text='Budget',
										hover_data={'Budget':':.1f'}
									)

fig_Awards_by_Intervention.update_layout(
												width=2000,
												height=700,
												plot_bgcolor="rgba(0,0,0,0)",
												xaxis=(dict(showgrid=False)),
											        title={
													"text": "<b>Budget by Foundational Element</b>",  # Title text
													"y": 0.95,  # Position title slightly below the top of the plot
													"x": 0.5,   # Center the title horizontally
													"xanchor": "center",  # Ensure horizontal center alignment
													"yanchor": "top",     # Align the title text to the top
													"font": {"size": 24}, # Set font size
											        }
                                         )

fig_Awards_by_Intervention.update_traces(
                                           texttemplate='%{text:,.0f}', textfont_size=14
                                        )
                                        
#Color-code based dataframe
# Define a function to apply row-wise styling based on the Status column
def highlight_row(row):
    if row["Status"] == "Delayed":
        return ["background-color: red"] * len(row)
    elif row["Status"] == "On track":
        return ["background-color: lightgreen"] * len(row)
    elif row["Status"] == "Complete":
        return ["background-color: green"] * len(row)
    else:
        return [""] * len(row)  # No styling for other statuses

# Apply the styling to the DataFrame
st.header("Budget breakdown")
df_selection_color = df_selection.style.apply(highlight_row, axis=1)
st.dataframe(df_selection_color)
#Add legend
# Create a legend below the DataFrame
legend_html = """
<div style="display: flex; flex-direction: column;">
    <div style="background-color: red; padding: 5px; color: white;">Delayed</div>
    <div style="background-color: lightgreen; padding: 5px;">On track</div>
    <div style="background-color: green; padding: 5px;">Complete</div>
</div>
"""
#Add plot
st.markdown(legend_html, unsafe_allow_html=True)

#Display the other dataframes
st.header("Purchase Orders")
# Display the DataFrame with hyperlinks
#st.markdown(df_POs_Selection.to_html(escape=False, index=False), unsafe_allow_html=True)
st.dataframe(df_POs_Selection)

#Display the other dataframes
st.header("Invoices")
# Display the DataFrame with hyperlinks
#st.markdown(df_Invoices_Selection.to_html(escape=False, index=False), unsafe_allow_html=True)
st.dataframe(df_Invoices_Selection)

#Display the other dataframes
st.header("Deliverables")
# Display the DataFrame with hyperlinks
#st.markdown(df_Invoices_Selection.to_html(escape=False, index=False), unsafe_allow_html=True)
st.dataframe(df_Deliverables_Selection)


#Add Bar chart
st.plotly_chart(fig_Awards_by_Intervention)


#Pie charts for Board Category and ACT-A Pillars
Awards_by_BoardCategory = (
							df_selection.groupby( by=["Provider"] ).sum()[['Budget']].sort_values(by="Budget")
			  )

Awards_by_ACTAPillar = (
							df_selection.groupby( by=["FundingSource"] ).sum()[['Budget']].sort_values(by="Budget")
		       )

fig_Awards_by_BoardCategory = px.pie(
						Awards_by_BoardCategory,
						values='Budget',
						names=Awards_by_BoardCategory.index,
						title="<b>Budget by Provider</b>"
				    )
									
fig_Awards_by_BoardCategory.update_layout(
					        title={
							"text": "<b>Budget by Provider</b>",  # Title text
							"y": 0.95,  # Position title slightly below the top of the plot
							"x": 0.5,   # Center the title horizontally
							"xanchor": "center",  # Ensure horizontal center alignment
							"yanchor": "top",     # Align the title text to the top
							"font": {"size": 16}, # Set font size
					        }
                                         )

									
fig_Awards_by_ACTAPillar = px.pie(
											Awards_by_ACTAPillar,
											values='Budget',
											names=Awards_by_ACTAPillar.index,
											title="<b>Budget by Founding Source</b>"
				 )

fig_Awards_by_ACTAPillar.update_layout(
					        title={
							"text": "<b>Budget by Founding Source</b>",  # Title text
							"y": 0.95,  # Position title slightly below the top of the plot
							"x": 0.5,   # Center the title horizontally
							"xanchor": "center",  # Ensure horizontal center alignment
							"yanchor": "top",     # Align the title text to the top
							"font": {"size": 16}, # Set font size
					        }
                                         )

#Place the pie charts
left_column, right_column = st.columns(2)
left_column.plotly_chart(fig_Awards_by_BoardCategory)
right_column.plotly_chart(fig_Awards_by_ACTAPillar)
