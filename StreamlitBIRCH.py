import streamlit as st
import pandas as pd
import numpy as np
import openpyxl
import plotly.express as px
import json
import os
import gspread
from google.oauth2.service_account import Credentials

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

df_Budget = get_data_from_excel('Budget')   #pd.read_excel(   'BirchDashboardData.xlsx', sheet_name='Budget'    )
df_ICCeiling = get_data_from_excel('IC Ceilings')
df_Invoices = get_data_from_excel('Invoices')
df_POs = get_data_from_excel('POs')

st.set_page_config(
							page_title="BIRCH Dashboard",
							page_icon=":bar_chart:",
							layout="wide"
				  )
	
#st.dataframe(df_Budget)

#Sidebar
st.sidebar.header("Please Filter here:")
country = st.sidebar.multiselect(
		"Select the country/organization:",
		options=df_Budget['OrganizationOrCountry'].unique(),
		default='Chad' #df_Budget['OrganizationOrCountry'].unique()
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
		"OrganizationOrCountry == @country & Provider == @Provider & FundingSource == @SourceOfFunds"
)

df_ICCeiling_Selection = df_ICCeiling.query(
		"Country == @country" 
)

df_Invoices_Selection = df_Invoices.query(
		"OrganizationOrCountry == @country" 
)

#Mainpage
st.title(":bar_chart: BIRCH Dashboard")
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
