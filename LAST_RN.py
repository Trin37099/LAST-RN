import pandas as pd
import numpy as np
from datetime import datetime
import datetime
from openpyxl import load_workbook
import matplotlib.pyplot as plt
import seaborn as sns
import plotly.express as px
import streamlit as st
import re
import warnings
import calendar
warnings.filterwarnings('ignore')

st.set_page_config(
    page_title="LAST RN",
    layout = 'wide',
)
st.markdown('**LAST 20 RN (AMBER 85)**')

all = pd.read_csv('reservations_summary_report (5).csv',thousands=',')
def convert_room_type(room_type):
  if re.search(r'\bGRAND DELUXE ROOM\b|\bGRAND DELUXE\b|\bGRAND DELUXE DOUBLE ROOM\b|\bGRAND DELUXE ROOM ONLY\b|\bGRAND DOUBLE OR TWIN ROOM\b|\bDOUBLE GRAND DELUXE DOUBLE ROOM\b', room_type):
    return 'GRAND DELUXE'
  elif re.search(r'\bDELUXE DOUBLE ROOM\b|\bDELUXE DOUBLE OR TWIN ROOM WITH CITY VIEW\b|\bDELUXE ROOM CITY VIEW\b|\bDELUXE ROOM ONLY\b|\bDELUXE DOUBLE OR TWIN ROOM\b|\bNEW DELUXE DOUBLE\b|\bDELUXE ROOM\b', room_type):
    return 'NEW DELUXE'
  elif re.search(r'\bNEW DELUXE TWIN\b|\bDELUXE TWIN ROOM\b|\bDOUBLE OR TWIN NEW DELUXE DOUBLE OR TWIN\b|\bDELUXE TWIN ROOM ONLY\b|\bTWIN NEW DELUXE TWIN ROOM\b', room_type):
    return 'NEW DELUXE TWIN'
  elif re.search(r'\bGRAND CORNER SUITES\b|\bGRAND DELUXE\b|\bSUITE WITH BALCONY\b|\bGRAND CORNER SUITES ROOM ONLY\b|\bSUITE SUITE GRAND CORNER\b|\bGRAND STUDIO SUITE\b|\bGRAND CORNER SUITE\b', room_type):
    return 'GRAND CORNER SUITES'
  elif re.search(r'\bMIXED ROOM\b', room_type):
    return 'MIXED'
  else: 
    return 'UNKNOWN'
def apply_discount(channel, adr):
    if channel == 'Booking.com':
      return adr * 0.82
    elif channel == 'Expedia':
      return adr * 0.80
    else:
      return adr

def clean_room_type(room_type):
    if ' X '  in room_type:
        room_type = 'MIXED ROOM'
    return room_type

def calculate_adr_per_rn_abf(row):
    if row['RO/ABF'] == 'ABF':
      return row['ADR'] - 260
    else:
      return row['ADR']
def convert_RF(room_type):
      if re.search(r'\bNON REFUNDABLE\b|\bไม่สามารถคืนเงินจอง\b|\bNON REFUND\b|\bNON-REFUNDABLE\b|\bNRF\b', room_type):
            return 'NRF'
      elif re.search(r'\bUNKNOWN ROOM\b', room_type):
            return 'UNKNOWN'
      elif  room_type == "1 X " or room_type == "2 X " or room_type == "3 X " or room_type == "4 X ":
            return 'UNKNOWN'
      else:
            return 'Flexible'

def convert_ABF(room_type):
      if re.search(r'\bBREAKFAST\b|\bWITH BREAKFAST\b|\bBREAKFAST INCLUDED\b', room_type):
            return 'ABF'
      elif re.search(r'\bUNKNOWN ROOM\b', room_type):
            return 'UNKNOWN'
      elif  room_type == "1 X " or room_type == "2 X " or room_type == "3 X " or room_type == "4 X ":
            return 'UNKNOWN'
      elif re.search(r'\bRO\b|\bROOM ONLY\b', room_type):
            return 'RO'
      else:
            return 'RO'

def perform(all): 
                all1 = all[['Booking reference'
                            ,'Guest names'
                            ,'Check-in'
                            ,'Check-out'
                            ,'Channel'
                            ,'Room'
                            ,'Booked-on date'
                            ,'Total price']]
                all1 = all1.dropna()

                all1["Check-in"] = pd.to_datetime(all1["Check-in"])
                all1['Booked-on date'] = pd.to_datetime(all1['Booked-on date'])
                all1['Booked'] = all1['Booked-on date'].dt.strftime('%m/%d/%Y')
                all1['Booked'] = pd.to_datetime(all1['Booked'])
                all1["Check-out"] = pd.to_datetime(all1["Check-out"])
                all1["Length of stay"] = (all1["Check-out"] - all1["Check-in"]).dt.days
                all1["Lead time"] = (all1["Check-in"] - all1["Booked"]).dt.days
                value_ranges = [-1, 0, 1, 2, 3, 4, 5, 6, 7,8, 14, 30, 90, 120]
                value_ranges1 = [1,2,3, 4,5,6,7,8,9,10,14,30,45,60]
                labels = ['-one', 'zero', 'one', 'two', 'three', 'four', 'five', 'six','seven', '8-14', '14-30', '31-90', '90-120', '120+']
                labels1 = ['one', 'two', 'three', 'four', 'five', 'six','seven','eight', 'nine', 'ten', '14-30', '30-45','45-60', '60+']
                all1['Lead time range'] = pd.cut(all1['Lead time'], bins=value_ranges + [float('inf')], labels=labels, right=False)
                all1['LOS range'] = pd.cut(all1['Length of stay'], bins=value_ranges1 + [float('inf')], labels=labels1, right=False)

                all1['Room'] = all1['Room'].str.upper()
                all1['Booking reference'] = all1['Booking reference'].astype('str')
                all1['Total price'] = all['Total price'].str.strip('THB')
                all1['Total price'] = all1['Total price'].astype('float64')

                all1['Quantity'] = all1['Room'].str.extract('^(\d+)', expand=False).astype(int)
                #all1['Room Type'] = all1['Room'].apply(lambda x: convert_room_type(x))
                #all1['Room Type'] = all1['Room'].str.replace('^DELUXE \(DOUBLE OR TWIN\) ROOM ONLY$', 'DELUXE TWIN')
                all1['Room Type'] = all1['Room'].str.replace('-.*', '', regex=True)
                all1['Room Type'] = all1['Room Type'].apply(lambda x: re.sub(r'^\d+\sX\s', '', x))
                all1['Room Type'] = all1['Room Type'].apply(clean_room_type)
                all1['Room Type'] = all1['Room Type'].apply(lambda x: convert_room_type(x))
                all1['F/NRF'] = all1['Room'].apply(lambda x: convert_RF(x))
                all1['RO/ABF'] = all1['Room'].apply(lambda x: convert_ABF(x))
                #all1['Room Type'] = all1['Room Type'].str.replace('(NRF)', '').apply(lambda x: x.replace('()', ''))
                #all1['Room Type'] = all1['Room Type'].str.replace('WITH BREAKFAST', '')
                #all1['Room Type'] = all1['Room Type'].str.replace('ROOM ONLY', '')
                #all1['Room Type'] = all1['Room Type'].replace('', 'UNKNOWN ROOM')
                #all1['Room Type'] = all1['Room Type'].str.strip()
                all1['ADR'] = (all1['Total price']/all1['Length of stay'])/all1['Quantity']
                all1['ADR'] = all1.apply(lambda row: apply_discount(row['Channel'], row['ADR']), axis=1)
                all1['RN'] = all1['Length of stay']*all1['Quantity']
                all1['ADR'] = all1.apply(calculate_adr_per_rn_abf, axis=1)

                all2 = all1[['Check-in'
                            ,'Check-out'
                            ,'Channel'
                            ,'Booked-on date'
                            ,'Total price'
                            ,'ADR'
                            ,'Length of stay'
                            ,'Lead time'
                            ,'RN'
                            ,'Quantity'
                            ,'Room'
                            ,'Room Type'
                            ,'RO/ABF'
                            ,'F/NRF'
                            ,'Lead time range'
                            ,'LOS range']]
                return all2


all3 =  perform(all)
filtered_df = all3
filtered_df['Stay'] = filtered_df.apply(lambda row: pd.date_range(row['Check-in'], row['Check-out']- pd.Timedelta(days=1)), axis=1)
filtered_df = filtered_df.explode('Stay').reset_index(drop=True)
filtered_df = filtered_df[['Stay','Check-in','Check-out','Booked-on date','Channel','ADR','Length of stay','Lead time','Lead time range','RN','Quantity','Room Type','Room']]


filtered_df =  filtered_df.sort_values(by='Booked-on date')
filtered_df['ADR'] = filtered_df['ADR'].apply('{:.2f}'.format)
filtered_df['ADR'] = filtered_df['ADR'].astype('float')

stay_last20_dict = {}

# Group the DataFrame by 'Stay' column
for stay, group in filtered_df.groupby('Stay'):
    # Select the last 40 rows for each group and reset the index
    last20 = group.tail(40).reset_index(drop=True)
    num_rows = len(last20)
    
    # Generate 'LAST RN' values based on the number of rows
    if num_rows < 40:
        last20['LAST RN'] = list(range(1, num_rows + 1))
    else:
        last20['LAST RN'] = list(range(1, 41))

    # Extract the required columns and convert to a list
    last20_bookings = last20[['Booked-on date', 'ADR', 'Room Type', 'LAST RN']].values.tolist()

    # Store the list of last 20 bookings in the dictionary
    stay_last20_dict[stay] = last20_bookings

# Create an empty DataFrame to store the final results
df_stay_last20 = pd.DataFrame(columns=['Stay', 'Booked-on date', 'ADR', 'Room Type', 'LAST RN'])

# Iterate over the dictionary and append the bookings to the DataFrame
for stay, bookings in stay_last20_dict.items():
    for booking in bookings:
        if len(booking) >= 4:  # Check if booking has at least four elements
            df_stay_last20 = df_stay_last20.append({
                'Stay': stay,
                'Booked-on date': booking[0],
                'ADR': booking[1],
                'Room Type': booking[2],
                'LAST RN': booking[3]
            }, ignore_index=True)
ALL = df_stay_last20
ALL['LAST RN'] = ALL['LAST RN'].astype(int)
ALL['Month'] = pd.to_datetime(ALL['Stay']).dt.month
ALL = ALL.drop(ALL[ALL['Month'] == 5].index)
ALL['Stay'] = ALL['Stay'].astype(str)
mean_by_month_and_rn = ALL.groupby(['Month', 'LAST RN'])['ADR'].mean().reset_index()
mean_by_month_and_rn['ADR'] = mean_by_month_and_rn['ADR'].apply('{:.2f}'.format)
mean_by_month_and_rn['ADR'] = mean_by_month_and_rn['ADR'].astype('float')

ND = df_stay_last20[df_stay_last20['Room Type']== 'NEW DELUXE']
ND['LAST RN'] = ND['LAST RN'].astype(int)
ND['Month'] = pd.to_datetime(ND['Stay']).dt.month
ND = ND.drop(ND[ND['Month'] == 5].index)
ND['Stay'] = ND['Stay'].astype(str)
mean_by_month_and_rn0 = ND.groupby(['Month', 'LAST RN'])['ADR'].mean().reset_index()
mean_by_month_and_rn0['ADR'] = mean_by_month_and_rn0['ADR'].apply('{:.2f}'.format)
mean_by_month_and_rn0['ADR'] = mean_by_month_and_rn0['ADR'].astype('float')


GD = df_stay_last20[df_stay_last20['Room Type']== 'GRAND DELUXE']
GD['LAST RN'] = GD['LAST RN'].astype(int)
GD['Month'] = pd.to_datetime(GD['Stay']).dt.month
GD = GD.drop(GD[GD['Month'] == 5].index)
GD['Stay'] = GD['Stay'].astype(str)
mean_by_month_and_rn1 = GD.groupby(['Month', 'LAST RN'])['ADR'].mean().reset_index()
mean_by_month_and_rn1['ADR'] = mean_by_month_and_rn1['ADR'].apply('{:.2f}'.format)
mean_by_month_and_rn1['ADR'] = mean_by_month_and_rn1['ADR'].astype('float')


NDT = df_stay_last20[df_stay_last20['Room Type']== 'NEW DELUXE TWIN']
NDT['LAST RN'] = NDT['LAST RN'].astype(int)
NDT['Month'] = pd.to_datetime(NDT['Stay']).dt.month
NDT = NDT.drop(NDT[NDT['Month'] == 5].index)
NDT['Stay'] = NDT['Stay'].astype(str)
mean_by_month_and_rn2 = NDT.groupby(['Month', 'LAST RN'])['ADR'].mean().reset_index()
mean_by_month_and_rn2['ADR'] = mean_by_month_and_rn2['ADR'].apply('{:.2f}'.format)
mean_by_month_and_rn2['ADR'] = mean_by_month_and_rn2['ADR'].astype('float')

GC = df_stay_last20[df_stay_last20['Room Type']== 'GRAND CORNER SUITES']
GC['LAST RN'] = GC['LAST RN'].astype(int)
GC['Month'] = pd.to_datetime(GC['Stay']).dt.month
GC = GC.drop(GC[GC['Month'] == 5].index)
GC['Stay'] = GC['Stay'].astype(str)
mean_by_month_and_rn3 = GC.groupby(['Month', 'LAST RN'])['ADR'].mean().reset_index()
mean_by_month_and_rn3['ADR'] = mean_by_month_and_rn3['ADR'].apply('{:.2f}'.format)
mean_by_month_and_rn3['ADR'] = mean_by_month_and_rn3['ADR'].astype('float')

UK = df_stay_last20[df_stay_last20['Room Type']== 'UNKNOWN']
UK['LAST RN'] = UK['LAST RN'].astype(int)
UK['Month'] = pd.to_datetime(UK['Stay']).dt.month
UK = UK.drop(UK[UK['Month'] == 5].index)
UK['Stay'] = UK['Stay'].astype(str)
mean_by_month_and_rn4 = UK.groupby(['Month', 'LAST RN'])['ADR'].mean().reset_index()
mean_by_month_and_rn4['ADR'] = mean_by_month_and_rn4['ADR'].apply('{:.2f}'.format)
mean_by_month_and_rn4['ADR'] = mean_by_month_and_rn4['ADR'].astype('float')

MIXED = df_stay_last20[df_stay_last20['Room Type']== 'MIXED']
MIXED['LAST RN'] = MIXED['LAST RN'].astype(int)
MIXED['Month'] = pd.to_datetime(MIXED['Stay']).dt.month
MIXED = MIXED.drop(MIXED[MIXED['Month'] == 5].index)
MIXED['Stay'] = MIXED['Stay'].astype(str)
mean_by_month_and_rn5 = MIXED.groupby(['Month', 'LAST RN'])['ADR'].mean().reset_index()
mean_by_month_and_rn5['ADR'] = mean_by_month_and_rn5['ADR'].apply('{:.2f}'.format)
mean_by_month_and_rn5['ADR'] = mean_by_month_and_rn5['ADR'].astype('float')

fig = px.line(mean_by_month_and_rn, x='LAST RN', y='ADR',color='Month',text='ADR')
fig.update_traces(textposition='top center')
fig.update_layout(title='Plot of ADR by LAST RN  (ALL ROOM TYPE)')
st.plotly_chart(fig,use_container_width=True)

C1,C2 = st.columns(2)
with C1:
  fig1 = px.line(mean_by_month_and_rn0, x='LAST RN', y='ADR',color='Month',text='ADR')
  fig1.update_traces(textposition='top center')
  fig1.update_layout(title='Plot of ADR by LAST RN  (NEW DELUXE)')
  st.plotly_chart(fig1,use_container_width=True)
with C2:
  fig2 = px.line(mean_by_month_and_rn2, x='LAST RN', y='ADR',color='Month',text='ADR')
  fig2.update_traces(textposition='top center')
  fig2.update_layout(title='Plot of ADR by LAST RN  (NEW DELUXE TWIN)')

C1,C2 = st.columns(2)
with C1:
  fig1 = px.line(mean_by_month_and_rn1, x='LAST RN', y='ADR',color='Month',text='ADR')
  fig1.update_traces(textposition='top center')
  fig1.update_layout(title='Plot of ADR by LAST RN  (GRAND DELUXE)')
  st.plotly_chart(fig1,use_container_width=True)
with C2:
  fig2 = px.line(mean_by_month_and_rn3, x='LAST RN', y='ADR',color='Month',text='ADR')
  fig2.update_traces(textposition='top center')
  fig2.update_layout(title='Plot of ADR by LAST RN  (GRAND CORNER SUITES)')

C1,C2 = st.columns(2)
with C1:
  fig1 = px.line(mean_by_month_and_rn4, x='LAST RN', y='ADR',color='Month',text='ADR')
  fig1.update_traces(textposition='top center')
  fig1.update_layout(title='Plot of ADR by LAST RN  (UNKNOWN)')
  st.plotly_chart(fig1,use_container_width=True)
with C2:
  fig2 = px.line(mean_by_month_and_rn5, x='LAST RN', y='ADR',color='Month',text='ADR')
  fig2.update_traces(textposition='top center')
  fig2.update_layout(title='Plot of ADR by LAST RN  (MIXED)')
