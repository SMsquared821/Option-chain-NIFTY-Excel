#!/usr/bin/env python
# coding: utf-8

# In[1]:


import requests
import pandas as pd
import xlwings as xw
import time


# In[2]:


url = 'https://www.nseindia.com/api/option-chain-indices?symbol=NIFTY'
headers = {
    'user-agent': '', #add your user agent here
    'accept-encoding': '', #add encoding in the single quotes
    'accept-language': ''} #add lamguage in the single quotes


# In[8]:


session = requests.Session()
data= session.get(url, headers=headers).json()["records"]["data"]
ocdata= []
for i in data:
    for j,k in i.items():
        if j=="CE" or j=="PE":
            info=k
            info["instrumentType"]=j
            ocdata.append(info)


# In[10]:


df= pd.DataFrame(ocdata)
wb=xw.Book("optionchaintracker.xlsx") #this is the MS Excel file that should be opened in the background where Option chain will be displayed
st=wb.sheets("nifty")
st.range("A1").value=df
time.sleep(180) #time for refresh of data


# In[ ]:




