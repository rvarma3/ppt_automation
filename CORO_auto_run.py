#!/usr/bin/env python
# coding: utf-8

# In[1]:


import os, os.path
import win32com.client


# In[2]:


os.chdir('C:\\Users\\s456781\\OneDrive - Emirates Group\\Documents\\My Data Sources')


# In[3]:


if os.path.exists("Skywards CORO Pack.xlsm"):
    xl=win32com.client.DispatchEx("Excel.Application")
    workbook1 = xl.Workbooks.Open(os.path.abspath("Skywards CORO Pack.xlsm"), ReadOnly = 1)
    xl.Application.Run("'Skywards CORO Pack.xlsm'!Module2.exltoppt_v2")
    #xl.Application.Save() # if you want to save then uncomment this line and change delete the ", ReadOnly=1" part from the open function.
    workbook1.Save() 
        # Comment this out if your excel script closes
    xl.Application.Quit()
    #del xl


# In[ ]:





# In[ ]:





# In[ ]:





# In[ ]:





# In[ ]:





# In[ ]:




