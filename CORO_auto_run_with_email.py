#!/usr/bin/env python
# coding: utf-8

# In[49]:


import os, os.path
import win32com.client
import datetime


# In[94]:


os.chdir('C:\\Users\\s456781\\OneDrive - Emirates Group\\Documents\\My Data Sources\\')


# In[96]:


if os.path.exists(".\Skywards CORO Pack.xlsm"):
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible= True
    wb  = excel.Workbooks.Open(os.path.abspath('Skywards CORO Pack.xlsm'))
    wb.RefreshAll()
    excel.CalculateUntilAsyncQueriesDone()
    wb.Save()
    excel.Application.Run("'Skywards CORO Pack.xlsm'!Module2.exltoppt_v2") 
    wb.Close(True)
    excel.Application.Quit()
del excel
del wb


# In[91]:


os.chdir('C:\\Users\\s456781\\OneDrive - Emirates Group\\Documents\\Coro Testing')


# In[92]:


# send the email out


outlook = win32com.client.Dispatch('Outlook.Application')


ol_msg = outlook.CreateItem(0)

ol_msg.Attachments.Add(Source = os.path.abspath("{} Skywards Forward Ticketed Report.pdf".format(datetime.datetime.today().strftime("%Y%m%d"))))
ol_msg.To = 'ruchir.varma@emirates.com'
ol_msg.CC = 'ruchir.varma@emirates.com'
ol_msg.Subject = 'This is the CORO Pack for 2021'
ol_msg.Body = 'Enclosed is the latest CORO pack for Skywards Commercial'
ol_msg.Display()






# ### Appendix

# In[19]:


# #if os.path.exists("Skywards CORO Pack.xlsm"):
# xl=win32com.client.DispatchEx("Excel.Application")
# xl.Workbooks.Open(os.path.abspath("testing excel.xlsx"), Editable = 1)
# xl.visible = True
#     #xl.Application.Run("'Skywards CORO Pack.xlsm'!Module2.exltoppt_v2")
# xl.Application.Save() # if you want to save then uncomment this line and change delete the ", ReadOnly=1" part from the open function.
#     #workbook1.Save() 
#         # Comment this out if your excel script closes
# xl.Application.Quit()
# del xl


# In[ ]:





# In[47]:


# #excel = win32com.client.Dispatch("Excel.Application")
# ppt = win32com.client.Dispatch("PowerPoint.Application")
# excel.Visible= True
# excel.Workbooks.Open(os.path.abspath('Skywards CORO Pack.xlsm'))
# excel.ActiveWorkbook.Saved = True
# excel.Application.Run("'Skywards CORO Pack.xlsm'!Module2.exltoppt_v2")  
# excel.ActiveWorkbook.Close()
# excel.Application.Quit()
# ppt.Quit()


# In[ ]:




