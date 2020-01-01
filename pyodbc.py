
# In[34]:


import pyodbc
import pandas as pd
conn = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};.....................')


# In[36]:


df = pd.read_excel(r'C:\............\Automation\SQL\pbperforming.xlsx') 
inv_array = df['invoiceID'].tolist()



# In[38]:


d = {}
sql=[]
start = """  '20191101 00:00:00' """
end = """  '20191130 23:59:59' """
for i in range(len(inv_array)):
    sql.append(""" 
    SELECT InvoiceNumberFull, cast(p.TimeStamp as date) as ImportDate, Code as Currency, Amount, p.Id from Payments p
				LEFT JOIN invoices i on i.id = p.InvoiceId
				LEFT JOIN CustomerOrders co on co.Id = i.CustomerOrderId
				LEFT JOIN Currencies cu on cu.id = co.CurrencyId
				WHERE PaymentType = 0
					and p.TimeStamp >= """ +  str(start) + 
               """ 
					and p.TimeStamp <= """ +  str(end) +
               """
    				AND i.Sold = 1
					and i.Id =  """ + str(inv_array[i])  )            
    d[i] = pd.read_sql(sql[i],conn)
    print(d[i])
    print(type(d[i]))


# In[39]:


result = pd.concat(d)


# In[40]:


result.to_csv(r'C:\.........\Automation\SQL\result.csv') 


# In[41]:


conn.close()


# In[42]:


print(result)






