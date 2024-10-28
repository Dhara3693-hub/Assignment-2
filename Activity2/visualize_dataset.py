# %% Import excel to dataframe
import pandas as pd
df = pd.read_excel("Online Retail.xlsx")
df

# %%  Show the first 10 rows

df.head(10)

# %% Generate descriptive statistics regardless the datatypes
df.info()
df.describe(include='all')  # include='all' to include all datatypes


# %% Remove all the rows with null value and generate stats again
df_cleaned = df.dropna()
df_cleaned.describe()
df_cleaned.info()

# %% Remove rows with invalid Quantity (Quantity being less than 0)
df_quantity = df[df['Quantity'] >= 0]
df_quantity    

# %% Remove rows with invalid UnitPrice (UnitPrice being less than 0)
df_unitprice = df[df['UnitPrice'] >= 0]
df_unitprice

# %% Only Retain rows with 5-digit StockCode

df_stockcode = df[df['StockCode'].astype(str).str.len() == 5]
df_stockcode.describe(include='all')

# %% strip all description

df['Description']= df['Description'].str.rstrip()
df['Description']

# %% Generate stats again and check the number of rows
df.describe()
df.info()

# %% Plot top 5 selling countries
import matplotlib.pyplot as plt
import seaborn as sns

top5_selling_countries = df["Country"].value_counts()[:5]
sns.barplot(x=top5_selling_countries.index, y=top5_selling_countries.values)
plt.xlabel("Country")
plt.ylabel("Amount")
plt.title("Top 5 Selling Countries")


# %% Plot top 20 selling products, drawing the bars vertically to save room for product description
import matplotlib.pyplot as plt
import seaborn as sns

import pandas as pd 
df = pd.read_excel("online Retail.xlsx")

top20_selling_products = df["Description"].value_counts()[:20]
sns.barplot(x=top20_selling_products.values, y=top20_selling_products.index, hue=top20_selling_products.index, palette='viridis', legend=False)
plt.xlabel("Amount")
plt.ylabel("Products")
plt.title("Top 20 Selling products")
plt.show()

# %% Focus on sales in UK
import matplotlib.pyplot as plt
import seaborn as sns

import pandas as pd
df = pd.read_excel("Online Retail.xlsx")

uk = df[df["Country"] == "United Kingdom"]
df['Total_sales']= df['Quantity'] * df['UnitPrice']
uk_sales= df.groupby('Country').sum('Total_sales')

sns.barplot(data=uk_sales, x='Country', y='Total_sales')
plt.xlabel('Country')
plt.ylabel('Total sales')
plt.title('United Kingdom Sales')
plt.show()


#%% Show gross revenue by year-month
import matplotlib.pyplot as plt
import seaborn as sns

import pandas as pd 
df = pd.read_excel("online Retail.xlsx")

from datetime import datetime

df["YearMonth"] = df["InvoiceDate"].apply(
    lambda dt: datetime(year=dt.year, month=dt.month, day=1)
)

df['Grossrevenue'] = df ['Quantity'] * df ['UnitPrice']
yearmonth_df = df.groupby('YearMonth').sum('Grossrevenue')

sns.lineplot(data=yearmonth_df, x='YearMonth', y='Grossrevenue')
plt.xlabel("Yearmonth")
plt.ylabel("Grossrevenue")
plt.title("UK Sales")
plt.show()

# %% save df in pickle format with name "UK.pkl" for next lab activity
# we are only interested in InvoiceNo, StockCode, Description columns
import pandas as pd
df = pd.read_excel("Online Retail.xlsx")
df 

uk_df = df[df['Country'] == "United Kingdom"]
UK_pkl = df[['InvoiceNo', 'StockCode', 'Description']]
UK_pkl

UK_pkl.to_pickle('UK.pkl')  


# %%
