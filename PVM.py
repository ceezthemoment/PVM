import copy
import pandas as pd
import plotly.graph_objects as go


df = pd.read_excel(io="C:/Users/cesar/Downloads/Dummy Business Data.xlsx", sheet_name='Sheet1')
df['Year'] = pd.DatetimeIndex(df['Date']).year

# data cleaning process to ensure data is unique, preventing duplication at merge
df = df.groupby(['Customer', 'Year', 'Product']).sum(['Revenue', 'Volume'])
df = df.reset_index()

# for reconciliation at the end
df_total = df[['Year',  'Revenue', 'Volume']].groupby('Year').sum(['Revenue', 'Volume'])
df_total = df_total.reset_index()

# deepcopy ensures that there are 2 separate dataframes in 2 separate variables
# otherwise year-1 will update both dataframes in each variable
df_cy = copy.deepcopy(df)
df_cy.Year = df_cy.Year - 1

df = df.merge(df_cy, how='outer', on=['Product', 'Customer', 'Year'], suffixes=["_py", "_cy"])
df[['Revenue_py', 'Volume_py', 'Revenue_cy', 'Volume_cy']] = df[['Revenue_py', 'Volume_py', 'Revenue_cy', 'Volume_cy']].fillna(0)

df['Price_py'] = df.Revenue_py / df.Volume_py
df['Price_cy'] = df.Revenue_cy / df.Volume_cy

df['Price effect'] = df.apply(lambda x: (x.Price_cy - x.Price_py) * x.Volume_py if (x['Revenue_py']>0) & (x['Revenue_cy']>0) else 0, axis=1)
df['Volume effect'] = df.apply(lambda x: (x.Volume_cy - x.Volume_py) * x.Price_py if (x['Revenue_py']>0) & (x['Revenue_cy']>0) else 0, axis=1)
df['Mix effect'] = df.apply(lambda x: (x.Price_cy - x.Price_py) * (x.Volume_cy - x.Volume_py) if (x['Revenue_py']>0) & (x['Revenue_cy']>0) else 0, axis=1)
df['New'] = df.apply(lambda x: x['Revenue_cy'] if (x['Revenue_py']==0) & (x['Revenue_cy']>0) else 0, axis=1)
df['Lost'] = df.apply(lambda x: (-1 * x['Revenue_py']) if (x['Revenue_py']>0) & (x['Revenue_cy']==0) else 0, axis=1)

# this section can be updated to allow for quick filtering to look at PVM at a singular product or customer level
rev_2018 = df['Revenue_py'].loc[df['Year']==2018].sum()
pe_2018 = df['Price effect'].loc[df['Year']==2018].sum()
ve_2018 = df['Volume effect'].loc[df['Year']==2018].sum()
me_2018 = df['Mix effect'].loc[df['Year']==2018].sum()
new_2018 = df['New'].loc[df['Year']==2018].sum()
lost_2018 = df['Lost'].loc[df['Year']==2018].sum()
rev_2019 = rev_2018 + pe_2018 + ve_2018 + me_2018 + new_2018 + lost_2018

pe_2019 = df['Price effect'].loc[df['Year']==2019].sum()
ve_2019 = df['Volume effect'].loc[df['Year']==2019].sum()
me_2019 = df['Mix effect'].loc[df['Year']==2019].sum()
new_2019 = df['New'].loc[df['Year']==2019].sum()
lost_2019 = df['Lost'].loc[df['Year']==2019].sum()
rev_2020 = rev_2019 + pe_2019 + ve_2019 + me_2019 + new_2019 + lost_2019

bar_labels = [rev_2018, pe_2018, ve_2018, me_2018, new_2018, lost_2018, rev_2019, pe_2019, ve_2019, me_2019, new_2019, lost_2019, rev_2020]
bar_labels = [label / 1000000 for label in bar_labels]
bar_labels = ["%.1f" % label for label in bar_labels]
bar_labels = ["{}".format(label)+"M" for label in bar_labels]

# reconciliation can be updated to be automated (i.e. display a true in the console)
df_rec = df[['Price effect', 'Volume effect', 'Mix effect', 'New', 'Lost', 'Year']].groupby('Year').sum(['Price effect', 'Volume effect', 'Mix effect', 'New', 'Lost'])

py_2018 = df['Revenue_py'].sum()
print(df_rec)
print(df_total)

start_pos = min(rev_2018,rev_2019,rev_2020)/3
fig = go.Figure(go.Waterfall(
    x = [["2018","2018-19","2018-19","2018-19","2018-19","2018-19","2019","2019-20","2019-20","2019-20","2019-20","2019-20","2020"],
        ["Total", "Price effect", "Volume effect", "Mix effect", "New", "Lost", "Total", "Price effect", "Volume effect", "Mix effect", "New", "lost", "Total"]],
    measure = ["absolute", "relative", "relative", "relative", "reletive", "relative", "total", "relative", "relative", "relative", "relative", "relative", "total"],
    y = [rev_2018 - start_pos, pe_2018, ve_2018, me_2018, new_2018, lost_2018, None, pe_2019, ve_2019, me_2019, new_2019, lost_2019, None],
    base = start_pos,
    text = bar_labels,
    decreasing = {"marker":{"color":"Maroon", "line":{"color":"red", "width":2}}},
    increasing = {"marker":{"color":"Teal"}},
    totals = {"marker":{"color":"deep sky blue", "line":{"color":"blue", "width":3}}}
))

fig.update_layout(title = "Price-volume mix 2018-2020", waterfallgap = 0.3)
fig.show()