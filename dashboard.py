# =============================
# dashboard.py
# =============================
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
import dash
from dash import dcc, html

# =============================
#  Load Data
# =============================
clients = pd.read_excel("data/clients.xlsx")
finance = pd.read_excel("data/finance.xlsx")

# =============================
#  Clean Client Data
# =============================
clients.columns = clients.columns.str.strip()
clients.rename(columns={
    "Company Name": "Company",
    "Department": "Department",
    "Renewal Number": "Renewal_No",
    "Renewal Date": "Renewal_Date"
}, inplace=True)

clients["Renewal_Date"] = pd.to_datetime(clients["Renewal_Date"].astype(str) + "-01-01", errors="coerce")
clients["Year"] = clients["Renewal_Date"].dt.year
clients = clients.dropna(subset=["Company", "Department", "Year"])

# =============================
#  Clean Finance Data
# =============================
finance.columns = finance.columns.str.strip()
finance["Year"] = finance["Year"].astype(int)

# =============================
#  1. Active Clients by Year
# =============================
clients_per_year = clients.groupby("Year")["Company"].nunique().reset_index(name="Active_Clients")
clients_per_year["Growth_%"] = (clients_per_year["Active_Clients"].pct_change() * 100).round(1)

fig1 = px.bar(clients_per_year, x="Year", y="Active_Clients", text="Active_Clients",
              title="Active Clients by Year", template="plotly_white", color_discrete_sequence=["#0077b6"])
fig1.update_traces(textposition="outside")

# =============================
#  2. Renewal Frequency
# =============================
renewal_counts = clients.groupby("Company")["Renewal_No"].max().value_counts().sort_index().reset_index()
renewal_counts.columns = ["Renewal_Times", "Number_of_Clients"]

fig2 = px.bar(renewal_counts, x="Renewal_Times", y="Number_of_Clients",
              title="Distribution of Client Renewals", text="Number_of_Clients",
              template="plotly_white", color_discrete_sequence=["#00b4d8"])
fig2.update_traces(textposition="outside")

# =============================
#  3. Market Share by Department
# =============================
dept_share = clients.groupby("Department")["Company"].count().reset_index(name="Total_Contracts")
dept_share["Share_%"] = round(100 * dept_share["Total_Contracts"] / dept_share["Total_Contracts"].sum(), 1)

fig3 = px.pie(dept_share, names="Department", values="Total_Contracts",
              title="Market Concentration by Contract Count", hole=0.45, template="plotly_white")
fig3.update_traces(textinfo="label+percent", pull=[0.05]*len(dept_share))

# =============================
#  4. First Contracts Only
# =============================
first_contracts = clients[clients["Renewal_No"] == 0]
first_share = first_contracts.groupby("Department")["Company"].count().reset_index(name="First_Contracts")
first_share["Share_%"] = round(100 * first_share["First_Contracts"] / first_share["First_Contracts"].sum(), 1)

fig4 = px.pie(first_share, names="Department", values="First_Contracts",
              title="Market Concentration by First Contracts Only", hole=0.45, template="plotly_white")
fig4.update_traces(textinfo="label+percent", pull=[0.05]*len(first_share))

# =============================
#  5. Client Retention
# =============================
years = sorted(clients["Year"].unique())
churn_data = []
for i, year in enumerate(years[:-1]):
    current = set(clients[clients["Year"] == year]["Company"])
    next_year = set(clients[clients["Year"] == years[i+1]]["Company"])
    if len(current) > 0:
        churn_rate = round(100 * len(current - next_year) / len(current), 1)
        retention_rate = 100 - churn_rate
    else:
        churn_rate = retention_rate = np.nan
    churn_data.append({"Year": year, "Churn_%": churn_rate, "Retention_%": retention_rate})

churn_df = pd.DataFrame(churn_data)

fig5 = go.Figure()
fig5.add_trace(go.Scatter(x=churn_df["Year"], y=churn_df["Churn_%"], mode="lines+markers", name="Churn Rate (%)",
                          line=dict(color="#ef476f", width=3)))
fig5.add_trace(go.Scatter(x=churn_df["Year"], y=churn_df["Retention_%"], mode="lines+markers", name="Retention Rate (%)",
                          line=dict(color="#06d6a0", width=3)))
fig5.update_layout(title="Yearly Client Churn and Retention Rates",
                   xaxis_title="Year", yaxis_title="Rate (%)", template="plotly_white")

# =============================
#  6. Financial Performance
# =============================
fig6 = px.bar(finance, x="Year", y="Total Sales", color="Department", title="Total Sales by Department and Year",
              template="plotly_white", barmode="group")
fig7 = px.bar(finance, x="Year", y="Total Profit", color="Department", title="Total Profit by Department and Year",
              template="plotly_white", barmode="group")
finance["Profit Margin_%"] = round((finance["Total Profit"] / finance["Total Sales"]) * 100, 1)
fig8 = px.line(finance, x="Year", y="Profit Margin_%", color="Department",
               title="Profit Margin by Department", markers=True, template="plotly_white")

# =============================
#  10. Key Clients Contribution
# =============================
client_contracts = clients.groupby(["Department", "Company"])["Renewal_No"].max().reset_index()
client_contracts["Total_Contracts"] = client_contracts["Renewal_No"] + 1
merged_clients_profit = pd.merge(client_contracts, finance[["Department", "Year", "Total Profit"]], on="Department", how="left")
dept_totals = client_contracts.groupby("Department")["Total_Contracts"].sum().reset_index(name="Dept_Total_Contracts")
merged_clients_profit = pd.merge(merged_clients_profit, dept_totals, on="Department", how="left")
merged_clients_profit["Contracts_Share_%"] = round(100 * merged_clients_profit["Total_Contracts"] / merged_clients_profit["Dept_Total_Contracts"], 1)
merged_clients_profit["Estimated_Profit"] = round(merged_clients_profit["Contracts_Share_%"] / 100 * merged_clients_profit["Total Profit"], 2)
top_clients_per_dept = merged_clients_profit.sort_values(["Department", "Estimated_Profit"], ascending=[True, False])
top_clients_15 = top_clients_per_dept.groupby("Department").head(15).reset_index(drop=True)

fig_top15 = px.bar(top_clients_15.sort_values(["Department", "Estimated_Profit"], ascending=[True, False]),
                   x="Department", y="Estimated_Profit", color="Company", text="Estimated_Profit",
                   title="Top 15 Clients Contribution per Department", labels={"Estimated_Profit": "Estimated Profit (EGP)", "Department": "Department"},
                   template="plotly_white", height=600)
fig_top15.update_traces(textposition="inside")
fig_top15.update_layout(yaxis_title="Estimated Profit (EGP)", xaxis_title="Department", legend_title_text="Client",
                        margin=dict(t=80, b=150))

# =============================
#  Dash App Layout
# =============================
app = dash.Dash(__name__)

app.layout = html.Div([
    html.H1("Proserv Strategic & Financial Dashboard"),
    html.H2("Active Clients by Year"), dcc.Graph(figure=fig1),
    html.H2("Distribution of Client Renewals"), dcc.Graph(figure=fig2),
    html.H2("Market Concentration by Contract Count"), dcc.Graph(figure=fig3),
    html.H2("Market Concentration by First Contracts Only"), dcc.Graph(figure=fig4),
    html.H2("Yearly Client Churn and Retention Rates"), dcc.Graph(figure=fig5),
    html.H2("Total Sales by Department"), dcc.Graph(figure=fig6),
    html.H2("Total Profit by Department"), dcc.Graph(figure=fig7),
    html.H2("Profit Margin by Department"), dcc.Graph(figure=fig8),
    html.H2("Top 15 Clients Contribution per Department"), dcc.Graph(figure=fig_top15),
])

if __name__ == "__main__":
    app.run_server(debug=True)
