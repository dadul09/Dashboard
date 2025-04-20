import pandas as pd
from dash import Dash, dash_table, dcc, html, Input, Output
import plotly.graph_objs as go
import dash_bootstrap_components as dbc
import socket

# Helper to load and format Final Dashboard data
def load_data():
    with open("NPCB Dash Board 25-26.xlsm", "rb") as f:
        df = pd.read_excel(f, sheet_name="Final Dashboard", skiprows=4, engine='openpyxl')
    df = df[df["Region"].str.upper() != "TOTAL"]
    df = df.dropna(subset=["Region"])
    df.reset_index(drop=True, inplace=True)
    df.iloc[:, 1:] = df.iloc[:, 1:].apply(pd.to_numeric, errors='coerce')
    return df

# Helper to generate ranking table
def create_ranking_table(df, column, title):
    ranking_df = df[["Region", column]].sort_values(by=column, ascending=False).reset_index(drop=True)
    ranking_df.insert(0, "Rank", ranking_df.index + 1)
    ranking_df[column] = ranking_df[column].map(lambda x: f"{x:.1f}" if pd.notnull(x) else "")
    return html.Div([
        html.H5(f"{title} Ranking"),
        dash_table.DataTable(
            data=ranking_df.to_dict("records"),
            columns=[{"name": i, "id": i} for i in ranking_df.columns],
            style_table={'overflowX': 'auto', 'width': '100%'},
            style_cell={'textAlign': 'left', 'padding': '5px'},
            style_header={'backgroundColor': 'lightgrey', 'fontWeight': 'bold'},
        )
    ])

# Build Dash app
app = Dash(__name__, external_stylesheets=[dbc.themes.BOOTSTRAP])

@app.callback(
    Output('page-content', 'children'),
    Input('interval-component', 'n_intervals')
)
def update_dashboard(_):
    df = load_data()
    display_df = df.copy()
    display_df.iloc[:, 1:] = display_df.iloc[:, 1:].applymap(lambda x: f"{x:.1f}" if pd.notnull(x) else "")

    charts = []
    for (title, col1, col2, color1, color2) in [
        ("AOP vs Actual Appropriation (₹ Cr.)", "AOP Appropriation", "Actual Appropriation", "steelblue", "seagreen"),
        ("AOP vs Actual Capex (₹ Cr.)", "AOP Capex", "Actual Capex", "orange", "darkgreen"),
        ("AOP vs Actual Capitalization (₹ Cr.)", "AOP Capitalization", "Actual Capitalization", "indianred", "mediumseagreen")
    ]:
        fig = go.Figure()
        fig.add_trace(go.Bar(x=df["Region"], y=df[col1], name=col1, marker_color=color1,
                             text=[f"{x:.1f} Cr." for x in df[col1]], textposition='outside'))
        fig.add_trace(go.Bar(x=df["Region"], y=df[col2], name=col2, marker_color=color2,
                             text=[f"{x:.1f} Cr." for x in df[col2]], textposition='outside'))
        fig.update_layout(barmode='group', title=title)
        charts.append(dbc.Row([
            dbc.Col(dcc.Graph(figure=fig), md=8),
            dbc.Col(create_ranking_table(df, col2, col2), md=4)
        ]))

    # CWIP and Technically Completed side-by-side
    fig4 = go.Figure([go.Bar(x=df["Region"], y=df["CWIP"], name="CWIP", marker_color="mediumpurple",
                             text=[f"{x:.1f} Cr." for x in df["CWIP"]], textposition='outside')])
    fig4.update_layout(title="CWIP (Capital Work in Progress) (₹ Cr.)")

    fig6 = go.Figure([go.Bar(x=df["Region"], y=df['No. of Projects "Technically Completed"'], name="Technically Completed",
                             marker_color="mediumaquamarine",
                             text=[f"{int(x)}" for x in df['No. of Projects "Technically Completed"']], textposition='outside')])
    fig6.update_layout(title="Technically Completed Projects")

    charts.append(dbc.Row([
        dbc.Col(dcc.Graph(figure=fig4), md=6),
        dbc.Col(dcc.Graph(figure=fig6), md=6)
    ]))

    # Delayed Projects and Negative CWIP side-by-side
    fig7 = go.Figure([go.Bar(x=df["Region"], y=df["No. of Delayed Projects (Time Overrrun)"], name="Delayed Projects",
                             marker_color="tomato",
                             text=[f"{int(x)}" for x in df["No. of Delayed Projects (Time Overrrun)"]], textposition='outside')])
    fig7.update_layout(title="Delayed Projects (Time Overrun)")

    fig8 = go.Figure([go.Bar(x=df["Region"], y=df["-VE CWIP"], name="Negative CWIP", marker_color="slateblue",
                             text=[f"{int(x)}" for x in df["-VE CWIP"]], textposition='outside')])
    fig8.update_layout(title="Projects with Negative CWIP")

    charts.append(dbc.Row([
        dbc.Col(dcc.Graph(figure=fig7), md=6),
        dbc.Col(dcc.Graph(figure=fig8), md=6)
    ]))

    return dbc.Container([
        html.H1("Project Performance Dashboard"),
        html.H4("Performance Metrics"),
        dash_table.DataTable(
            data=display_df.to_dict('records'),
            columns=[{"name": i, "id": i} for i in display_df.columns],
            style_table={'overflowX': 'auto'},
            style_cell={'textAlign': 'left', 'padding': '5px'},
            style_header={'backgroundColor': 'lightgrey', 'fontWeight': 'bold'}
        ),
        html.Br(),
        *charts
    ], fluid=True)

app.layout = html.Div([
    dcc.Interval(id='interval-component', interval=60000, n_intervals=0),
    html.Div(id='page-content')
])

if __name__ == '__main__':
    ip = socket.gethostbyname(socket.gethostname())
    print(f"Dashboard running on http://{ip}:8050")
    app.run(debug=True, host='0.0.0.0', port=8050)

 
