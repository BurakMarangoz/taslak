import dash
from dash import dcc, html, dash_table
from dash.dependencies import Input, Output
import pandas as pd
import plotly.graph_objs as go
import dash_bootstrap_components as dbc
import plotly.express as px
import datetime
import openpyxl

xls = pd.ExcelFile('kurallar.xlsx')
sheet_to_df_map = {sheet_name: xls.parse(sheet_name) for sheet_name in xls.sheet_names}

company_info_data = {'Künye Bilgileri': ['Firma Adı', 'Müşteri Numarası', 'Şube Adı', 'Şube Kodu',
                                         'Teklif Numarası', 'Tarih', 'Müşteri Sınıfı', 'PD', 'KDS Skoru'],
                     'Değerler': ['ABC Ltd. Şti.', '12345678', 'Levent', '8833', '823483248', '15.11.2023', '2', '1', '3']}

company_info_df = pd.DataFrame(company_info_data)

company_info_table = dash_table.DataTable(
    id='company-info-table',
    columns=[{'name': i, 'id': i} for i in company_info_df.columns],
    data=company_info_df.to_dict('records'),
    style_as_list_view=True,
    style_cell={'textAlign': 'left', 'padding': '5px', 'fontSize': '16px', 'fontFamily': 'Arial'},
    style_header={'backgroundColor': '#000066', 'fontWeight': 'bold', 'color': 'white'},
    style_data_conditional=[{'if': {'row_index': 'odd'},'backgroundColor': 'rgb(248, 248, 248)'}])


data = {'0':   [1, 1, 2, 2, 3],
        '100': [1, 1, 2, 2, 3],
        '200': [1, 1, 2, 2, 3],
        '300': [2, 2, 2, 3, 3],
        '400': [3, 3, 3, 3, 3]}
df_matrix = pd.DataFrame(data, index=[0, 100, 200, 300, 400])

color_scale = [(0, '#00FF00'),  # Green for 1
               (0.5, '#FFFF00'),  # Yellow for 2
               (1, '#FF0000')]  # Red for 3

fig = px.imshow(df_matrix,
                labels=dict(x="TKN", y="TBE", color="Value"),
                x=df_matrix.columns,
                y=df_matrix.index,
                color_continuous_scale=color_scale)

fig.update_layout(width=320, height=320,margin=dict(t=0, l=0),coloraxis_showscale=False)
fig.update_xaxes(side='top')
fig.update_traces(xgap=2, ygap=2)
fig.update_coloraxes(colorscale=color_scale, cmin=1, cmax=3)
fig.add_annotation(x=1,y='300', text="X", showarrow=False, font=dict(color="black", size=20))

data2 = {'0-1':[2,3,3,3, 3, 3, 3, 3],
        '1-2': [2,2,3,3, 3, 3, 3, 3],
        '2-3': [1,2,2,2, 2, 3, 3, 3],
        '3-4': [1,1,2,2, 2, 2, 3, 3],
        '>4':  [1,1,1,2, 2, 2, 3, 3]}
df_matrix2 = pd.DataFrame(data2, index=["0-1", "1-1.50", "1.50-2", "2-2.50", "2.50-3","3-3.50","3.50-4",">4"])

fig2 = px.imshow(df_matrix2,
                labels=dict(x="Kuruluş Tarihi", y="PD", color="Value"),
                x=df_matrix2.columns,
                y=df_matrix2.index,
                color_continuous_scale=color_scale)

fig2.update_layout(width=320, height=320,margin=dict(t=0, l=0),coloraxis_showscale=False)
fig2.update_xaxes(side='top')
fig2.update_traces(xgap=2, ygap=2)
fig2.update_coloraxes(colorscale=color_scale, cmin=1, cmax=3)
fig2.add_annotation(x=2,y='1.50-2', text="X", showarrow=False, font=dict(color="black", size=20))

data3 = {'0-5':   [1,1,1,2, 2, 3],
         '5-10':  [1,1,2,3, 3, 3],
         '10-15': [1,2,2,3, 3, 3],
         '15-25': [2,2,3,3, 3, 3],
         '>25':   [3,3,3,3, 3, 3]}
df_matrix3 = pd.DataFrame(data3, index=["0", "1-5", "5-10", "10-15", "15-25",">25"])

fig3 = px.imshow(df_matrix3,
                labels=dict(x="Talep/GL Oranı", y="KDS", color="Value"),
                x=df_matrix3.columns,
                y=df_matrix3.index,
                color_continuous_scale=color_scale)

fig3.update_layout(width=320, height=320,margin=dict(t=0, l=0),coloraxis_showscale=False)
fig3.update_xaxes(side='top')
fig3.update_traces(xgap=2, ygap=2)
fig3.update_coloraxes(colorscale=color_scale, cmin=1, cmax=3)
fig3.add_annotation(x=2,y='1-5', text="X", showarrow=False, font=dict(color="black", size=20))

data_pd = {"Tarih": ["Ocak 23", "Şubat 23", "Mart 23", "Nisan 23", "Mayıs 23", "Haziran 23", "Temmuz 23", "Ağustos 23", "Eylül 23", "Ekim 23"],
           "PD": [4, 3.8, 3.6, 3.4, 3.2, 2.3, 2.8, 2.6, 2.4, 2.2]}
data_memzuc_risk = {"Tarih": ["Ocak 23", "Şubat 23", "Mart 23", "Nisan 23", "Mayıs 23", "Haziran 23", "Temmuz 23", "Ağustos 23", "Eylül 23", "Ekim 23"],
                    "Memzuç Riski Değişimi %": [3, -2, -5, 10, 38, 2, 3, 6, 4, 5]}
data_bank_memzuc_risk = {"Tarih": ["Ocak 23", "Şubat 23", "Mart 23", "Nisan 23", "Mayıs 23", "Haziran 23", "Temmuz 23", "Ağustos 23", "Eylül 23", "Ekim 23"],
                         "Bankamız Memzuc Riski Oranı": [20, 15, 30, 12, 24, 40, 50, 60, 40, 50]}
data_tbe_tkn = {"Tarih": ["Ocak 23", "Şubat 23", "Mart 23", "Nisan 23", "Mayıs 23", "Haziran 23", "Temmuz 23", "Ağustos 23", "Eylül 23", "Ekim 23"],
                "TBE": [250, 255, 300, 330, 300, 250, 200, 150, 100, 100],
                "TKN": [300, 310, 350, 380, 330, 253, 180, 160, 150, 154]}

df_pd = pd.DataFrame(data_pd)
df_memzuc_risk = pd.DataFrame(data_memzuc_risk)
df_bank_memzuc_risk = pd.DataFrame(data_bank_memzuc_risk)
df_tbe_tkn = pd.DataFrame(data_tbe_tkn)


fig_pd = px.line(df_pd, x='Tarih', y='PD', title='PD Trend')
fig_memzuc_risk = px.line(df_memzuc_risk, x='Tarih', y='Memzuç Riski Değişimi %', title='Memzuç Riski Değişimi %')
fig_bank_memzuc_risk = px.line(df_bank_memzuc_risk, x='Tarih', y='Bankamız Memzuc Riski Oranı', title='Bankamız Memzuc Riski Oranı')
fig_tbe_tkn = px.line(df_tbe_tkn, x='Tarih', y=['TBE', 'TKN'], title='TBE ve TKN Trendleri')


def style_figure(fig):
    fig.update_layout(
        plot_bgcolor='white',  # Set the background color to white
        height=300,  # Set a smaller height for the figure
        title_x=0.5,  # Center the title
        title_font_size=20,  # Set a suitable font size for the title
        margin=dict(l=20, r=20, t=70, b=40),  # Reduce the margin to make the chart more prominent
    )

    # Update axes styles
    fig.update_xaxes(
        showgrid=True,
        tickangle=45,  # Angle the x-axis tick labels if needed
        title_font=dict(size=16)  # Style for the axis title
    )

    fig.update_yaxes(
        showgrid=True,  # Show the y-axis grid lines for better readability
        gridcolor='lightgrey',  # Set the grid line color to light grey
        title_font=dict(size=16)  # Style for the axis title
    )

    # Update the legend to a more subtle and professional style
    fig.update_layout(legend=dict(
        orientation="h",
        yanchor="bottom",
        y=1.02,
        xanchor="right",
        x=1
    ))

    return fig


# Apply the style to all your figures
fig_pd = style_figure(fig_pd)
fig_memzuc_risk = style_figure(fig_memzuc_risk)
fig_bank_memzuc_risk = style_figure(fig_bank_memzuc_risk)
fig_tbe_tkn = style_figure(fig_tbe_tkn)

fig_pd.update_xaxes(title_text='')
fig_memzuc_risk.update_xaxes(title_text='')
fig_bank_memzuc_risk.update_xaxes(title_text='')
fig_tbe_tkn.update_xaxes(title_text='')

excel_path = 'Banka iç bilgiler.xlsx'
excel_data = pd.ExcelFile(excel_path)
sheet_names = excel_data.sheet_names

def create_table_from_sheet(sheet_name):
    df = pd.read_excel(excel_path, sheet_name=sheet_name)

    # Sütun başlıklarını kontrol et ve datetime olanları düzelt
    df.columns = [col.strftime('%Y-%m-%d') if isinstance(col, datetime.datetime) else col for col in df.columns]

    return dash_table.DataTable(
        id=f'table-{sheet_name}',
        columns=[{"name": str(i), "id": str(i)} for i in df.columns],
        data=df.to_dict('records'),
        style_table={'maxWidth': '960px', 'margin': 'auto'},
        style_cell={
            'textAlign': 'center',
            'padding': '5px',
            'fontSize': '16px',
            'fontFamily': 'Arial',
            'whiteSpace': 'pre-line'
        }
    )

# Akordiyon öğelerini oluşturma
banka_ici_bilgiler_accordion = dbc.Accordion(
    [dbc.AccordionItem(
        title=sheet_name,
        children=create_table_from_sheet(sheet_name)
    ) for sheet_name in excel_data.sheet_names],
    start_collapsed=True)

# Function to read and format the Excel file
def read_and_format_excel(excel_path):
    # Read the Excel file
    excel_data = pd.ExcelFile(excel_path)

    # Create a dictionary to store formatted dataframes
    formatted_dfs = {}

    for sheet_name in excel_data.sheet_names:
        # Read each sheet into a dataframe
        df = pd.read_excel(excel_path, sheet_name=sheet_name)

        # Convert datetime columns to date strings and remove time part if it is 00:00:00
        for col in df.select_dtypes(include=['datetime64', 'datetime']).columns:
            df[col] = df[col].dt.strftime('%Y-%m-%d')

        # Convert index to string if it's a datetime
        if isinstance(df.index, pd.DatetimeIndex):
            df.index = df.index.strftime('%Y-%m-%d')

        # Ensure column names are strings
        df.columns = df.columns.map(str)

        # Remove decimals and add thousand separator to numeric columns
        for col in df.select_dtypes(include=['number']).columns:
            # Round to the nearest whole number and convert to int
            df[col] = df[col].fillna(0).astype(int)
            # Add thousand separator
            df[col] = df[col].apply(lambda x: "{:,}".format(x))

        # Store the formatted dataframe
        formatted_dfs[sheet_name] = df

    return formatted_dfs

# Function to create Dash DataTable from dataframe
def create_dash_datatable(sheet_name, df):
    return dash_table.DataTable(
        id=f'table-{sheet_name}',
        columns=[{"name": str(i), "id": str(i)} for i in df.columns],
        data=df.to_dict('records'),
        style_table={'maxWidth': '960px', 'margin': 'auto'},
        style_cell={
            'textAlign': 'center',
            'padding': '5px',
            'fontSize': '16px',
            'fontFamily': 'Arial',
            'whiteSpace': 'pre-line'
        },
        style_cell_conditional=[
            # Right-align the numeric columns
            {'if': {'column_id': c}, 'textAlign': 'right'} for c in df.select_dtypes(include=['number']).columns
        ]
    )

# Function to create an accordion with DataTables for each sheet in the Excel file
def create_accordion_with_tables(formatted_dfs):
    accordion_items = []
    for sheet_name, df in formatted_dfs.items():
        accordion_item = dbc.AccordionItem(
            title=sheet_name,
            children=create_dash_datatable(sheet_name, df)
        )
        accordion_items.append(accordion_item)

    return dbc.Accordion(accordion_items)

# Read and format the Excel file
formatted_dfs = read_and_format_excel('Banka iç bilgiler.xlsx')

# Create the accordion
banka_ici_bilgiler_accordion = create_accordion_with_tables(formatted_dfs)

formatted_dfs2 = read_and_format_excel('Çek Haciz TSPS.xlsx')

# Create the accordion
cek_haciz_tsps = create_accordion_with_tables(formatted_dfs2)

formatted_dfs3 = read_and_format_excel('ticari kredi bilgileri.xlsx')

# Create the accordion
ticari_kredi_akordiyon = create_accordion_with_tables(formatted_dfs3)

app = Dash(__name__, external_stylesheets=[dbc.themes.BOOTSTRAP])
server = app.server

app.layout = html.Div([
    # This div represents the grey background
    html.Div(style={'backgroundColor': '#f8f9fa', 'minHeight': '100vh'}, children=[
        # This is the container for your content, centered with max width
        dbc.Container(fluid=True, children=[
            # Header with logo
            dbc.Row([
                dbc.Col(html.H1("Kredi Karar Destek Sistemi", style={'color': 'white', 'margin': '0'}),
                        width=9, style={'padding': '10px 20px', 'backgroundColor': '#FF6200'}),
                dbc.Col(html.Img(src=app.get_asset_url('ING_Logo.png'), height="50px"), width=3,
                        style={'padding': '10px 20px', 'backgroundColor': '#FF6200', 'textAlign': 'right'}),
            ]),
            html.Br(),
            # Company information table
            dbc.Row([
                dbc.Col(company_info_table, width=12)
            ]),
            html.Hr(),
            dbc.Row(dbc.Col(html.H5("A - KURALLAR"), width=12)),
            # Accordion for the sheets data
            dbc.Accordion(
                [dbc.AccordionItem(
                    dash_table.DataTable(
                        id=f'table-{sheet_name}',
                        columns=[{"name": i, "id": i} for i in df.columns],
                        data=df.to_dict('records'),
                        style_table={'maxWidth': '960px', 'margin': 'auto'},
                        style_cell={
                            'textAlign': 'center',
                            'padding': '5px',
                            'fontSize': '16px',
                            'fontFamily': 'Arial',
                            'whiteSpace': 'pre-line'
                        }
                    ),
                    title=sheet_name
                ) for sheet_name, df in sheet_to_df_map.items()],
                start_collapsed=True
            ),
            html.Hr(),
            dbc.Row(dbc.Col(html.H5("B - GÖSTERGELER"), width=12)),
            dbc.Accordion([dbc.AccordionItem(title="Matrisler",
                                             children=[dbc.Row([
                                                 dbc.Col(dcc.Graph(id='matrix-heatmap', figure=fig), width=4),
                                                 dbc.Col(dcc.Graph(id='matrix-heatmap2', figure=fig2), width=4),
                                                 dbc.Col(dcc.Graph(id='matrix-heatmap3', figure=fig3), width=4)
                                             ]), ]),
                           ], start_collapsed=True),
            dbc.Accordion([
                dbc.AccordionItem(title="Trendler", children=[
                    dbc.Row([
                        dbc.Col(dcc.Graph(figure=fig_pd), width=6),
                        dbc.Col(dcc.Graph(figure=fig_memzuc_risk), width=6)
                    ]),
                    dbc.Row([
                        dbc.Col(dcc.Graph(figure=fig_bank_memzuc_risk), width=6),
                        dbc.Col(dcc.Graph(figure=fig_tbe_tkn), width=6)
                    ])
                ])
            ], start_collapsed=True),

            html.Hr(),
            dbc.Row(dbc.Col(html.H5("C - BANKA İÇİ VERİLER"), width=12)),
            banka_ici_bilgiler_accordion,

            html.Hr(),
            dbc.Row(dbc.Col(html.H5("D - ÇEK/HACİZ/TSPS VERİLERİ"), width=12)),
            cek_haciz_tsps,

            html.Hr(),
            dbc.Row(dbc.Col(html.H5("E - KREDİ VERİLERİ / KKB-MEMZUÇ"), width=12)),
            ticari_kredi_akordiyon

        ], style={'maxWidth': '1080px', 'margin': 'auto', 'padding': '20px', 'marginTop': '20px',
                  'boxShadow': '0 4px 8px 0 rgba(0,0,0,0.2)'})
    ])
])

if __name__ == '__main__':
    app.run_server()