import dash_table
import dash_core_components as dcc
import dash_html_components as html
import dash_bootstrap_components as dbc
from dash.dependencies import Input, Output, State
from dash.exceptions import PreventUpdate
import pandas as pd
import win32com
from datetime import datetime
import json
import requests
from waitress import serve
from paste.translogger import TransLogger

from app import app
from iress import Iress


# declare global variables
interval_threshold = 900
tick_threshold = 1
inav_threshold = 10
thresholds = {
    'Bid Spread to iNAV (ticks)': tick_threshold,
    'Ask Spread to iNAV (ticks)': tick_threshold,
    'iNAV Diff (bps)': inav_threshold
}

ice_url = 'https://iml.factsetdigitalsolutions.com/application/index/quote?t=LSGE'
sol_url = 'https://clients.solactive.com/api/rest/v1/indices/4386924db2b1d848621a188a90a3a855/DE000SL0DQU2/performance'
etf_codes = ['LSGE']
codes = ['LSGE', 'LSGEAUDINAV', 'SPFUT']
exchanges = ['AXW', 'ETF', 'ID']
input_dict = {
    'SecurityCode': codes,
    'Exchange': exchanges,
}

method = 'pricingquoteexget'
fields = ['SecurityCode', 'BidPrice', 'AskPrice', 'LastPrice', 'MovementPercent']
etf_fields = ['SecurityCode', 'BidPrice', 'AskPrice', 'LastPrice', 'MovementPercent']
inav_fields = ['SecurityCode', 'LastPrice']
futures_fields = ['SecurityCode', 'MovementPercent']
table_fields = ['SecurityCode', 'BidPrice', 'AskPrice', 'LastPrice', 'MovementPercent', 'ICE iNAV', 'Solactive iNAV',
             'iNAV Diff (bps)', 'Bid Spread to iNAV (ticks)', 'Ask Spread to iNAV (ticks)', 'SP500 Futures %']


def serve_layout():

    layout = dbc.Container([

        # interval component
        dcc.Interval(
            id='interval',
            interval=1 * 1000,  # in milliseconds
            n_intervals=0
        ),

        # storage component to store last interval when tick threshold was exceeded
        dcc.Store(
            id='store',
            data={
                'Bid Spread to iNAV (ticks)': 0,
                'Ask Spread to iNAV (ticks)': 0,
                'iNAV Diff (bps)': 0
            }
        ),

        # Top Banner
        dbc.Navbar([
            html.Div(
                dbc.Row([
                    dbc.Col(html.H1('IML ETFs'))
                ])
            ),

            dbc.Row([
                dbc.Col(html.Img(src=app.get_asset_url('IML_Logo.png'), height="40px")),
            ],
                className="ml-auto flex-nowrap mt-3 mt-md-0",
            )
        ]),

        # table with links to indicators
        dbc.Row([
            dbc.Col(
                dash_table.DataTable(
                    id='table-etf',
                    columns=[{'name': i, 'id': i} for i in table_fields],
                    style_cell={
                        'textAlign': 'left',
                    },
                    style_data_conditional=[
                        {
                            'if': {
                                'filter_query': '{Bid Spread to iNAV (ticks)} > ' + str(tick_threshold),
                                'column_id': 'Bid Spread to iNAV (ticks)'
                            },
                            'backgroundColor': 'tomato',
                            'color': 'white'
                        },
                        {
                            'if': {
                                'filter_query': '{Ask Spread to iNAV (ticks)} > ' + str(tick_threshold),
                                'column_id': 'Ask Spread to iNAV (ticks)'
                            },
                            'backgroundColor': 'tomato',
                            'color': 'white'
                        },
                        {
                            'if': {
                                'filter_query': '{iNAV Diff (bps)} > ' + str(inav_threshold),
                                'column_id': 'iNAV Diff (bps)'
                            },
                            'backgroundColor': 'tomato',
                            'color': 'white'
                        }
                    ]
                )
            )
        ]),

        # create charts for indicators vs share price
        dbc.Row(
            id='chart-layout',
        )
    ],
        fluid=True
    )

    return layout


@app.callback(Output('table-etf', 'data'),
              Output('store', 'data'),
              Input('interval', 'n_intervals'),
              State('store', 'data'))
def update_etfs(n, last_interval):
    if last_interval is None:
        raise PreventUpdate

    iress_obj = Iress(method, fields, input_dict)
    iress_obj.set_inputs()
    iress_obj.execute()
    data = iress_obj.retrieve_data()
    df = pd.DataFrame(data=data, columns=fields)
    # print(df)

    etf_filter = df['SecurityCode'].isin(etf_codes)
    df.loc[etf_filter, ['BidPrice', 'AskPrice', 'LastPrice']] = df.loc[etf_filter, ['BidPrice', 'AskPrice', 'LastPrice']].div(100)
    df['MovementPercent'] = df['MovementPercent'].round(2)
    ice_res = requests.get(ice_url)
    ice_json_data = json.loads(ice_res.text)
    ice_json_data = ice_json_data['quote']['etf'][0]
    df['ICE iNAV'] = float(ice_json_data['inav'])
    df['Solactive iNAV'] = df[df['SecurityCode'] == 'LSGEAUDINAV']['LastPrice'].values[0]
    df['SP500 Futures %'] = df[df['SecurityCode'] == 'SPFUT']['MovementPercent'].values[0]
    df['iNAV Diff (bps)'] = round(abs((df['ICE iNAV'] / df['Solactive iNAV'] - 1)*10000), 0)
    df['Bid Spread to iNAV (ticks)'] = round((round(df['ICE iNAV'], 2) - df['BidPrice']) * 100, 0)
    df['Ask Spread to iNAV (ticks)'] = round((df['AskPrice'] - round(df['ICE iNAV'], 2)) * 100, 0)
    df = df[df['SecurityCode'].isin(etf_codes)][table_fields]

    for col in ['Bid Spread to iNAV (ticks)', 'Ask Spread to iNAV (ticks)', 'iNAV Diff (bps)']:
        if (df[col] > thresholds[col]).any():
            print(col, n - last_interval[col])
            if n - last_interval[col] > interval_threshold:
                last_interval[col] = n
                email_alert(df[df[col] > thresholds[col]], col)
                print(col, last_interval)

    return df.to_dict('records'), last_interval


def email_alert(data, alert_type):
    Outlook = win32com.client.Dispatch("Outlook.Application")
    objMail = Outlook.CreateItem(0)
    objMail.To = 'jeremy.tang@iml.com.au; richard.rudenko@iml.com.au'
    objMail.Subject = alert_type + ' Alert: ' + datetime.now().strftime('%d-%b-%Y %H:%M:%S')
    html = 'IML ETF Dashboard: <br><a href=http://aud0100ck4:8085>http://aud0100ck4:8085</a><br><br>' + \
           data.to_html(index=False)
    objMail.HTMLBody = html
    objMail.Send()
    # objMail.Display()