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
intervals = {
    'Bid Spread to iNAV (ticks)': 50,
    'Ask Spread to iNAV (ticks)': 50,
    'iNAV Diff (bps)': 50
}

tick_threshold_dict = {
    'LSGE': 1,
    'VNGS': 1
}

inav_threshold_dict = {
    'LSGE': 100,
    'VNGS': 100
}

thresholds = {
    'Bid Spread to iNAV (ticks)': tick_threshold_dict,
    'Ask Spread to iNAV (ticks)': tick_threshold_dict,
    'iNAV Diff (bps)': inav_threshold_dict
}

ice_url = 'https://iml.ppe.factsetdigitalsolutions.com/application/index/quote?t=LSGE,VNGS'
sol_url = 'https://clients.solactive.com/api/rest/v1/indices/4386924db2b1d848621a188a90a3a855/DE000SL0DQU2/performance'
etf_codes = ['LSGE', 'VNGS']
solactive_codes = ['LSGEAUDINAV', 'VNGSAUDINAV'] # change to VN solactive inav code when it goes live
codes = ['LSGE', 'LSGEAUDINAV', 'SPFUT', 'VNGS', 'VNGSAUDINAV'] # change to VN solactive inav code when it goes live
exchanges = ['AXW', 'ETF', 'ID', 'AXW', 'ETF']
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
            interval=4 * 1000,  # in milliseconds
            n_intervals=0
        ),

        # storage component to store last interval when tick threshold was exceeded
        dcc.Store(
            id='store',
            data={etf: {
                    'Bid Spread to iNAV (ticks)': 0,
                    'Ask Spread to iNAV (ticks)': 0,
                    'iNAV Diff (bps)': 0
                }
                for etf in etf_codes
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
                    style_data_conditional=(
                        [
                            {
                                'if': {
                                    'filter_query': '{Bid Spread to iNAV (ticks)} > ' + str(tick_threshold_dict[i]) + ' && {SecurityCode} = ' + i,
                                    'column_id': 'Bid Spread to iNAV (ticks)'
                                },
                                'backgroundColor': 'tomato',
                                'color': 'white'
                            }
                            for i in etf_codes
                        ] +
                        [
                            {
                                'if': {
                                    'filter_query': '{Ask Spread to iNAV (ticks)} > ' + str(tick_threshold_dict[i]) + ' && {SecurityCode} = ' + i,
                                    'column_id': 'Ask Spread to iNAV (ticks)'
                                },
                                'backgroundColor': 'tomato',
                                'color': 'white'
                            }
                            for i in etf_codes
                        ] +
                        [
                            {
                                'if': {
                                    'filter_query': '{iNAV Diff (bps)} > ' + str(inav_threshold_dict[i]) + ' && {SecurityCode} = ' + i,
                                    'column_id': 'iNAV Diff (bps)'
                                },
                                'backgroundColor': 'tomato',
                                'color': 'white'
                            }
                            for i in etf_codes
                        ]
                    )
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
    ice_json_data = ice_json_data['quote']['etf']
    for index, (etf, sol_inav) in enumerate(zip(etf_codes, solactive_codes)):
        df.loc[df['SecurityCode'] == etf, 'Solactive iNAV'] = df[df['SecurityCode'] == sol_inav]['LastPrice'].values[0]
        if ice_json_data[index]['inav'] == '-':
            df.loc[df['SecurityCode'] == etf, 'ICE iNAV'] = 0
        else:
            df.loc[df['SecurityCode'] == etf, 'ICE iNAV'] = float(ice_json_data[index]['inav'])
    df['SP500 Futures %'] = df[df['SecurityCode'] == 'SPFUT']['MovementPercent'].values[0]
    df['iNAV Diff (bps)'] = round(abs((df['ICE iNAV'] / df['Solactive iNAV'] - 1)*10000), 0)
    df['Bid Spread to iNAV (ticks)'] = round((round(df['ICE iNAV'], 2) - df['BidPrice']) * 100, 0)
    df['Ask Spread to iNAV (ticks)'] = round((df['AskPrice'] - round(df['ICE iNAV'], 2)) * 100, 0)
    df = df[df['SecurityCode'].isin(etf_codes)][table_fields]

    '''recode this part for etf specific thresholds'''
    for col in ['Bid Spread to iNAV (ticks)', 'Ask Spread to iNAV (ticks)', 'iNAV Diff (bps)']:
        for etf in etf_codes:
            # print(etf, col, n, last_interval[etf][col])
            df_etf = df[df['SecurityCode'] == etf]
            if (df_etf[col] > thresholds[col][etf]).any():
                print(etf, col, n, last_interval[etf][col])
                if last_interval[etf][col] <= n:
                    last_interval[etf][col] = n
                    if last_interval[etf][col] % intervals[col] == 0 and last_interval[etf][col] != 0 and n != 0:
                        print('email alert')
                        email_alert(df_etf[df_etf[col] > thresholds[col][etf]], col)

                else:
                    last_interval[etf][col] = 0

    del iress_obj

    return df.to_dict('records'), last_interval


def email_alert(data, alert_type):
    Outlook = win32com.client.Dispatch("Outlook.Application")
    objMail = Outlook.CreateItem(0)
    objMail.To = 'jeremy.tang@iml.com.au; richard.rudenko@iml.com.au'
    # objMail.To = 'jeremy.tang@iml.com.au;'
    objMail.Subject = alert_type + ' Alert: ' + datetime.now().strftime('%d-%b-%Y %H:%M:%S')
    html = 'IML ETF Dashboard: <br><a href=http://aud0100ck4:8085>http://aud0100ck4:8085</a><br><br>' + \
           data.to_html(index=False)
    objMail.HTMLBody = html
    objMail.Send()
    # objMail.Display()