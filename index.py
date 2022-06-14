import dash_core_components as dcc
import dash_html_components as html
from dash.dependencies import Input, Output
from waitress import serve
from paste.translogger import TransLogger

from app import app
from apps import app_home


app.layout = html.Div([
    dcc.Location(id='url', refresh=True),
    html.Div(id='page-content')
])


@app.callback(Output('page-content', 'children'),
              Input('url', 'pathname'))
def display_page(pathname):
    return app_home.serve_layout()


if __name__ == '__main__':
    app.run_server(debug=True, host='AUD0100CK4', port=8085)
    # serve(TransLogger(app.server, logging_level=30), host='AUD0100CK4', port=8085)