from dash import Dash, dcc, html, Input, Output, State, dash_table
import dash_bootstrap_components as dbc 
import pandas as pd
import plotly.express as px
import io
import base64
from datetime import datetime
import os
# Initialize Dash app
app = Dash(__name__, external_stylesheets=[dbc.themes.FLATLY])
app.title = "Excel Analysis Dashboard"

# App Layout
app.layout = dbc.Container([
    # Header
    dbc.Row([
        dbc.Col(html.H1("Excel Analysis Dashboard", className="text-center text-primary mb-4"))
    ]),

    # Main Tabs
    dbc.Tabs([
        # Tab 1: Split Excel
        dbc.Tab(label="Split Excel", children=[
            dbc.Card(dbc.CardBody([
                dbc.Row([
                    dbc.Col([
                        html.H4("Upload Excel File"),
                        dcc.Upload(
                            id='upload-data',
                            children=dbc.Card(
                                dbc.CardBody([
                                    html.I(className="fas fa-file-upload fa-2x mb-3"),
                                    html.H5("Drag and Drop or Click to Upload")
                                ], className="text-center"),
                                className="border-dashed"
                            ),
                            className="mb-3"
                        ),
                        html.Div(id='upload-status'),
                        dbc.Button(
                            "Download Processed File",
                            id="btn-download",
                            color="primary",
                            className="w-100 mt-3",
                            disabled=True
                        ),
                        dcc.Download(id="download-dataframe-xlsx")
                    ])
                ]),
                html.Div(id='sheet-counts', className="mt-4")
            ]))
        ]),

        # Tab 2: Analytics
        dbc.Tab(label="Analytics", children=[
            dbc.Card(dbc.CardBody([
                # Summary Statistics
                dbc.Row([
                    dbc.Col(dbc.Card(dbc.CardBody([
                        html.H5("Total Pre-available", className="text-muted"),
                        html.H2(id="total-count", className="text-primary text-center")
                    ])), width=12, md=4),
                    dbc.Col(dbc.Card(dbc.CardBody([
                        html.H5("Live In (CC)", className="text-muted"),
                        html.H2(id="live-in-count", className="text-success text-center")
                    ])), width=12, md=4),
                    dbc.Col(dbc.Card(dbc.CardBody([
                        html.H5("Live Out", className="text-muted"),
                        html.H2(id="live-out-count", className="text-info text-center")
                    ])), width=12, md=4),
                ], className="mb-4"),

                # Nationality Distribution Chart
                dbc.Card([
                    dbc.CardHeader("Nationality Distribution by Live In/Out"),
                    dbc.CardBody(dcc.Graph(id="nationality-chart"))
                ], className="mb-4"),

                # Visa Alerts and Risk Tables
                dbc.Card([
                    dbc.CardHeader([
                        html.H5("Visa Alerts", className="mb-0 text-danger"),
                        html.Small("Entry Visa: >3 days, Tourist Visa: >8 days", className="text-muted")
                    ]),
                    dbc.CardBody([
                        dash_table.DataTable(
                            id='visa-alerts-table',
                            style_data_conditional=[
                                {
                                    'if': {'row_index': 'odd'},
                                    'backgroundColor': 'rgba(248, 249, 250, 0.5)'
                                }
                            ],
                            style_header={
                                'backgroundColor': '#f8f9fa',
                                'fontWeight': 'bold'
                            },
                            style_cell={
                                'textAlign': 'left',
                                'padding': '12px'
                            },
                            sort_action='native'
                        )
                    ])
                ], className="mb-4"),
                
                # At Risk Cases
                dbc.Card([
                    dbc.CardHeader([
                        html.H5("At Risk Cases", className="mb-0 text-warning"),
                        html.Small("Entry Visa: 2 days, Tourist Visa: 7 days", className="text-muted")
                    ]),
                    dbc.CardBody([
                        dash_table.DataTable(
                            id='at-risk-table',
                            style_data_conditional=[
                                {
                                    'if': {'row_index': 'odd'},
                                    'backgroundColor': 'rgba(248, 249, 250, 0.5)'
                                }
                            ],
                            style_header={
                                'backgroundColor': '#f8f9fa',
                                'fontWeight': 'bold'
                            },
                            style_cell={
                                'textAlign': 'left',
                                'padding': '12px'
                            },
                            sort_action='native'
                        )
                    ])
                ])
            ]))
        ]),

        # Tab 3: Name Comparison
        dbc.Tab(label="Name Comparison", children=[
            dbc.Card(dbc.CardBody([
                dbc.Row([
                    dbc.Col([
                        html.H5("First File"),
                        dcc.Upload(
                            id='upload-file-1',
                            children=dbc.Card(
                                dbc.CardBody([
                                    html.I(className="fas fa-file-upload fa-2x mb-3"),
                                    html.H6("Upload First File")
                                ], className="text-center"),
                                className="border-dashed"
                            )
                        ),
                        html.Div(id='file1-status')
                    ], width=12, md=6),
                    dbc.Col([
                        html.H5("Second File"),
                        dcc.Upload(
                            id='upload-file-2',
                            children=dbc.Card(
                                dbc.CardBody([
                                    html.I(className="fas fa-file-upload fa-2x mb-3"),
                                    html.H6("Upload Second File")
                                ], className="text-center"),
                                className="border-dashed"
                            )
                        ),
                        html.Div(id='file2-status')
                    ], width=12, md=6),
                ], className="mb-4"),

                dbc.Card([
                    dbc.CardHeader("Matching Names"),
                    dbc.CardBody(
                        dash_table.DataTable(
                            id='matching-names-table',
                            style_data_conditional=[
                                {
                                    'if': {'row_index': 'odd'},
                                    'backgroundColor': 'rgba(248, 249, 250, 0.5)'
                                }
                            ],
                            style_header={
                                'backgroundColor': '#f8f9fa',
                                'fontWeight': 'bold'
                            },
                            style_cell={
                                'textAlign': 'left',
                                'padding': '12px'
                            },
                            sort_action='native'
                        )
                    )
                ])
            ]))
        ])
    ])
], fluid=True)

def parse_contents(contents, filename):
    """Parse uploaded Excel file contents"""
    content_type, content_string = contents.split(',')
    decoded = base64.b64decode(content_string)
    
    try:
        if 'xlsx' not in filename.lower():
            return None, "Please upload an Excel file."
            
        df = pd.read_excel(io.BytesIO(decoded))
        return df, None
    except Exception as e:
        return None, str(e)

def process_data(df):
    """Process data into categories based on visa step and other conditions"""
    # Initialize categories
    categories = {}
    
    # Special handling for medical and bio appointments
    both_medical_and_bio = df[
        df['Current Visa Step'].str.contains('Pending maid to go for EID Biometrics', na=False) & 
        df['Current Visa Step'].str.contains('Waiting for the maid to go to medical test and EID fingerprinting', na=False)
    ]
    
    only_medical = df[
        ~df.index.isin(both_medical_and_bio.index) & 
        df['Current Visa Step'].str.contains('Waiting for the maid to go to medical test and EID fingerprinting', na=False)
    ]
    
    only_bio = df[
        ~df.index.isin(both_medical_and_bio.index) & 
        df['Current Visa Step'].str.contains('EID fingerprinting', na=False)
    ]
    
    # Add these to categories
    categories['Push for medical and book bio'] = both_medical_and_bio
    categories['Push for medical'] = only_medical
    categories['Push to book bio appointment'] = only_bio
    
    # Other categories
    categories['As Aya to book Bio Appointment'] = df[
        ~df.index.isin(both_medical_and_bio.index) & 
        ~df.index.isin(only_medical.index) & 
        ~df.index.isin(only_bio.index) & 
        df['Current Visa Step'].str.contains('Prepare EID application', na=False)
    ]
    
    categories['Apply Entry Visa'] = df[
        df['Current Visa Step'].str.contains('Apply for entry Visa', na=False)
    ]
    
    categories['Create Offer Letter'] = df[
        df['Current Visa Step'].str.contains('Create Regular Offer Letter', na=False)
    ]
    
    categories['Check Complaints'] = df[
        df['Current Visa Step'].str.contains('Waiting for the PRO Update|Pending to fix MOHRE issue', 
                                           na=False, regex=True)
    ]
    
    categories['Coaches'] = df[
        df['pending arrival task'].str.contains('STAND_UP_SHOOTING|MATCHING TYPES AND DATA GATHERING', 
                                              na=False, regex=True)
    ]
    
    categories['Onboarding'] = df[
        df['pending arrival task'].str.contains('TAWJEEH_TRAINING|ORIENTATION|UPLOAD_CERTIFICATE|MAID_INFO', 
                                              na=False, regex=True)
    ]
    
    categories['Media'] = df[
        df['pending arrival task'].str.contains('VIDEO_EDITING', na=False)
    ]
    
    categories['Apply for GCC'] = df[
        (df['Nationality'].isin(['Ugandan', 'Kenyan'])) & 
        (df['GCC'] == 'No') & 
        df['GCC Application Reference Number Upload Date'].isna()
    ]
    
    # Add remaining rows to "unmatched" category
    all_matched = pd.concat([df_ for df_ in categories.values()]).index
    categories['Unmatched'] = df[~df.index.isin(all_matched)]
    
    return categories

def calculate_days_since_landing(landing_date):
    """Calculate days since landing date"""
    if pd.isna(landing_date):
        return None
    try:
        landing = pd.to_datetime(landing_date)
        today = pd.Timestamp.now()
        return (today - landing).days
    except:
        return None

@app.callback(
    [Output('total-count', 'children'),
     Output('live-in-count', 'children'),
     Output('live-out-count', 'children'),
     Output('nationality-chart', 'figure'),
     Output('visa-alerts-table', 'data'),
     Output('visa-alerts-table', 'columns'),
     Output('at-risk-table', 'data'),
     Output('at-risk-table', 'columns'),
     Output('upload-status', 'children'),
     Output('btn-download', 'disabled'),
     Output('sheet-counts', 'children')],
    [Input('upload-data', 'contents')],
    [State('upload-data', 'filename')]
)
def update_analytics(contents, filename):
    """Update analytics based on uploaded file"""
    if contents is None:
        return "0", "0", "0", {}, [], [], [], [], "", True, ""
    
    df, error = parse_contents(contents, filename)
    if error:
        return "0", "0", "0", {}, [], [], [], [], html.Div(error, style={'color': 'red'}), True, ""
    
    # Process data
    categories = process_data(df)
    
    # Calculate counts
    total_count = len(df)
    df['Live out type'] = df['Live out type'].fillna('')
    live_in_count = len(df[df['Live out type'].str.strip() == 'CC'])
    live_out_count = len(df[df['Live out type'].str.strip() == 'CC (Live out)'])
    
    # Create nationality chart
    nationality_data = []
    for nat in df['Nationality'].unique():
        if pd.notna(nat):
            live_in = len(df[(df['Nationality'] == nat) & (df['Live out type'].str.strip() == 'CC')])
            live_out = len(df[(df['Nationality'] == nat) & (df['Live out type'].str.strip() == 'CC (Live out)')])
            if live_in > 0 or live_out > 0:
                nationality_data.append({
                    'Nationality': nat,
                    'Live In': live_in,
                    'Live Out': live_out
                })
    
    fig = px.bar(nationality_data, 
                 x='Nationality',
                 y=['Live In', 'Live Out'],
                 barmode='group',
                 title='Nationality Distribution',
                 color_discrete_sequence=['#28a745', '#17a2b8'])
    
    # Process visa alerts and at-risk cases
    df['Days Since Landing'] = df['Landed In Dubai'].apply(calculate_days_since_landing)
    
    # Define thresholds
    entry_visa_alert_threshold = 3
    entry_visa_risk_threshold = 2
    tourist_visa_alert_threshold = 8
    tourist_visa_risk_threshold = 7
    
    # Process exceeded threshold (alerts)
    visa_alerts = df[
        ((df['Type of Visa'] == 'Entry Visa') & (df['Days Since Landing'] > entry_visa_alert_threshold)) |
        ((df['Type of Visa'].isin(['Tourist Visa', '']) | df['Type of Visa'].isna()) & 
         (df['Days Since Landing'] > tourist_visa_alert_threshold))
    ].sort_values('Days Since Landing', ascending=False)
    
    # Process at-risk cases (1 day before alert threshold)
    at_risk = df[
        ((df['Type of Visa'] == 'Entry Visa') & (df['Days Since Landing'] == entry_visa_risk_threshold)) |
        ((df['Type of Visa'].isin(['Tourist Visa', '']) | df['Type of Visa'].isna()) & 
         (df['Days Since Landing'] == tourist_visa_risk_threshold))
    ].sort_values('Days Since Landing', ascending=False)
    
    alerts_data = []
    for _, row in visa_alerts.iterrows():
        alerts_data.append({
            'Housemaid Name': row['Housemaid Name'],
            'Type of Visa': row['Type of Visa'],
            'Landed In Dubai': row['Landed In Dubai'],
            'Days Since Landing': row['Days Since Landing'],
            'Live Type': 'Live In' if str(row['Live out type']).strip() == 'CC' 
                        else 'Live Out' if str(row['Live out type']).strip() == 'CC (Live out)' 
                        else 'Unknown',
            'Visa Step': str(row['Current Visa Step'])
        })
    
    at_risk_data = []
    for _, row in at_risk.iterrows():
        at_risk_data.append({
            'Housemaid Name': row['Housemaid Name'],
            'Type of Visa': row['Type of Visa'],
            'Landed In Dubai': row['Landed In Dubai'],
            'Days Since Landing': row['Days Since Landing'],
            'Live Type': 'Live In' if str(row['Live out type']).strip() == 'CC' 
                        else 'Live Out' if str(row['Live out type']).strip() == 'CC (Live out)' 
                        else 'Unknown',
            'Visa Step': str(row['Current Visa Step']),
            'Threshold': f"{entry_visa_alert_threshold} days" if row['Type of Visa'] == 'Entry Visa' 
                        else f"{tourist_visa_alert_threshold} days"
        })
    
    alerts_columns = [
        {"name": "Name", "id": "Housemaid Name"},
        {"name": "Visa Type", "id": "Type of Visa"},
        {"name": "Landing Date", "id": "Landed In Dubai"},
        {"name": "Days Since Landing", "id": "Days Since Landing"},
        {"name": "Live Type", "id": "Live Type"},
        {"name": "Visa Step", "id": "Visa Step"}
    ]
    
    at_risk_columns = [
        {"name": "Name", "id": "Housemaid Name"},
        {"name": "Visa Type", "id": "Type of Visa"},
        {"name": "Landing Date", "id": "Landed In Dubai"},
        {"name": "Days Since Landing", "id": "Days Since Landing"},
        {"name": "Live Type", "id": "Live Type"},
        {"name": "Visa Step", "id": "Visa Step"},
        {"name": "Threshold", "id": "Threshold"}
    ]
    
    # Create sheet counts display
    sheet_counts = dbc.Card([
        dbc.CardHeader("Records per Category"),
        dbc.CardBody([
            dbc.Row([
                dbc.Col(
                    dbc.Card(dbc.CardBody([
                        html.H6(name, className="mb-2"),
                        html.H4(str(len(df_cat)), className="text-primary text-center")
                    ])),
                    width=12, md=6, lg=4, className="mb-3"
                )
                for name, df_cat in categories.items()
                if len(df_cat) > 0
            ])
        ])
    ])
    
    return (
        f"{total_count:,}",
        f"{live_in_count:,}",
        f"{live_out_count:,}",
        fig,
        alerts_data,
        alerts_columns,
        at_risk_data,
        at_risk_columns,
        html.Div("File processed successfully!", className="text-success"),
        False,
        sheet_counts
    )

@app.callback(
    Output("download-dataframe-xlsx", "data"),
    Input("btn-download", "n_clicks"),
    State('upload-data', 'contents'),
    State('upload-data', 'filename'),
    prevent_initial_call=True
)
def download_excel(n_clicks, contents, filename):
    """Create and download processed Excel file"""
    if contents is None:
        return None
    
    df, error = parse_contents(contents, filename)
    if error:
        return None
    
    categories = process_data(df)
    
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        for sheet_name, df_sheet in categories.items():
            if not df_sheet.empty:
                df_sheet.to_excel(writer, sheet_name=sheet_name[:31], index=False)
    
    return dcc.send_bytes(output.getvalue(), f"processed_{filename}")

@app.callback(
    [Output('matching-names-table', 'data'),
     Output('matching-names-table', 'columns'),
     Output('file1-status', 'children'),
     Output('file2-status', 'children')],
    [Input('upload-file-1', 'contents'),
     Input('upload-file-2', 'contents')],
    [State('upload-file-1', 'filename'),
     State('upload-file-2', 'filename')]
)
def compare_files(contents1, contents2, filename1, filename2):
    """Compare names between two files"""
    if contents1 is None or contents2 is None:
        return [], [], "", ""
    
    # Parse both files
    df1, error1 = parse_contents(contents1, filename1)
    df2, error2 = parse_contents(contents2, filename2)
    
    if error1:
        return [], [], html.Div(error1, style={'color': 'red'}), ""
    if error2:
        return [], [], "", html.Div(error2, style={'color': 'red'})
    
    # Find matching names with same visa step
    matches = []
    
    # Check for required columns in both dataframes
    required_columns = ['Housemaid Name', 'Current Visa Step', 'Live out type']
    missing_columns_df1 = [col for col in required_columns if col not in df1.columns]
    missing_columns_df2 = [col for col in required_columns if col not in df2.columns]
    
    if missing_columns_df1 or missing_columns_df2:
        error_msg = "Missing required columns: "
        if missing_columns_df1:
            error_msg += f"\nFile 1: {', '.join(missing_columns_df1)}"
        if missing_columns_df2:
            error_msg += f"\nFile 2: {', '.join(missing_columns_df2)}"
        return [], [], html.Div(error_msg, style={'color': 'red'}), html.Div(error_msg, style={'color': 'red'})
    
    for _, row1 in df1.iterrows():
        name1 = row1['Housemaid Name']
        step1 = row1['Current Visa Step']
        
        # Find matching rows in second file
        matching_rows = df2[
            (df2['Housemaid Name'] == name1) & 
            (df2['Current Visa Step'] == step1)
        ]
        
        if not matching_rows.empty:
            # Get live type for both files
            live_type_1 = 'Live In' if str(row1['Live out type']).strip() == 'CC' else 'Live Out' if str(row1['Live out type']).strip() == 'CC (Live out)' else 'Unknown'
            live_type_2 = 'Live In' if str(matching_rows.iloc[0]['Live out type']).strip() == 'CC' else 'Live Out' if str(matching_rows.iloc[0]['Live out type']).strip() == 'CC (Live out)' else 'Unknown'
            
            match_data = {
                'Name': name1,
                'Current Visa Step': step1,
                'Live Type (File 1)': live_type_1,
                'Live Type (File 2)': live_type_2
            }
            
            # Add IDs if available
            if 'Housemaid Id' in df1.columns:
                match_data['File 1 ID'] = row1['Housemaid Id']
            if 'Housemaid Id' in df2.columns:
                match_data['File 2 ID'] = matching_rows.iloc[0]['Housemaid Id']
                
            matches.append(match_data)
    
    if not matches:
        success_msg = html.Div([
            html.I(className="fas fa-check-circle me-2"),
            "Files processed - No matches found"
        ], className="text-warning")
        return [], [], success_msg, success_msg
    
    # Create columns based on available data
    columns = [
        {'name': 'Name', 'id': 'Name'},
        {'name': 'Current Visa Step', 'id': 'Current Visa Step'},
        {'name': 'Live Type (File 1)', 'id': 'Live Type (File 1)'},
        {'name': 'Live Type (File 2)', 'id': 'Live Type (File 2)'}
    ]
    
    # Add ID columns if available
    if 'Housemaid Id' in df1.columns:
        columns.insert(2, {'name': 'File 1 ID', 'id': 'File 1 ID'})
    if 'Housemaid Id' in df2.columns:
        columns.insert(3 if 'File 1 ID' in [col['id'] for col in columns] else 2, 
                      {'name': 'File 2 ID', 'id': 'File 2 ID'})
    
    success_msg = html.Div([
        html.I(className="fas fa-check-circle me-2"),
        f"Files processed - Found {len(matches)} matches"
    ], className="text-success")
    
    return matches, columns, success_msg, success_msg

# Add custom CSS
app.index_string = '''
<!DOCTYPE html>
<html>
    <head>
        {%metas%}
        <title>{%title%}</title>
        {%favicon%}
        {%css%}
        <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/5.15.4/css/all.min.css">
        <style>
            .border-dashed {
                border-style: dashed !important;
                border-width: 2px !important;
                border-radius: 10px !important;
                background-color: #fafafa !important;
                cursor: pointer;
            }
            .border-dashed:hover {
                background-color: #f0f0f0 !important;
            }
            .card {
                box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
                border: none !important;
                border-radius: 10px !important;
            }
            .card-header {
                background-color: #ffffff !important;
                border-bottom: 1px solid rgba(0,0,0,.125);
                border-top-left-radius: 10px !important;
                border-top-right-radius: 10px !important;
            }
            .dash-table-container .dash-spreadsheet-container .dash-spreadsheet-inner td, 
            .dash-table-container .dash-spreadsheet-container .dash-spreadsheet-inner th {
                border: 1px solid #dee2e6 !important;
            }
            .nav-tabs .nav-link.active {
                font-weight: bold;
                border-bottom: 3px solid #007bff;
            }
        </style>
    </head>
    <body>
        {%app_entry%}
        <footer>
            {%config%}
            {%scripts%}
            {%renderer%}
        </footer>
    </body>
</html>
'''


if __name__ == '__main__':
    port = int(os.environ.get("PORT", 8050))  # Use PORT from env, default to 8050
    app.run_server(debug=False, host="0.0.0.0", port=port)
