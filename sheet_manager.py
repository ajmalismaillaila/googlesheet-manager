from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
import gspread
from gspread_formatting import *
import pandas as pd
import numpy as np

pd.options.mode.chained_assignment = None

letters  = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z']


def authenticate():
    credentials = Credentials.from_service_account_file('secrets/client_secrets.json', scopes=['https://www.googleapis.com/auth/spreadsheets'])
    return credentials

creds = authenticate()
service = build('sheets','v4', credentials=creds)
gc = gspread.authorize(creds)

# data range you can specift the sheet name and the cell range eg: 'Sheet1!A1:B10'
def read_sheet(SHEETS_ID, data_range):
    sheets = service.spreadsheets()
    result = sheets.values().get(spreadsheetId=SHEETS_ID, range = data_range).execute()#
    return result


def create_new_sheet(service, SHEETS_ID, new_sheet_title):
    # Create the new sheet request
    request = {
        'addSheet': {
            'properties': {
                'title': new_sheet_title,
            },
        },
    }

    # Execute the batchUpdate request
    service.spreadsheets().batchUpdate(spreadsheetId=SHEETS_ID, body={'requests': [request]}).execute()

#delete sheet
def delete_sheet(service,SHEETS_ID,sheet_name,sheet_ids):
    request_body = {
                    "requests": [
                                    {
                                      "deleteSheet": {
                                        "sheetId": sheet_ids[sheet_name]
                                                    }
                                    }
                                ]
                    }
    service.spreadsheets().batchUpdate(spreadsheetId=SHEETS_ID, body=request_body).execute()


def delete_row(service,SHEETS_ID,sheet_name,sheet_ids):
    request_body = {
                      "requests": [
                        {
                          "deleteDimension": {
                            "range": {
                              "sheetId": sheet_ids[sheet_name],
                              "dimension": "ROWS",
                              "startIndex": 0,
                              "endIndex": 34
                            }
                          }
                        },
                      ],
                    }
    service.spreadsheets().batchUpdate(spreadsheetId=SHEETS_ID, body=request_body).execute()



# To Get sheet Ids
def get_sheet_ids(SHEETS_ID, service):
    spreadsheet = service.spreadsheets().get(spreadsheetId=SHEETS_ID).execute()
    sheets_list = [{sheet['properties']['title']:sheet['properties']['sheetId']} for sheet in spreadsheet['sheets']]
    sheet_ids = {}
    for ids in sheets_list:
        sheet_ids.update(ids)
    return sheet_ids

def set_column_index(column_list):
    index = 0
    col_index = {}
    for title in column_list:
        col_index.update({title:index})
        index+=1
    return col_index

# define column range 
def col_range(col_index):
    if len(list(col_index.keys())) > len(letters):
        ll= 'A'
        ll+=letters[len(list(col_index.keys()))- len(letters)]
    else:
        ll = letters[len(col_index.keys())]
    return ll

def create_app_group_pivot(SHEETS_ID, sheet_name, col_index, sheet_ids,os_category='None'):
    request_body = {
    'requests': [
        {
            'updateCells': {
                'rows': {
                    'values': [
                        {
                            'pivotTable': {
                                # Data Source
                                'source': {
                                    'sheetId': sheet_ids['Filtered_data'], # Source 
                                    'startRowIndex': 0,#1
                                    'startColumnIndex': 0,
                                    # 'endRowIndex': 702,  # Need to define the 
                                    # 'endColumnIndex': len(list(col_index.keys())) # base index is 1
                                },
                                
                                # Rows Field(s)
                                'rows': [
                                    # row field #1
                                    {
                                        'sourceColumnOffset': col_index['Group'],
                                        'showTotals': False, # display subtotals
                                        'sortOrder': 'ASCENDING',
                                        'repeatHeadings': False
                                    },
                                     {
                                        'sourceColumnOffset': col_index['Title'],
                                        'showTotals': False, # display subtotals
                                        'sortOrder': 'ASCENDING',
                                        'repeatHeadings': False
                                    },
                                     {
                                        'sourceColumnOffset': col_index['IP'],
                                        'showTotals': False, # display subtotals
                                        'sortOrder': 'ASCENDING',
                                        'repeatHeadings': False
                                    },
                                     {
                                        'sourceColumnOffset': col_index['DNS'],
                                        'showTotals': False, # display subtotals
                                        'sortOrder': 'ASCENDING',
                                        'repeatHeadings': False
                                    }
                                ],

                                # Columns Field(s)
                                # 'columns': [
                                #     # column field #1
                                #     {
                                #         'sourceColumnOffset': 14,
                                #         'sortOrder': 'ASCENDING', 
                                #         'showTotals': True
                                #     }
                                # ],

                                # Values Field(s)
                                'values': [
                                    # value field #1
                                    {
                                        'sourceColumnOffset': 1,
                                        'summarizeFunction': 'COUNTA'
                                    }
                                ],

                                'criteria': {
                                            col_index['Category_filter']: {
                                                'visibleValues': ['Application','No Data']
                                                },
                                            col_index['Os_category']: {'visibleValues': os_category} if os_category!='None' else None
                                            },

                                'valueLayout': 'HORIZONTAL'
                            }
                        }
                    ]
                },
                
                'start': {
                    'sheetId': sheet_ids[sheet_name],
                    'rowIndex': 1, # 4th row
                    'columnIndex': 1 # 3rd column
                },
                'fields': 'pivotTable'
            }
            }
        ]
    }
    response = service.spreadsheets().batchUpdate(spreadsheetId=SHEETS_ID,body=request_body).execute()
    print('Application-Group Pivot Created')


def create_app_ip_pivot(SHEETS_ID, sheet_name, col_index, sheet_ids, os_category='None'):
    request_body = {
    'requests': [
        {
            'updateCells': {
                'rows': {
                    'values': [
                        {
                            'pivotTable': {
                                # Data Source
                                'source': {
                                    'sheetId': sheet_ids['Filtered_data'], # Source 
                                    'startRowIndex': 0,#1
                                    'startColumnIndex': 0,
                                    # 'endRowIndex': 702,  # Need to define the 
                                    # 'endColumnIndex': len(list(col_index.keys())) # base index is 1
                                },
                                
                                # Rows Field(s)
                                'rows': [
                                    # row field #1
                                    {
                                        'sourceColumnOffset': col_index['IP'],
                                        'showTotals': False, # display subtotals
                                        'sortOrder': 'ASCENDING',
                                        'repeatHeadings': False
                                    },
                                     {
                                        'sourceColumnOffset': col_index['Group'],
                                        'showTotals': False, # display subtotals
                                        'sortOrder': 'ASCENDING',
                                        'repeatHeadings': False
                                    },
                                     {
                                        'sourceColumnOffset': col_index['Title'],
                                        'showTotals': False, # display subtotals
                                        'sortOrder': 'ASCENDING',
                                        'repeatHeadings': False
                                    },
                                     {
                                        'sourceColumnOffset': col_index['DNS'],
                                        'showTotals': False, # display subtotals
                                        'sortOrder': 'ASCENDING',
                                        'repeatHeadings': False
                                    }
                                ],

                                # Columns Field(s)
                                # 'columns': [
                                #     # column field #1
                                #     {
                                #         'sourceColumnOffset': 14,
                                #         'sortOrder': 'ASCENDING', 
                                #         'showTotals': True
                                #     }
                                # ],

                                # Values Field(s)
                                'values': [
                                    # value field #1
                                    {
                                        'sourceColumnOffset': 1,
                                        'summarizeFunction': 'COUNTA'
                                    }
                                ],

                                'criteria': {
                                            col_index['Category_filter']: {
                                                'visibleValues': ['Application','No Data']
                                                },
                                            col_index['Os_category']: {'visibleValues': os_category} if os_category!='None' else None
                                            },

                                'valueLayout': 'HORIZONTAL'
                            }
                        }
                    ]
                },
                
                'start': {
                    'sheetId': sheet_ids[sheet_name],
                    'rowIndex': 1, # 4th row
                    'columnIndex': 1 # 3rd column
                },
                'fields': 'pivotTable'
            }
            }
        ]
    }
    response = service.spreadsheets().batchUpdate(spreadsheetId=SHEETS_ID,body=request_body).execute()
    print('Application-IP Pivot Created')


def create_app_age_pivot(SHEETS_ID, sheet_name, col_index, sheet_ids,os_category='None'):
    request_body = {
    'requests': [
        {
            'updateCells': {
                'rows': {
                    'values': [
                        {
                            'pivotTable': {
                                # Data Source
                                'source': {
                                    'sheetId': sheet_ids['Filtered_data'], # Source 
                                    'startRowIndex': 0,#1
                                    'startColumnIndex': 0,
                                    # 'endRowIndex': 702,  # Need to define the 
                                    # 'endColumnIndex': len(list(col_index.keys())) # base index is 1
                                },
                                
                                # Rows Field(s)
                                'rows': [
                                    # row field #1
                                    {
                                        'sourceColumnOffset': col_index['Aging_group'],
                                        'showTotals': False, # display subtotals
                                        'sortOrder': 'ASCENDING',
                                        'repeatHeadings': False
                                    },
                                    {
                                        'sourceColumnOffset': col_index['Group'],
                                        'showTotals': False, # display subtotals
                                        'sortOrder': 'ASCENDING',
                                        'repeatHeadings': False
                                    },
                                     {
                                        'sourceColumnOffset': col_index['Title'],
                                        'showTotals': False, # display subtotals
                                        'sortOrder': 'ASCENDING',
                                        'repeatHeadings': False
                                    },
                                     {
                                        'sourceColumnOffset': col_index['IP'],
                                        'showTotals': False, # display subtotals
                                        'sortOrder': 'ASCENDING',
                                        'repeatHeadings': False
                                    },
                                     {
                                        'sourceColumnOffset': col_index['DNS'],
                                        'showTotals': False, # display subtotals
                                        'sortOrder': 'ASCENDING',
                                        'repeatHeadings': False
                                    },
                                ],

                                # Columns Field(s)
                                # 'columns': [
                                #     # column field #1
                                #     {
                                #         'sourceColumnOffset': 14,
                                #         'sortOrder': 'ASCENDING', 
                                #         'showTotals': True
                                #     }
                                # ],

                                # Values Field(s)
                                'values': [
                                    # value field #1
                                    {
                                        'sourceColumnOffset': 1,
                                        'summarizeFunction': 'COUNTA'
                                    }
                                ],

                                'criteria': {
                                            col_index['Category_filter']: {
                                                'visibleValues': ['Application','No Data']
                                                },
                                            col_index['Os_category']: {'visibleValues': os_category} if os_category!='None' else None
                                            },

                                'valueLayout': 'HORIZONTAL'
                            }
                        }
                    ]
                },
                
                'start': {
                    'sheetId': sheet_ids[sheet_name],
                    'rowIndex': 1, # 2nd row
                    'columnIndex': 1 # 2dd column
                },
                'fields': 'pivotTable'
            }
            }
        ]
    }
    response = service.spreadsheets().batchUpdate(spreadsheetId=SHEETS_ID,body=request_body).execute()
    print('Application-Age Pivot Created')


def create_os_group_pivot(SHEETS_ID, sheet_name, col_index, sheet_ids, os_category='None'):
    request_body = {
    'requests': [
        {
            'updateCells': {
                'rows': {
                    'values': [
                        {
                            'pivotTable': {
                                # Data Source
                                'source': {
                                    'sheetId': sheet_ids['Filtered_data'], # Source 
                                    'startRowIndex': 0,#1
                                    'startColumnIndex': 0,
                                    # 'endRowIndex': 702,  # Need to define the 
                                    # 'endColumnIndex': len(list(col_index.keys())) # base index is 1
                                },
                                
                                # Rows Field(s)
                                'rows': [
                                    # row field #1
                                    {
                                        'sourceColumnOffset': col_index['Group'],
                                        'showTotals': False, # display subtotals
                                        'sortOrder': 'ASCENDING',
                                        'repeatHeadings': False
                                    },
                                     {
                                        'sourceColumnOffset': col_index['Title'],
                                        'showTotals': False, # display subtotals
                                        'sortOrder': 'ASCENDING',
                                        'repeatHeadings': False
                                    },
                                     {
                                        'sourceColumnOffset': col_index['IP'],
                                        'showTotals': False, # display subtotals
                                        'sortOrder': 'ASCENDING',
                                        'repeatHeadings': False
                                    },
                                     {
                                        'sourceColumnOffset': col_index['DNS'],
                                        'showTotals': False, # display subtotals
                                        'sortOrder': 'ASCENDING',
                                        'repeatHeadings': False
                                    }
                                ],

                                # Columns Field(s)
                                # 'columns': [
                                #     # column field #1
                                #     {
                                #         'sourceColumnOffset': 14,
                                #         'sortOrder': 'ASCENDING', 
                                #         'showTotals': True
                                #     }
                                # ],

                                # Values Field(s)
                                'values': [
                                    # value field #1
                                    {
                                        'sourceColumnOffset': 1,
                                        'summarizeFunction': 'COUNTA'
                                    }
                                ],



                                'criteria': {
                                            col_index['Category_filter']: {
                                                'visibleValues': ['OS']
                                                },
                                            col_index['Os_category']: {'visibleValues': os_category} if os_category!='None' else None
                                            },

                                'valueLayout': 'HORIZONTAL'
                            }
                        }
                    ]
                },
                
                'start': {
                    'sheetId': sheet_ids[sheet_name],
                    'rowIndex': 1, # 4th row
                    'columnIndex': 1 # 3rd column
                },
                'fields': 'pivotTable'
            }
            }
        ]
    }
    response = service.spreadsheets().batchUpdate(spreadsheetId=SHEETS_ID,body=request_body).execute()
    print('OS-Group Pivot Created')

def create_owners_group_pivot(SHEETS_ID, sheet_name, col_index, sheet_ids, os_category='None'):
    request_body = {
    'requests': [
        {
            'updateCells': {
                'rows': {
                    'values': [
                        {
                            'pivotTable': {
                                # Data Source
                                'source': {
                                    'sheetId': sheet_ids['Filtered_data'], # Source 
                                    'startRowIndex': 0,#1
                                    'startColumnIndex': 0,
                                    # 'endRowIndex': 702,  # Need to define the 
                                    # 'endColumnIndex': len(list(col_index.keys())) # base index is 1
                                },
                                
                                # Rows Field(s)
                                'rows': [
                                    # row field #1
                                    {
                                        'sourceColumnOffset': col_index['Owners'],
                                        'showTotals': False, # display subtotals
                                        'sortOrder': 'ASCENDING',
                                        'repeatHeadings': False
                                    },
                                    {
                                        'sourceColumnOffset': col_index['Group'],
                                        'showTotals': False, # display subtotals
                                        'sortOrder': 'ASCENDING',
                                        'repeatHeadings': False
                                    },
                                     {
                                        'sourceColumnOffset': col_index['Title'],
                                        'showTotals': False, # display subtotals
                                        'sortOrder': 'ASCENDING',
                                        'repeatHeadings': False
                                    },
                                     {
                                        'sourceColumnOffset': col_index['IP'],
                                        'showTotals': False, # display subtotals
                                        'sortOrder': 'ASCENDING',
                                        'repeatHeadings': False
                                    },
                                     {
                                        'sourceColumnOffset': col_index['DNS'],
                                        'showTotals': False, # display subtotals
                                        'sortOrder': 'ASCENDING',
                                        'repeatHeadings': False
                                    }
                                ],

                                # Columns Field(s)
                                # 'columns': [
                                #     # column field #1
                                #     {
                                #         'sourceColumnOffset': 14,
                                #         'sortOrder': 'ASCENDING', 
                                #         'showTotals': True
                                #     }
                                # ],

                                # Values Field(s)
                                'values': [
                                    # value field #1
                                    {
                                        'sourceColumnOffset': 1,
                                        'summarizeFunction': 'COUNTA'
                                    }
                                ],



                                'criteria': {
                                            col_index['Category_filter']: {
                                                'visibleValues': ['OS']
                                                },
                                            col_index['Os_category']: {'visibleValues': os_category} if os_category!='None' else None
                                            },

                                'valueLayout': 'HORIZONTAL'
                            }
                        }
                    ]
                },
                
                'start': {
                    'sheetId': sheet_ids[sheet_name],
                    'rowIndex': 1, # 4th row
                    'columnIndex': 1 # 3rd column
                },
                'fields': 'pivotTable'
            }
            }
        ]
    }
    response = service.spreadsheets().batchUpdate(spreadsheetId=SHEETS_ID,body=request_body).execute()
    print('Owner-Group Pivot Created')

def create_os_ip_pivot(SHEETS_ID, sheet_name, col_index, sheet_ids, os_category='None'):
    request_body = {
    'requests': [
        {
            'updateCells': {
                'rows': {
                    'values': [
                        {
                            'pivotTable': {
                                # Data Source
                                'source': {
                                    'sheetId': sheet_ids['Filtered_data'], # Source 
                                    'startRowIndex': 0,#1
                                    'startColumnIndex': 0,
                                    # 'endRowIndex': 702,  # Need to define the 
                                    # 'endColumnIndex': len(list(col_index.keys())) # base index is 1
                                },
                                
                                # Rows Field(s)
                                'rows': [
                                    # row field #1
                                    {
                                        'sourceColumnOffset': col_index['IP'],
                                        'showTotals': False, # display subtotals
                                        'sortOrder': 'ASCENDING',
                                        'repeatHeadings': False
                                    },
                                     {
                                        'sourceColumnOffset': col_index['Group'],
                                        'showTotals': False, # display subtotals
                                        'sortOrder': 'ASCENDING',
                                        'repeatHeadings': False
                                    },
                                     {
                                        'sourceColumnOffset': col_index['Title'],
                                        'showTotals': False, # display subtotals
                                        'sortOrder': 'ASCENDING',
                                        'repeatHeadings': False
                                    },
                                     {
                                        'sourceColumnOffset': col_index['DNS'],
                                        'showTotals': False, # display subtotals
                                        'sortOrder': 'ASCENDING',
                                        'repeatHeadings': False
                                    }
                                ],

                                # Columns Field(s)
                                # 'columns': [
                                #     # column field #1
                                #     {
                                #         'sourceColumnOffset': 14,
                                #         'sortOrder': 'ASCENDING', 
                                #         'showTotals': True
                                #     }
                                # ],

                                # Values Field(s)
                                'values': [
                                    # value field #1
                                    {
                                        'sourceColumnOffset': 1,
                                        'summarizeFunction': 'COUNTA'
                                    }
                                ],

                                'criteria': {
                                            col_index['Category_filter']: {
                                                'visibleValues': ['OS']
                                                },
                                            col_index['Os_category']: {'visibleValues': os_category} if os_category!='None' else None
                                            },

                                'valueLayout': 'HORIZONTAL'
                            }
                        }
                    ]
                },
                
                'start': {
                    'sheetId': sheet_ids[sheet_name],
                    'rowIndex': 1, # 4th row
                    'columnIndex': 1 # 3rd column
                },
                'fields': 'pivotTable'
            }
            }
        ]
    }
    response = service.spreadsheets().batchUpdate(spreadsheetId=SHEETS_ID,body=request_body).execute()
    print('OS-IP Pivot Created')

def create_os_age_pivot(SHEETS_ID, sheet_name, col_index, sheet_ids, os_category='None'):
    request_body = {
    'requests': [
        {
            'updateCells': {
                'rows': {
                    'values': [
                        {
                            'pivotTable': {
                                # Data Source
                                'source': {
                                    'sheetId': sheet_ids['Filtered_data'], # Source 
                                    'startRowIndex': 0,#1
                                    'startColumnIndex': 0,
                                    # 'endRowIndex': 702,  # Need to define the 
                                    # 'endColumnIndex': len(list(col_index.keys())) # base index is 1
                                },
                                
                                # Rows Field(s)
                                'rows': [
                                    # row field #1
                                    {
                                        'sourceColumnOffset': col_index['Aging_group'],
                                        'showTotals': False, # display subtotals
                                        'sortOrder': 'ASCENDING',
                                        'repeatHeadings': False
                                    },
                                    {
                                        'sourceColumnOffset': col_index['Group'],
                                        'showTotals': False, # display subtotals
                                        'sortOrder': 'ASCENDING',
                                        'repeatHeadings': False
                                    },
                                     {
                                        'sourceColumnOffset': col_index['Title'],
                                        'showTotals': False, # display subtotals
                                        'sortOrder': 'ASCENDING',
                                        'repeatHeadings': False
                                    },
                                     {
                                        'sourceColumnOffset': col_index['IP'],
                                        'showTotals': False, # display subtotals
                                        'sortOrder': 'ASCENDING',
                                        'repeatHeadings': False
                                    },
                                     {
                                        'sourceColumnOffset': col_index['DNS'],
                                        'showTotals': False, # display subtotals
                                        'sortOrder': 'ASCENDING',
                                        'repeatHeadings': False
                                    },
                                ],

                                # Columns Field(s)
                                # 'columns': [
                                #     # column field #1
                                #     {
                                #         'sourceColumnOffset': 14,
                                #         'sortOrder': 'ASCENDING', 
                                #         'showTotals': True
                                #     }
                                # ],

                                # Values Field(s)
                                'values': [
                                    # value field #1
                                    {
                                        'sourceColumnOffset': 1,
                                        'summarizeFunction': 'COUNTA'
                                    }
                                ],

                                'criteria': {
                                            col_index['Category_filter']: {
                                                'visibleValues': ['OS']
                                                },
                                            col_index['Os_category']: {'visibleValues': os_category} if os_category!='None' else None
                                            },

                                'valueLayout': 'HORIZONTAL'
                            }
                        }
                    ]
                },
                
                'start': {
                    'sheetId': sheet_ids[sheet_name],
                    'rowIndex': 1, # 2nd row
                    'columnIndex': 1 # 2dd column
                },
                'fields': 'pivotTable'
            }
            }
        ]
    }
    response = service.spreadsheets().batchUpdate(spreadsheetId=SHEETS_ID,body=request_body).execute()
    print('OS-Age Pivot Created')

def read_sheet_as_df(SHEETS_ID,sheet_index):
    sheet = gc.open_by_key(SHEETS_ID)
    worksheet = sheet.get_worksheet(sheet_index)
    all_values = worksheet.get_all_values()
    df = pd.DataFrame(all_values[1:], columns=all_values[0])
    # should be added the title [['AGE_Group','Vulnerability count']]
    return df

def convert_df_into_sheet_values(df):
    df = df.where(pd.notnull(df), None)
    value_list = df.values.tolist()
    values = [df.columns.tolist()]+value_list
    return values

def write_df_to_sheet(SHEETS_ID, sheet_name,df, location):
    sheet = gc.open_by_key(SHEETS_ID)
    worksheet = sheet.worksheet(sheet_name)
    # worksheet.set_dataframe(df)
    content =  convert_df_into_sheet_values(df)
    worksheet.update(location, content)
    print("successfully written to Google Sheet.")

def create_bar_chart(SHEETS_ID, sheet_name, col_index, sheet_ids,data_range):
    request_body = {
    "requests": [
    {
      "addChart": {
        "chart": {
          "spec": {
            "title": "Top 10 Vulnerability contributted IPs",
            "basicChart": {
              "chartType": "COLUMN",
              "legendPosition": "BOTTOM_LEGEND",
              "axis": [
                {
                  "position": "BOTTOM_AXIS",
                  "title": "Group"
                },
                {
                  "position": "LEFT_AXIS",
                  "title": "Count"
                }
              ],
              "domains": [
                {
                  "domain": {
                    "sourceRange": {
                      "sources": [
                        {
                          "sheetId": sheet_ids['Summary'],
                          "startRowIndex": data_range['row']-1,
                          "endRowIndex": data_range['row']+10,
                          "startColumnIndex": data_range['col']+3,
                      "endColumnIndex": data_range['col']+4
                        }
                      ]
                    }
                  }
                }
              ],
              "series": [
                {
                  "series": {
                    "sourceRange": {
                      "sources": [
                        {
                          "sheetId": sheet_ids['Summary'],
                          "startRowIndex": data_range['row']-1,
                          "endRowIndex": data_range['row']+10,
                          "startColumnIndex": data_range['col']+4,
                          "endColumnIndex": data_range['col']+5
                        }
                      ]
                    }
                  },
                  "targetAxis": "LEFT_AXIS"
                }
              ],
              "headerCount": 1
            }
          },
          "position": {
            "overlayPosition": {
              "anchorCell": {
                "sheetId": sheet_ids['Summary'],
                "rowIndex": data_range['row']-8,
                "columnIndex": 14
              },
              "offsetXPixels": 50,
              "offsetYPixels": 50
            }
          }
        }
      }
    }
    ]
    }
    response = service.spreadsheets().batchUpdate(spreadsheetId=SHEETS_ID,body=request_body).execute()
    print('Bar-Chart Created')

def create_pie_chart(SHEETS_ID, sheet_name, col_index, sheet_ids,data_range):
    request_body = {
  "requests": [
    {
      "addChart": {
        "chart": {
          "spec": {
            "title": "Top 10 Group's Vulnerability Counts",
            "pieChart": {
              "legendPosition": "RIGHT_LEGEND",
              "threeDimensional": True,
              "domain": {
                "sourceRange": {
                  "sources": [
                    {
                      "sheetId": sheet_ids['Summary'],
                      "startRowIndex": data_range['row'],
                      "endRowIndex": data_range['row']+10,
                      "startColumnIndex": data_range['col'],
                      "endColumnIndex": data_range['col']+1
                    }
                  ]
                }
              },
              "series": {
                "sourceRange": {
                  "sources": [
                    {
                      "sheetId": sheet_ids['Summary'],
                      "startRowIndex": data_range['row'],
                      "endRowIndex": data_range['row']+10,
                      "startColumnIndex": data_range['col']+1,
                      "endColumnIndex": data_range['col']+2
                    }
                  ]
                }
              },
            }
          },
          "position": {
            "overlayPosition": {
              "anchorCell": {
                "sheetId": sheet_ids['Summary'],
                "rowIndex": data_range['row']-8,
                "columnIndex": 8
              },
              "offsetXPixels": 60,
              "offsetYPixels": 50
            }
          }
        }
      }
    }
    ]
    }
    response = service.spreadsheets().batchUpdate(spreadsheetId=SHEETS_ID,body=request_body).execute()
    print('Pie-Chart Created')

#--------------------------------------------FORMATTING------------------------------------------#
def data_format(SHEETS_ID, loc):
    sheet = gc.open_by_key(SHEETS_ID)
    worksheet = sheet.worksheet('Summary')
    # worksheet.update('A1:B4', values)
    fmt = cellFormat(
        backgroundColor=color(0.8745098039, 0.9490196078, 0.9490196078),
        textFormat=textFormat(bold=False, 
        foregroundColor=color(0, 0, 0)),
        horizontalAlignment='CENTER'
    )
    format_cell_range(worksheet, loc, fmt)

def header_format(SHEETS_ID, loc):
    sheet = gc.open_by_key(SHEETS_ID)
    worksheet = sheet.worksheet('Summary')
    # worksheet.update('A1:B4', values)
    fmt = cellFormat(
        backgroundColor=color(0.4078431373, 0.5058823529, 0.768627451),
        textFormat=textFormat(bold=True, 
        foregroundColor=color(0, 0, 0)),
        horizontalAlignment='CENTER'
    )
    format_cell_range(worksheet, loc, fmt)
   
def set_boarder(SHEETS_ID, loc):
    range_to_format = loc
    borders = {
    'style': 'SOLID',
    'width': 1,
    'color': {
        'red': 0,
        'green': 0,
        'blue': 0
    }
    }
    sheet_ids = get_sheet_ids(SHEETS_ID,service)
    border_request = {
    'requests': [
        {
            'updateBorders': {
                'range': {
                    'sheetId': sheet_ids['Summary'],
                    'startRowIndex': int(range_to_format.split(':')[0][1:]),
                    'endRowIndex': int(range_to_format.split(':')[1][1:]),
                    'startColumnIndex': ord(range_to_format.split(':')[0][0]) - ord('A'),
                    'endColumnIndex': ord(range_to_format.split(':')[1][0]) - ord('A') + 1,
                },
                'top': borders,
                'bottom': borders,
                'left': borders,
                'right': borders
            }
        }
    ]
    }
    response = service.spreadsheets().batchUpdate(spreadsheetId=SHEETS_ID,body=border_request).execute()

#---------------------set of format---------------------
def data_formatting(SHEETS_ID, cell_range):
    header_format(SHEETS_ID,f"A{cell_range['start']}:B{cell_range['start']}")
    data_format(SHEETS_ID,f"A{cell_range['start']+1}:B{cell_range['end']}")
    set_boarder(SHEETS_ID,f"A{cell_range['start']}:B{cell_range['end']}")

    header_format(SHEETS_ID,f"D{cell_range['start']}:E{cell_range['start']}")
    data_format(SHEETS_ID,f"D{cell_range['start']+1}:E{cell_range['end']}")
    set_boarder(SHEETS_ID,f"D{cell_range['start']}:E{cell_range['end']}")

    header_format(SHEETS_ID,f"G{cell_range['start']}:H{cell_range['start']}")
    data_format(SHEETS_ID,f"G{cell_range['start']+1}:H{cell_range['start']+4}")
    set_boarder(SHEETS_ID,f"G{cell_range['start']}:H{cell_range['start']+4}")
#-------------------------------------------------------

#formating for categories
def category_format(SHEETS_ID, loc):
    sheet = gc.open_by_key(SHEETS_ID)
    worksheet = sheet.worksheet('Summary')
    # worksheet.update('A1:B4', values)
    fmt = cellFormat(
        backgroundColor=color(0.9607843137,0.8549019608,0.2588235294),
        textFormat=textFormat(bold=True, 
        foregroundColor=color(0, 0, 0)),
        horizontalAlignment='CENTER'
    )
    format_cell_range(worksheet, loc, fmt)

def write_category_title(SHEETS_ID, sheet_name,content, location):
    sheet = gc.open_by_key(SHEETS_ID)
    worksheet = sheet.worksheet(sheet_name)
    # worksheet.set_dataframe(df)
    worksheet.update(location, content)
