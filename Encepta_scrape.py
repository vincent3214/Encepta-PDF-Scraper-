import pandas as pd
import PyPDF2
import re
from openpyxl import load_workbook

def load_file():

    file=input('Enter the file:')
    file_name=input('Enter xlsx file name to save:')
    sheet_name=input('sheet_name:')
    try:
        input1 = PyPDF2.PdfFileReader(open(file, "rb"))
        # Try open pdf file (use full path)
    except:
        raise ValueError()

    Comments = [] # list for comments
    nPages = input1.getNumPages()
    # get number of pages and parse through for comments
    for i in range(nPages):
        page = input1.getPage(i)
        try:
            for annot in page['/Annots']:
                Comments.append(annot.getObject())
                # append comment to comment list
        except:
            pass
                # if no comment on page pass

    Comment_df = pd.DataFrame(Comments) #change list to dataframe
    Comment_df['Error'] = 'N/a'
    Comment_df['Cat_Error']='None'
    User_com = Comment_df[['/Contents', '/Subtype', '/T', 'Error','Cat_Error']]

    df = User_com[User_com['/Contents'].notna()] # make sure we only get notna values
    index_content = df.columns.get_loc('/Contents')
    index_error = df.columns.get_loc('Error')
    index_cat_error = df.columns.get_loc('Cat_Error')
    error_pattern = r'@\w+'
    for row in range(0, len(df)):
        try:
            Comment = re.search(error_pattern, df.iat[row, index_content]).group()
        except:
            Comment = 'N/a'
        df.iat[row, index_error] = Comment
        df.iat[row, index_cat_error] = Comment
    df.rename(columns={"/Contents": "Annotations", "/Subtype": "Annot_Type", '/T': 'User'}, inplace=True)
    cat_map = {'N/a': '', '@1': 'Typo', '@2': "Missing Count (draft/design)", '@3': 'Wrong SNET Count',
               '@4': 'Design Error', '@5': 'More'}

    df['Cat_Error'].map((cat_map))
    df['Cat_Error'] = df['Cat_Error'].map((cat_map))
    df = df.loc[df['User'] != 'AutoCAD SHX Text']
    try:
        book = load_workbook(f'{file_name}.xlsx')
        writer = pd.ExcelWriter(f'{file_name}.xlsx', engine='openpyxl')
        writer.book = book
        writer.sheets = dict((ws.title, ws) for ws in book.worksheets)
        df.to_excel(writer,
                    f"{sheet_name}",
                    na_rep='None',
                    startcol=3,
                    startrow=3)
        writer.save()

    except:
        xlwriterDF = pd.ExcelWriter(f'{file_name}.xlsx')
        df.to_excel(
                excel_writer=xlwriterDF,
                sheet_name=f'{sheet_name}',
                na_rep='None',
                startcol=3,
                startrow=3)
        xlwriterDF.save()


    print('File analysed')

load_file()
