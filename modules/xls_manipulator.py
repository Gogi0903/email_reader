import pandas as pd
import xlwings as wg 


class XlsProcessor:

    '''
    A class to handle processing and writing data to Excel files using xlwings.
    
    Attributes:
        file_path (str): The path to the Excel file.
    
    Methods:
        find_last_row(sheet: wg.Sheet, col: str = 'A', max_empty: int = 10) -> int:
            Finds the last non-empty row in the specified column of the Excel sheet.
        
        data_to_excel(sheet_name: str, data: pd.DataFrame) -> None:
            Writes data from a Pandas DataFrame to the specified sheet in the Excel file.
    '''
    
    def __init__(self, file_path):

        '''
        Initializes the XlsProcessor with the path to the Excel file.
        
        Parameters:
            file_path (str): The path to the Excel file.
        '''

        self.file_path = file_path
    
    @staticmethod
    def find_last_row(sheet: wg.Sheet, col='A', max_empty=10) -> int:

        '''
        Finds the last non-empty row in the specified column of the Excel sheet.
        
        Parameters:
            sheet (wg.Sheet): The Excel sheet to search.
            col (str): The column to search for the last non-empty row. Default is 'A'.
            max_empty (int): The number of consecutive empty rows to consider before stopping. Default is 10.
        
        Returns:
            int: The row number of the last non-empty row.
        '''

        last_row = 1
        empty_count = 0
        
        while True:
            if sheet.range(f'{col}{last_row}').value is None:
                empty_count += 1
                if empty_count >= max_empty:
                    break
            else:
                empty_count = 0
            last_row += 1

        return last_row - max_empty
    

    @staticmethod
    def reverse_date(df: pd.DataFrame) -> list:

        '''
        Converts the date's format from english to hungarian.

        Parameters:
            df (pd.DataFrame): pandas dataframe containing the date.

        Returns:
            list: list of converted dates.
        '''

        dates = list()
        for index, row in df.iterrows():
            date = row.iloc[3]
            splitted_date = date.split('.')

            reversed_date_list = [i for i in splitted_date[::-1]]        
            hun_formated_date = '.'.join(reversed_date_list)
            dates.append(hun_formated_date)

        return dates

    @staticmethod
    def df_modding(df: pd.DataFrame, date_list: list) -> pd.DataFrame:
        
        '''
        Exchanges the date's format to hungarian.

        Parameters:
            df (pd.DataFrame): pandas dataframe to change.
            date_list (list): this is the formated dates' list.

        Returns:
            pd.DataFrame: dataframe with the newly formated date.
        '''

        for i, date in enumerate(date_list):
            df.iloc[i, 3] = date
        return df

    def additional_datas(self, col_m: int=0, col_n=None, col_o=None, col_p=None, col_q=None, col_r: str='x', col_s: str='x') -> list:
        return [col_m, col_n, col_o, col_p, col_q, col_r, col_s]

    def data_to_excel(self, sheet_name: str, data: pd.DataFrame, add_datas: list) -> None:

        '''
        Writes data from a Pandas DataFrame to the specified sheet in the Excel file.
        
        Parameters:
            sheet_name (str): The name of the sheet to write the data to.
            data (pd.DataFrame): The DataFrame containing the data to be written.
        '''
        
        formated_date = XlsProcessor.reverse_date(df=data)
        msg_data = XlsProcessor.df_modding(
            df=data,
            date_list=formated_date
        )
        data_from_msg_list = msg_data.values.tolist()

        app = wg.App(visible=False)
        wb = app.books.open(self.file_path)
        sheet = wb.sheets(sheet_name)
        last_row = XlsProcessor.find_last_row(sheet=sheet)

        for row_data in data_from_msg_list:
            current_row = last_row + 1
            
            # datas from DF to excel
            for col_index, value in enumerate(row_data):
                sheet.range((current_row, col_index + 1)).value = value

            # adding col P to the additional list
            old_col_p = sheet.range(f'P{last_row}').formula
            col_p = old_col_p.replace(f'{last_row}', f'{current_row}')
            add_datas[3] = col_p

            # datas from 'all_datas' to excel
            for col_index, value in enumerate(add_datas):
                cell = sheet.range((current_row, col_index + 13))
                
                cell.value = value

            last_row += 1

        wb.save()
        wb.close()
        app.quit()
    