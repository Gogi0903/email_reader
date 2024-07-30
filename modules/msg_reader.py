import pandas as pd
import os
import win32com.client as wc
from io import StringIO


class MsgReader:

    '''
    A class to handle reading, converting, and processing Outlook MSG files.
    
    Attributes:
        file_dir (str): The directory where MSG files are stored.
    
    Methods:
        list_of_files() -> list[str]:
            Returns a list of filenames in the directory.
        
        converting_msg_to_html(file_path: str, file_name: str) -> str:
            Converts the content of the specified Outlook MSG file to HTML format.
        
        converting_html_to_df(html: str) -> pd.DataFrame:
            Converts HTML content into a Pandas DataFrame.
        
        processing_dataframe(dataframe: pd.DataFrame) -> pd.DataFrame:
            Cleans and processes the DataFrame to remove unnecessary rows and columns.
    '''

    def __init__(self, file_dir: str):

        '''
        Initializes the MsgReader with the directory where MSG files are stored.
        
        Parameters:
            file_dir (str): The directory path containing the MSG files.
        '''

        self.file_dir = file_dir

    def list_of_files(self) -> list[str]:

        '''
        Returns a list of filenames in the directory.
        
        Returns:
            list[str]: A list containing the names of all files in the specified directory.
        '''

        return os.listdir(
            path=self.file_dir
            )

    def converting_msg_to_html(self, file_path: str, file_name: str) -> str:

        '''
        Converts the content of the Outlook MSG file into HTML format.
        
        Parameters:
            file_path (str): The path to the directory containing the MSG file.
            file_name (str): The name of the MSG file to convert.
        
        Returns:
            str: The HTML content of the MSG file.
        '''

        # opening outlook
        outlook = wc.Dispatch('Outlook.Application').GetNamespace('MAPI')
        # reading msg
        msg = outlook.OpenSharedItem(f'{file_path}/{file_name}')
        # converting msg to html
        html = msg.HTMLBody
        
        return html

    def converting_html_to_df(self, html: str) -> pd.DataFrame:

        '''
        Converts the content of the HTML into a Pandas DataFrame.
        
        Parameters:
            html (str): The HTML content to be converted.
        
        Returns:
            pd.DataFrame: A DataFrame containing the data parsed from the HTML content.
        '''
        html_io = StringIO(html)

        return pd.read_html(html_io)[0]

    def processing_dataframe(self, dataframe: pd.DataFrame) -> pd.DataFrame:

        ''' 
        Transforms the input DataFrame into a clean format.
        
        Parameters:
            dataframe (pd.DataFrame): The DataFrame to be processed.
        
        Returns:
            pd.DataFrame: A cleaned and processed DataFrame with unnecessary rows and columns removed.
        '''
        
        # setting first row's values to dataframe's columns' names 
        dataframe.columns = dataframe.iloc[0]

        # deleting the first row, since it isn't needed anymore
        dataframe = dataframe.drop(
            axis=0, 
            index=[0, 1]
            )

        # deleting the last two columns, since they aren't needed (NaN values)
        dataframe = dataframe.drop(
            [dataframe.columns[0], dataframe.columns[-2], dataframe.columns[11]], 
            axis=1
            )

        # deleting blank lines
        dataframe = dataframe.dropna()

        return dataframe
