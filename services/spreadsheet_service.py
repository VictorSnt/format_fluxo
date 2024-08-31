from pathlib import Path
import pandas as pd

class SpreadsheetService:
    """
    A service class for handling spreadsheet operations with Excel files.
    
    Attributes:
        file_path (str): The path to the Excel file.
        data (pandas.DataFrame): The DataFrame containing the data from the Excel file.
    """

    def __init__(self, file_path: str|Path):
        """
        Initializes the SpreadsheetService with the path to the Excel file.
        
        Args:
            file_path (str): The path to the Excel file.
        """
        self.file_path = file_path
        self.data = self._load_data()

    def _load_data(self) -> pd.DataFrame:
        """
        Loads the data from the Excel file into a pandas DataFrame.
        
        Returns:
            pandas.DataFrame: The DataFrame containing the data from the Excel file.
        """
        return pd.read_excel(self.file_path)

    def save_data(self) -> None:
        """
        Saves the current DataFrame to the Excel file specified by file_path.
        
        The DataFrame is saved without the index column.
        """
        self.data.to_excel(f'outtput{self.file_path}', index=False)

    def remove_duplicates(self) -> None:
        """
        Removes all duplicate rows from the DataFrame.
        
        This method removes duplicates based on all columns, keeping only unique rows.
        """
        # Remove all duplicate rows
        self.data = self.data.drop_duplicates()
        # Optionally, reset index after dropping duplicates
        self.data.reset_index(drop=True, inplace=True)
    
    def get_column_numbers(self):
        """
        Retrieves all data from the 'Número' column in the DataFrame.
        
        Returns:
            pandas.Series: A Series containing all values from the 'Número' column.
        
        Raises:
            KeyError: If the 'Número' column does not exist in the DataFrame.
        """
        if 'Número' not in self.data.columns:
            raise KeyError("A coluna 'Número' não existe no DataFrame.")
        return self.data['Número']
    
    
    def update_idpessoa_from_tuples(self, updates: list[tuple]) -> None:
        """
        Updates the 'idpessoa' column in the DataFrame based on the provided list of tuples.
        
        Args:
            updates (list of tuples): Each tuple should contain (idpessoa, Número).
            
        Raises:
            KeyError: If the 'Número' column does not exist in the DataFrame.
        """
        if 'Número' not in self.data.columns:
            raise KeyError("A coluna 'Número' não existe no DataFrame.")
        
        if 'idpessoa' not in self.data.columns:
            # Add the 'idpessoa' column if it does not exist
            self.data['idpessoa'] = None
        
        # Create a DataFrame from the updates list
        updates_df = pd.DataFrame(updates, columns=['idpessoa', 'Número'])
        
        # Merge the updates DataFrame with the existing data
        self.data = self.data.merge(updates_df, on='Número', how='left', suffixes=('', '_update'))
        
        # Update the 'idpessoa' column with values from the updates DataFrame
        self.data['idpessoa'] = self.data['idpessoa_update'].combine_first(self.data['idpessoa'])
        
        # Drop the temporary update column
        self.data.drop(columns='idpessoa_update', inplace=True)
        
        # Save changes to the file
   
    
    
   
    
