from pathlib import Path
from typing import List
from database.db_connect import DBConector
from database.query import Query
from services.spreadsheet_service import SpreadsheetService


if __name__ == '__main__':
    connector = DBConector()
    if not connector.get_connection():
        raise Exception('Need to connect to the database')
    
    file_path = Path('./contas.xlsx')
    if not file_path.exists():
        raise Exception('Sheet path has no file')
        
    sheet_service = SpreadsheetService(file_path)
    numbers: List[str] = sheet_service.get_column_numbers().to_list()
    str_numbers = [num for num in numbers if isinstance(num, str)]
    query = Query("""
          SELECT idpessoa, nrtitulo FROM wshop.fluxo
          WHERE nrtitulo IN ({}) AND dtbaixa IS NULL AND dtexclusao IS NULL
    """.format(','.join([f"'{n}'" for n in str_numbers])))
    print(query.text)
    result = connector.execute_query(query)
    
    sheet_service.update_idpessoa_from_tuples(result)
    sheet_service.remove_duplicates()
    sheet_service.save_data()
    
    
   

