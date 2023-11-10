import requests
import pandas as pd
from pathlib import Path

class Metabase:
    output_folder: Path = None
    username: str = None
    password: str = None
    domain: str = None
    
    def __init__(self):
        """Interact with metabase
        
        Params:
            username (str): Credential to login
            password (str): Credential to login
            domain (str): url of metabase app
        """
        
        ## Authentication
        self.session = requests.post(
            self.domain + '/api/session',
            json = {
                'username': self.username,
                'password': self.password
            }
        )
        self.session_id = self.session.json()['id']
        self.session_header = {'X-Metabase-Session': self.session_id}
        
    ## Ex.: Employee is db. 36 and tb. 516
    def get_table(self, database: int = None, table: int = None, format: str = 'json'):
        """Gets the data of a table from metabase

        Args:
            database (int): id of the database
            table (int): id of the table
            format (str): output format ['json']

        Returns:
            Dataframe: Data of the table
        """
        
        query = f'query=%7B%22database%22%3A{database}%2C%22query%22%3A%7B%22source-table%22%3A{table}%7D%2C%22type%22%3A%22query%22%7D'    # Visualization settings and middleware removed
        response = requests.post(f'{self.domain}/api/dataset/{format}?{query}', headers = self.session_header)
        
        match format:
            case 'json':
                data = pd.DataFrame(response.json())
        
        return data