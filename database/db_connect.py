from os import getenv
from typing import Any, Optional
from dotenv import load_dotenv
import psycopg2

from database.errors import NotConnectedException
from database.query import Query
from database.types import Cursor, Connection


load_dotenv()
class DBConector:
    """"
    Handle all database operations
    """
    conn: Connection
    cursor: Cursor
    def get_connection(self) -> bool:
        """
        Connect to the database
        """
        self.conn: Connection = psycopg2.connect(
            host=getenv('HOST'),
            port=getenv('PORT'),
            dbname=getenv('DBNAME'),
            user=getenv('USER'),
            password=getenv('PASSWD')
        )

        self.cursor = self.conn.cursor()
        return True

    def execute_query(self, query: Query, params: Optional[tuple] = None):
        """
        Execute a query with parms if given
        """
        if not isinstance(self.cursor, Cursor):
            raise NotConnectedException('You must connect to databse first')

        self.cursor.execute(query.text, (params,))

        return self.cursor.fetchall()

    def destroy(self):
        """
        close connection and cursor
        """

        self.cursor.close()
        self.conn.close()
