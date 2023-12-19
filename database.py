import os
import sqlite3


class Database:
    """
    Class to manage and interact with the sqlite database

    Attributes:

    - db_name :    :class:`str` --> The name of the database file

    Methods:

    - _exist_chk() --> This modified version of exist_chk provided in the starting materials checks to see if the database file exists. If the file does not exist, it will be created. The insurance table will then be created, regardless of if the file existed prior to the program running.
    - _create_database() --> Create the file to use for the database.
    - _create_table() --> This modified version of create_table provided in the starting materials. This method connects to the database and creates a table if one does not exist already.
    - write_data() --> Take in data and write the data to the sqlite database
    - read_data() --> Read in data from the sqlite database and return it
    - read_unique_data() --> Read and return the unique values from the field indicated
    - read_filtered_data() --> Read and return all rows containing the same value for the field provided
    """

    def __init__(self):
        self.db_name = "insurance_data.db"
        self._exist_chk()

    def _exist_chk(self):
        """
        This modified version of exist_chk provided in the starting materials checks to see if the database file exists.
        If the file does not exist, it will be created. The insurance table will then be created, regardless of if the
        file existed prior to the program running.
        :return None:
        """
        exist_chk = os.path.exists(os.path.join(os.getcwd(), self.db_name))

        if not exist_chk:
            self._create_database()

        self._create_table()

    def _create_database(self):
        """
        Create the file to use for the database.
        :return None:
        """
        with open(self.db_name, 'w') as fp:
            print('Database has been created.')

    def _create_table(self):
        """
        This modified version of create_table provided in the starting materials. This method connects to the database and creates a table if one does not exist already.

        :return bool:
        """
        try:
            conn = sqlite3.connect(self.db_name)
            cursor = conn.cursor()
            # Make sure the policy # is unique
            query = '''CREATE TABLE IF NOT EXISTS insurance
                            (insurance_id INTEGER PRIMARY KEY,
                            policy INTEGER NOT NULL UNIQUE,
                            expiry TEXT NOT NULL,
                            location TEXT NOT NULL,
                            state TEXT NOT NULL,
                            region TEXT NOT NULL,
                            insurance_value INTEGER DEFAULT 0,
                            construction TEXT NOT NULL,
                            business_type TEXT NOT NULL,
                            earthquake INTEGER DEFAULT 0,
                            flood INTEGER DEFAULT 0);'''

            cursor.execute(query)
            conn.commit()
            print('The insurance table has been created')

        except sqlite3.Error as error:
            print(f'Error ocured - {error}')

        finally:
            if conn:
                conn.close()
                return True

    def write_data(self, data):
        """
        Take in data and write the data to the sqlite database
        :param data: The data from Excel to write to the database
        :return None:
        """
        conn = sqlite3.connect(self.db_name)
        cursor = conn.cursor()
        # Insert data into the "insurance" table
        insert_query = '''
            INSERT OR IGNORE INTO insurance (policy, expiry, location, state, region, insurance_value, construction, 
            business_type, earthquake, flood) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        '''

        # Iterate through rows in the Excel sheet and insert each row into the table
        for row in data.iter_rows(min_row=2, values_only=True):
            cursor.execute(insert_query, row)

        # Commit the changes and close the connection
        conn.commit()
        conn.close()

    def read_data(self):
        """
        Read in data from the sqlite database and return it
        :return data: A list of tuples containing data from the sqlite database
        """
        conn = sqlite3.connect(self.db_name)
        cursor = conn.cursor()
        select_query = "SELECT * FROM insurance"
        cursor.execute(select_query)
        data = cursor.fetchall()

        conn.close()
        return data

    def read_unique_data(self, field):
        """
        Read and return the unique values from the field indicated
        :param field: The database field to get unique values from
        :return data: The unique values of the field
        """
        conn = sqlite3.connect(self.db_name)
        cursor = conn.cursor()

        # Query the database for unique values in the selected field
        cursor.execute(f"SELECT DISTINCT {field} FROM insurance")
        data = [row[0] for row in cursor.fetchall()]

        # Close the connection
        conn.close()

        return data

    def read_filtered_data(self, field, value):
        """
        Read and return all rows containing the same value for the field provided
        :param field: The field to select
        :param value: The value to find duplicates of
        :return data: The rows containing the same value for the field
        """
        conn = sqlite3.connect(self.db_name)
        cursor = conn.cursor()
        if field == "*":
            cursor.execute("SELECT * FROM insurance")
        else:
            cursor.execute(f"SELECT * FROM insurance WHERE {field} = ?", (value,))

        # Fetch all the matching entries
        data = cursor.fetchall()

        # Close the connection
        conn.close()
        return data
