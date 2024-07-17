import mysql.connector
from mysql.connector import Error


def connect_to_db():
    try:
        connection = mysql.connector.connect(
            host='roidital.mysql.pythonanywhere-services.com',
            database='roidital$default',
            user='roidital',
            password='mypass11'
        )
        if connection.is_connected():
            return connection
    except Error as e:
        print(f"Error while connecting to MySQL: {e}")
        return None


def create_table():
    connection = connect_to_db()
    if connection:
        try:
            cursor = connection.cursor()
            create_table_query = """
            CREATE TABLE IF NOT EXISTS users_details (
                username VARCHAR(255) NOT NULL PRIMARY KEY,
                password VARCHAR(255) NOT NULL,
                unit_name VARCHAR(255) NOT NULL
            );
            """
            cursor.execute(create_table_query)
            connection.commit()
        except Error as e:
            print(f"Failed to create table: {e}")
        finally:
            cursor.close()
            connection.close()


def insert_user_strings(user_data_list):
    connection = connect_to_db()
    if connection:
        try:
            cursor = connection.cursor()
            insert_query = """
            INSERT INTO users_details (username, password, unit_name)
            VALUES (%s, %s, %s)
            """

            cursor.execute(insert_query, tuple(user_data_list))
            connection.commit()
        except Error as e:
            print(f"Failed to insert data: {e}")
        finally:
            cursor.close()
            connection.close()


def register_user_to_receive_auto_excel(username, password, unit_name):
    user_data = [username, password, unit_name]
    create_table()  # Create the table if it doesn't exist
    insert_user_strings(user_data)