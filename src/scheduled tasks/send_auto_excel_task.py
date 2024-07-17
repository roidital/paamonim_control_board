from src.common.common_utils import connect_to_db
from mysql.connector import Error


def fetch_all_users_details_from_db():
    connection = connect_to_db()
    if connection:
        try:
            cursor = connection.cursor()
            fetch_query = "SELECT username, password, unit_name FROM users_details"
            cursor.execute(fetch_query)
            for (username, password, unit_name) in cursor:
                print(f"Username: {username}, Password: {password}, Unit Name: {unit_name}")
        except Error as e:
            print(f"Failed to fetch data: {e}")
        finally:
            cursor.close()
            connection.close()