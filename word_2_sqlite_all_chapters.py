import docx
import sqlite3

def read_word_table(file_path):
    doc = docx.Document(file_path)
    data = []

    for table in doc.tables:
        for row in table.rows:
            row_data = [cell.text.strip() for cell in row.cells]
            data.append(row_data)

    return data

def create_sqlite_db(db_file):
    conn = sqlite3.connect(db_file)
    cursor = conn.cursor()

    # Create a table to store the data
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS YourTable (
            Level TEXT,
            Chapter1 TEXT,
            Korean TEXT,
            English TEXT
        )
    ''')

    conn.commit()
    conn.close()

def insert_into_sqlite(db_file, data):
    conn = sqlite3.connect(db_file)
    cursor = conn.cursor()

    # Insert data into the YourTable table
    for row in data:
        cursor.execute('''
            INSERT INTO YourTable (Level, Chapter1, Korean, English)
            VALUES (?, ?, ?, ?)
        ''', (row[0], row[1], row[2], row[3]))

    conn.commit()
    conn.close()

if __name__ == "__main__":
    word_file_path = "D:/Level2_vocabulary.docx"
    sqlite_db_file = "D:/Level2_vocabulary_database.db"

    # Read data from Word table
    table_data = read_word_table(word_file_path)

    # Create SQLite database and table
    create_sqlite_db(sqlite_db_file)

    # Insert data into SQLite database
    insert_into_sqlite(sqlite_db_file, table_data)

    print("Data has been successfully imported into SQLite.")
