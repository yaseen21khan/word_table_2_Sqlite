import docx
import sqlite3

def read_from_sqlite(db_file, table_name):
    conn = sqlite3.connect(db_file)
    cursor = conn.cursor()

    # Fetch data from the SQLite table
    cursor.execute(f'SELECT * FROM {table_name}')
    data = cursor.fetchall()

    conn.close()

    return data

def create_word_table(doc, data):
    # Create a table in the Word document
    table = doc.add_table(rows=1, cols=len(data[0]))

    # Add headers to the table
    for col, header in enumerate(data[0]):
        table.cell(0, col).text = header

    # Add data to the table
    for row in data[1:]:
        new_row = table.add_row()
        for col, value in enumerate(row):
            new_row.cells[col].text = str(value)

def main():
    word_file_path = "E:/Chapter1.docx"
    sqlite_db_file = "D:/database.db"
    table_name = "Chapter1"

    # Read data from SQLite
    sqlite_data = read_from_sqlite(sqlite_db_file, table_name)

    # Create a Word document
    doc = docx.Document()

    # Create a Word table and populate it with data from SQLite
    create_word_table(doc, sqlite_data)
    print(sqlite_data)
    # Save the Word document
    doc.save(word_file_path)

    print("Data has been successfully exported to Word.")

if __name__ == "__main__":
    main()
