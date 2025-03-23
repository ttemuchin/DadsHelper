import sys
from PyQt5.QtWidgets import QApplication, QMessageBox, QInputDialog, QMainWindow, QVBoxLayout, QPushButton, QTableWidget, QTableWidgetItem, QFileDialog, QWidget
from PyQt5.QtCore import Qt
from bs4 import BeautifulSoup
import pyodbc

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("DadsHelper")
        self.setGeometry(600, 400, 400, 300)

        layout = QVBoxLayout()

        self.table = QTableWidget()
        layout.addWidget(self.table)

        self.button = QPushButton("Open HTML File")
        self.button.clicked.connect(self.open_file)
        layout.addWidget(self.button)

        container = QWidget()
        container.setLayout(layout)
        self.setCentralWidget(container)

    def open_file(self):
        file_name, _ = QFileDialog.getOpenFileName(self, "Open HTML File", "", "HTML Files (*.html *.htm)")
        if file_name:
            data = self.parse_html_type1(file_name)
            self.display_data(data)
            self.insert_into_access(data)

    def parse_html_type1(self, file_name):
        with open(file_name, 'r', encoding='utf-8') as file:
            soup = BeautifulSoup(file, 'html.parser')
        # div-table-tr-td-table-tr-td-font     
        data = []
        for div in soup.find_all('div'):
            batch = []
            tables = div.find_all('table')
            for table in tables:
                rows = table.find_all('tr')
                for row in rows:
                    cells = row.find_all('td')
                    for cell in cells:
                        innerTables = cell.find_all('table', border='1')
                        for inT in innerTables:
                            innerTrs = inT.find_all('tr')                         
                            for innerTr in innerTrs:
                                tds = innerTr.find_all('td')
                                for td in tds:
                                    fonts = td.find_all('font', size='2')
                                    for f in fonts:
                                        span = f.find('span')
                                        batch.append(span.text.strip())
            data.append({"sys": batch[0], "user": batch[1], "password": batch[2]})
        # print(f"Parsed data: {data}")
        return data

    def display_data(self, data):
        self.table.setRowCount(len(data))
        self.table.setColumnCount(3)
        self.table.setHorizontalHeaderLabels(["System", "User", "Password"])

        for i, row in enumerate(data):
            self.table.setItem(i, 0, QTableWidgetItem(row["sys"]))
            self.table.setItem(i, 1, QTableWidgetItem(row["user"]))
            self.table.setItem(i, 2, QTableWidgetItem(row["password"]))

    def get_db_path(self):
        db_path, _ = QFileDialog.getOpenFileName(
            self, "Open database file (*.accdb *.mdb)", "", "Access Files (*.accdb *.mdb)"
        )
        return db_path
    
    def get_db_password(self):
        password, ok = QInputDialog.getText(self, "DB password", "Password:")
        if ok:
            return password
        return None
    
    def connect_to_db(self, db_path, password=None):
        try:
            conn_str = rf'DRIVER={{Microsoft Access Driver (*.mdb, *.accdb)}};DBQ={db_path};'
            if password:
                conn_str += f'PWD={password};'
            conn = pyodbc.connect(conn_str)
            print("Successful connection to db!")
            return conn
        except Exception as e:
            print(f"Failed conn to db: {e}")
            return None
    
    def fetch_user_data(self, cursor, user_name):
        try:
            cursor.execute("SELECT * FROM UserResurses WHERE UserName = ?", (user_name,))
            row = cursor.fetchone()  # только первую найденную строку
            if row:
                print(f"Found data for user {user_name}: {row}")
                return row
            else:
                print(f"User {user_name} not found.")
                return None
        except Exception as e:
            print(f"Select ERROR: {e}")
            return None
        
    def insert_into_access(self, data):
        # таблица UserResurses => Система это ресурс. в поле Pwd вставляем значение пароля из html файла. 
        # путь к БД
        db_path = self.get_db_path()
        if not db_path:
            QMessageBox.warning(self, "Error", "Path to database not found")
            return
        # вход в БД
        while True:
            password = self.get_db_password()
            if password is None:
                # cancel operation
                return

            conn = self.connect_to_db(db_path, password)
            if conn:
                break
            else:
                QMessageBox.warning(self, "Error", "Wrong password, please try again")

        cursor = conn.cursor()

        NotFound = [] # user не найден в БД
        SysConflict = [] # система и ресурc не совпали

        for dict in data:
            user_name = dict["user"]
            # print(f"Данные для пользователя: {user_name}")
            user_data = self.fetch_user_data(cursor, user_name)
            if user_data:
                if dict["sys"] == user_data[1]:  # точно знаем, что Resurs — второе поле в строке
                    cursor.execute("UPDATE UserResurses SET Pwd = ? WHERE UserName = ?", (dict["password"], user_name))
                    print(f"Пароль для пользователя {user_name} обновлен.")
                else:
                    SysConflict.append(dict)
            else:
                NotFound.append(dict)
        conn.commit()

        conn.close()

        if not NotFound and not SysConflict:
            QMessageBox.information(self, "Successful operation", "All passwords updated")
        else:
            if NotFound:
                not_found_message = "NOT FOUND users below:\n"
                for item in NotFound:
                    not_found_message += f"{item['sys']} - {item['user']} - {item['password']}\n"
                QMessageBox.warning(self, "Attention!", not_found_message)
            if SysConflict:
                sys_conflict_message = "System/resource doesn't match for users below:\n"
                for item in SysConflict:
                    sys_conflict_message += f"{item['sys']} - {item['user']} - {item['password']}\n"
                QMessageBox.warning(self, "System conflict", sys_conflict_message)

if __name__ == "__main__":
    app = QApplication(sys.argv)
    w = MainWindow()
    w.show()
    sys.exit(app.exec())