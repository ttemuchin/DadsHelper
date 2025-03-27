import sys
from PyQt5.QtWidgets import QApplication, QMessageBox, QInputDialog, QLineEdit, QMainWindow, QVBoxLayout, QPushButton, QTableWidget, QTableWidgetItem, QFileDialog, QWidget
from PyQt5.QtCore import Qt
from bs4 import BeautifulSoup
import pyodbc, datetime

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("DadsHelper")
        self.setGeometry(600, 400, 600, 300)

        layout = QVBoxLayout()

        self.table = QTableWidget()
        layout.addWidget(self.table)

        self.buttonOpen = QPushButton("Update passwords with HTML")
        self.buttonOpen.clicked.connect(self.open_file)
        layout.addWidget(self.buttonOpen)

        self.buttonQuit = QPushButton("Quit")
        self.buttonQuit.clicked.connect(sys.exit)
        layout.addWidget(self.buttonQuit)

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
        # div-table-tr-td-table-tr-td-p-font-b-span     
        data = []
        for div in soup.find_all('div'):
            batch = []
            tables = div.find_all('table')
            for table in tables:
                rows = table.find_all('tr')
                for row in rows:
                    cells = row.find_all('td')
                    for cell in cells:
                        innerTables = cell.find_all('table')
                        for inT in innerTables:
                            innerTrs = inT.find_all('tr')                         
                            for innerTr in innerTrs:
                                tds = innerTr.find_all('td')
                                for td in tds:
                                    ps = td.find_all('p')
                                    for p in ps:
                                        batch.append(p.text.strip())

                                    fonts = td.find_all('font', size='2')
                                    for f in fonts:
                                        span = f.find('span')
                                        batch.append(span.text.strip())
            data.append({"name": batch[2], "sys": batch[3], "user": batch[4], "password": batch[5]})
        # print(f"Parsed data: {data}")
        return data

    def display_data(self, data):
        self.table.setRowCount(len(data))
        self.table.setColumnCount(3)
        self.table.setHorizontalHeaderLabels(["User", "Name", "System"])

        for i, row in enumerate(data):
            self.table.setItem(i, 0, QTableWidgetItem(row["user"]))
            self.table.setItem(i, 1, QTableWidgetItem(row["name"]))
            self.table.setItem(i, 2, QTableWidgetItem(row["sys"]))

    def get_db_path(self):
        db_path, _ = QFileDialog.getOpenFileName(
            self, "Open database file (*.accdb *.mdb)", "", "Access Files (*.accdb *.mdb)"
        )
        return db_path
    
    def get_db_password(self):
        password, ok = QInputDialog.getText(self, "DB password", "Password:", echo=QLineEdit.Password)
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
        
    def get_user_FIO(self, cursor, kod):
        try:
            # cursor.execute(f"SELECT TOP 1 * FROM Users")
            # columns = cursor.description       
            # print(f"Структура таблицы Users:", columns[0][0], columns[0][1])
            # Структура таблицы Users: kod <class 'int'>
            cursor.execute("SELECT * FROM Users WHERE kod = ?", (kod,))
            row = cursor.fetchone()
            print(row)
            if row:
                return [row[1], row[2], row[3]]
            else:
                print(f"User with 'Номер' {kod} not found in Users.")
                return None
        except Exception as e:
            print(f"Select ERROR: {e}")
            return None
        
    def compare_fio(self, db_fio, html_fio):
        html_fio = html_fio.split()
        for i in range(3):
            if html_fio[i].lower() != db_fio[i].lower():
                return False
        return True
    
    def write_to_log(self, message):
        with open('log-file.txt', 'a', encoding='utf-8') as log_file:
            log_file.write(message + '\n')
        
        
    def insert_into_access(self, data):
        # таблица UserResurses => в поле Pwd вставляем значение пароля из html файла. 
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

        SuccessUpdates = []
        UsernameNotFound = [] # username не найден в userResurses
        NameConflict = [] # fio не совпали
        UserIsNotInUsers = [] # пользователь не найден по ID в таблице Users

        for user in data:
            user_name = user["user"]
            print(f"Данные для пользователя: {user_name}")
            user_data = self.fetch_user_data(cursor, user_name)
            
            if user_data:
                kod = user_data[0]
                fio_DB = self.get_user_FIO(cursor, int(kod))
                if fio_DB:
                    if self.compare_fio(fio_DB, user["name"]):  # сравниваем ФИО в БД и в HTML
                        newDatetime = datetime.datetime.now() #08.05.2024 11:19:31
                        cursor.execute("UPDATE UserResurses SET Pwd = ? WHERE UserName = ?", (user["password"], user_name))
                        cursor.execute("UPDATE UserResurses SET datepar = ? WHERE UserName = ?", (newDatetime, user_name))
                        SuccessUpdates.append(user)
                        print(f"Пароль для пользователя {user_name} обновлен.")
                    else:
                        NameConflict.append(user)
                else:
                    UserIsNotInUsers.append(user)
            else:
                UsernameNotFound.append(user)
        conn.commit()

        conn.close()

        log_message = "PASSWORD UPDATE REPORT:\n(fullname - username - system)\n"
        open('log-file.txt', 'w', encoding='utf-8').close()
        success_message = "SUCCESSFUL UPDATE FOR:\n\n"
        for item in SuccessUpdates:
            success_message += f"{item['name']} - {item['user']} - {item['sys']}\n"
        log_message += success_message

        if not UsernameNotFound and not NameConflict and not UserIsNotInUsers:
            QMessageBox.information(self, "Successful operation", "All passwords updated")
        else:    
            if UsernameNotFound:
                not_found_message = "\nUSERNAME NOT FOUND IN 'USERRESURSES':\n\n"
                for item in UsernameNotFound:
                    not_found_message += f"{item['name']} - {item['user']} - {item['sys']}\n"
                log_message += not_found_message
            if NameConflict:
                name_conflict_message = "\nNAME IN HTML DIDN'T MATCH WITH NAME IN 'USERS':\n\n"
                for item in NameConflict:
                    name_conflict_message += f"{item['name']} - {item['user']} - {item['sys']}\n"
                log_message += name_conflict_message
            if UserIsNotInUsers:
                uiniu_message = "\nCOULDN'T FIND USER BY ID IN 'USERS':\n\n"
                for item in uiniu_message:
                    uiniu_message += f"{item['name']} - {item['user']} - {item['sys']}\n"
                log_message += uiniu_message
            QMessageBox.information(self, "Successful operation", "Passwords updated with some exceptions.\nDetailed report saved to log-file.txt")
        
        self.write_to_log(log_message)



if __name__ == "__main__":
    app = QApplication(sys.argv)
    w = MainWindow()
    w.show()
    sys.exit(app.exec())