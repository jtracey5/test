import tkinter as tk
from tkinter import ttk
import sqlite3
import xlsxwriter
from xlsxwriter.workbook import Workbook


class Main(tk.Frame):
    def __init__(self, root):
        super().__init__(root)
        self.init_main()
        self.db = db
        self.view_records()


    def init_main(self):
        toolbar =tk.Frame(bg='#FFFFFF', bd=2)
        toolbar.pack(side=tk.TOP, fill=tk.X)

        self.add_img = tk.PhotoImage(file='add.gif')
        btn_open_dialog = tk.Button(toolbar, text='Добавить', command=self.open_dialog, bg='#FFFFFF', bd=0,
                                    compound=tk.TOP, image=self.add_img)
        btn_open_dialog.pack(side=tk.LEFT)

        self.update_img = tk.PhotoImage(file='update.gif')
        btn_edit_dialog = tk.Button(toolbar, text ='Редактировать', bg='#FFFFFF', bd=0, image=self.update_img,
                                    compound = tk.TOP, command=self.open_update_dialog)
        btn_edit_dialog.pack(side=tk.LEFT)

        self.delete_img = tk.PhotoImage(file='delete.gif')
        btn_edit_dialog = tk.Button(toolbar, text='Удалить', bg='#FFFFFF', bd=0, image=self.delete_img,
                                    compound=tk.TOP, command=self.delete_records)
        btn_edit_dialog.pack(side=tk.LEFT)

        self.show_img = tk.PhotoImage(file='users.gif')
        btn_edit_dialog = tk.Button(toolbar, text='Наши клиенты', bg='#FFFFFF', bd=0, image=self.show_img,
                                    compound=tk.TOP, command=self.show_users)
        btn_edit_dialog.pack(side=tk.LEFT)

        self.refresh_img = tk.PhotoImage(file='refresh.gif')
        btn_edit_dialog = tk.Button(toolbar, text='Обновить', bg='#FFFFFF', bd=0,
                                    image=self.refresh_img,
                                    compound=tk.TOP, command=self.view_records)
        btn_edit_dialog.pack(side=tk.LEFT)

        self.export_img = tk.PhotoImage(file='export.gif')
        btn_edit_dialog = tk.Button(toolbar, text='Экспорт', bg='#FFFFFF', bd=0,
                                    image=self.export_img,
                                    compound=tk.TOP, command=self.export_data)
        btn_edit_dialog.pack(side=tk.LEFT)


        self.tree=ttk.Treeview(self, column=("id", "fio", "login", "date_insert", "deleted"),
                               height=500, show='headings')

        self.tree.column("id", width=150, anchor=tk.CENTER)
        self.tree.column("fio", width=150, anchor=tk.CENTER)
        self.tree.column("login", width=150, anchor=tk.CENTER)
        self.tree.column("date_insert", width=150, anchor=tk.CENTER)
        self.tree.column("deleted", width=150, anchor=tk.CENTER)

        self.tree.heading("id", text='ID')
        self.tree.heading("fio", text='ФИО')
        self.tree.heading("login", text='Логин')
        self.tree.heading("date_insert", text='Дата добавления')
        self.tree.heading("deleted", text='Удален')
        self.tree.pack()

    def records(self, fio, login, date_insert, deleted):
        self.db.insert_data(fio, login, date_insert, deleted)
        self.view_records()

    def view_records(self):
        self.db.c.execute('''SELECT * FROM user''')
        [self.tree.delete(i) for i in self.tree.get_children()]
        [self.tree.insert('', 'end', values=row) for row in self.db.c.fetchall()]

    def delete_records(self):
        for selection_item in self.tree.selection():
            self.db.c.execute('''UPDATE user SET deleted = 'Да' where ID =?''', ( self.tree.set(selection_item, '#1'),))
        self.db.conn.commit()
        self.view_records()

    def update_records(self, fio, login, date_insert, deleted):
        self.db.c.execute('''UPDATE user SET fio=?, login=?, date_insert=?, deleted=? WHERE ID=?''',
                          (fio, login, date_insert, deleted, self.tree.set(self.tree.selection()[0], '#1')))
        self.db.conn.commit()
        self.view_records()

    def export_data(self):
        workbook = Workbook('alfa_test.xlsx')
        worksheet = workbook.add_worksheet()
        conn = sqlite3.connect('task.db')
        c = conn.cursor()
        mysel = c.execute("select * from user")
        for i, row in enumerate(mysel):
            for j, value in enumerate(row):
                worksheet.write(i, j, value)
        workbook.close()


    def show_users(self):
        self.db.c.execute('''SELECT * FROM user WHERE deleted = 'Нет' ''')
        [self.tree.delete(i) for i in self.tree.get_children()]
        [self.tree.insert('', 'end', values=row) for row in self.db.c.fetchall()]

    def open_dialog(self):
        Child()

    def open_update_dialog(self):
        Update()

class Child(tk.Toplevel):
    def __init__(self):
        super().__init__(root)
        self.init_child()
        self.view = app

    def init_child(self):
        self.title("Добавить пользователя")
        self.geometry('500x400+400+300')
        self.resizable(False, False)

        label_fio = tk.Label(self, text ='ФИО')
        label_fio.place(x=50, y=50)
        label_login = tk.Label(self, text='Логин')
        label_login.place(x=50, y=80)
        label_date_insrt = tk.Label(self, text='Дата добавления')
        label_date_insrt.place(x=50, y=110)
        label_del = tk.Label(self, text='Удален')
        label_del.place(x=50, y=140)

        self.entry_fio=ttk.Entry(self)
        self.entry_fio.place(x=200, y=50)

        self.entry_login= ttk.Entry(self)
        self.entry_login.place(x=200, y=80)

        self.entry_dates = ttk.Entry(self)
        self.entry_dates.place(x=200, y=110)

        self.entry_deleted = ttk.Combobox(self, values=[u"Да", u"Нет"])
        self.entry_deleted.current(1)
        self.entry_deleted.place(x=200, y=140)


        btn_cancel = ttk.Button(self, text='Закрыть', command=self.destroy)
        btn_cancel.place(x=400, y=350)

        self.btn_ok = ttk.Button(self, text='Добавить')
        self.btn_ok.place(x=200, y=200)
        self.btn_ok.bind('<Button-1>', lambda event: self.view.records(self.entry_fio.get(),
                                                                  self.entry_login.get(),
                                                                  self.entry_dates.get(),
                                                                  self.entry_deleted.get()))


        self.grab_set()
        self.focus_set()

class Update(Child):
    def __init__(self):
        super().__init__()
        self.init_edit()
        self.view = app

    def init_edit(self):
        self.title('Редактировать данные')
        btn_edit = ttk.Button(self, text='Редактировать')
        btn_edit.place(x=200, y=350)
        btn_edit.bind('<Button-1>', lambda event: self.view.update_records(self.entry_fio.get(),
                                                                  self.entry_login.get(),
                                                                  self.entry_dates.get(),
                                                                  self.entry_deleted.get()))
        self.btn_ok.destroy()




class DB:
    def __init__(self):
        self.conn = sqlite3.connect('task.db')
        self.c = self.conn.cursor()
        self.c.execute(
            '''CREATE TABLE IF NOT EXISTS user(id integer primary key, fio text, login text, date_insert text, 
            deleted text)''')
        self.conn.commit()




    def insert_data(self, fio,login, date_insert, deleted):
        self.c.execute('''INSERT INTO user (fio, login, date_insert, deleted) VALUES (?, ?, ?, ?)''',
                       (fio, login, date_insert, deleted))
        self.conn.commit()


if __name__ == "__main__":
    root = tk.Tk()
    db = DB()
    app = Main(root)
    app.pack()
    root.title("Alfa Test")
    root.geometry("900x600+300+200")
    root.resizable(False, False)
    root.mainloop()
