from tkinter import *
import Controller.Flat_Database as FD_M
import Controller.Bill_Database as BD_M


class FlatProfile:
    def __init__(self, root):
        self.root = root
        Label(root, text="        ", bg="#CAE9F5", fg='black').grid(row=0, column=0)
        Label(root, text="        ", bg="#CAE9F5", fg='black').grid(row=1, column=0)
        Label(root, text="        ", bg="#CAE9F5", fg='black').grid(row=2, column=0)
        Label(root, text="        ", bg="#CAE9F5", fg='black').grid(row=3, column=0)

        # ------------------Title----------------------
        Label(root, text="FLAT PROFILE", bg="#CAE9F5", fg='Black', font=("Helvetica", 15, 'bold')).place(x=200, y=15)
        Label(root, text="Click flat number for details of specific flat information:", bg="#CAE9F5", fg='Black',
              font=("Helvetica", 10)).place(x=0, y=60)

        refresh_button = Button(root, text="Refresh", bg="#ADECDF")
        refresh_button.place(x=10, y=20)

        delete_flat_row_button = Button(root, text="Delete row :", bg="#ADECDF", command=self.delete_bill_row)
        delete_flat_row_button.place(x=610, y=20)
        self.delete_flat_row_value = Entry(width=5)
        self.delete_flat_row_value.place(x=690, y=23)

        home_button = Button(root, text="Home Page", bg="#ADECDF", command=self.goto_home_page)
        home_button.place(x=400, y=20)

        quit_button = Button(root, text="Quit", bg="#ADECDF", command=root.destroy)
        quit_button.place(x=490, y=20)

        for i in range(self.total_rows):
            for j in range(self.total_columns):
                if i == 0:
                    self.e = Label(root, text=str(self.flat_data_list[i][j]) + "    ", bg="#CAE9F5", fg='black',
                                   font=("Helvetica", 10, 'bold'))
                    self.e.grid(row=i + 4, column=j)
                else:
                    if j == 1:
                        self.btn = Button(root, text=self.flat_data_list[i][j], width=5)
                        self.btn.grid(row=i + 5, column=j, pady=10)
                    else:
                        self.e = Label(root, text=self.flat_data_list[i][j], fg='black').grid(row=i + 5, column=j)

    def goto_home_page(self):
        self.root.destroy()

    def delete_bill_row(self):
        FD_M.Flat_database_edit().delete_information_row(int(self.delete_flat_row_value.get()))
        FD_M.Flat_database_create().save_excel()

        BD_M.Bill_database_edit().delete_information_row(int(self.delete_flat_row_value.get()))
        BD_M.Bill_database_create().save_excel()

    flat_data_list = FD_M.Flat_database_read().load_workbook_values()
    total_rows = len(flat_data_list)
    total_columns = len(flat_data_list[0])


def present_FlatProfile_gui_frame():
    root = Tk()
    FlatProfile(root)

    windowWidth = root.winfo_reqwidth()
    windowHeight = root.winfo_reqheight()

    positionRight = int(root.winfo_screenwidth() / 3 - windowWidth / 2)
    positionDown = int(root.winfo_screenheight() / 3 - windowHeight / 2)

    root.title('Apartment Management System')
    root.geometry("860x350+{}+{}".format(positionRight, positionDown))
    root['background'] = '#CAE9F5'
    root.mainloop()


present_FlatProfile_gui_frame()
