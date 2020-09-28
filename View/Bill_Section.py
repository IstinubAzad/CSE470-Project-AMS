from tkinter import *
import Controller.Bill_Database as BD_M


class Bill_Section:
    def __init__(self, root):

        self.root = root

        Label(root, text="        ", bg="#CAE9F5", fg='black').grid(row=0, column=0)
        Label(root, text="        ", bg="#CAE9F5", fg='black').grid(row=1, column=0)
        Label(root, text="        ", bg="#CAE9F5", fg='black').grid(row=2, column=0)

        # ------------------Title----------------------
        Label(root, text="BILL SECTION", bg="#CAE9F5", fg='Black', font=("Helvetica", 15, 'bold')).place(x=60, y=10)
        Label(root, text="Click flat number & bill titles for detailed information & money voucher:", bg="#CAE9F5",
              fg='Black',
              font=("Helvetica", 10,)).place(x=0, y=50)

        refresh_button = Button(root, text="Refresh", bg="#ADECDF")
        refresh_button.place(x=700, y=140)

        Home_button = Button(root, text="Home Page", bg="#ADECDF")
        Home_button.place(x=295, y=20)

        quit_button = Button(root, text="Quit", bg="#ADECDF", command=root.destroy)
        quit_button.place(x=260, y=20)

        generate_monthly_bill = Button(root, text="Generate monthly bill each", bg="#ADECDF",
                                       command=self.monthly_values)
        generate_monthly_bill.place(x=700, y=100)

        delete_bill_row_button = Button(root, text="Clear values of row :", bg="#ADECDF", command=self.delete_bill_row)
        delete_bill_row_button.place(x=565, y=20)
        self.delete_bill_row_value = Entry(width=5)
        self.delete_bill_row_value.place(x=690, y=23)

        for i in range(self.total_rows):
            for j in range(self.total_columns):
                if i == 0:
                    if j == 3:
                        self.btn = Button(root, text=str(self.bill_data_list[i][j]) + "   ", bg="#ADECDF",
                                          font=("Helvetica", 9, 'bold'), pady=4, command=self.goto_electric_bill)
                        self.btn.grid(row=i + 3, column=j, pady=10)
                    elif j == 4:
                        self.btn = Button(root, text=str(self.bill_data_list[i][j]) + "   ", bg="#ADECDF",
                                          font=("Helvetica", 9, 'bold'), pady=4, command=self.goto_gas_bill)
                        self.btn.grid(row=i + 3, column=j, pady=10)
                    elif j == 5:
                        self.btn = Button(root, text=str(self.bill_data_list[i][j]) + "   ", bg="#ADECDF",
                                          font=("Helvetica", 9, 'bold'), pady=4, command=self.goto_water_bill)
                        self.btn.grid(row=i + 3, column=j, pady=10)
                    else:
                        self.e = Label(root, text=str(self.bill_data_list[i][j]) + "   ", bg="#CAE9F5", fg='black',
                                       font=("Helvetica", 10, 'bold'))
                        self.e.grid(row=i + 3, column=j)
                else:
                    if j == 1:
                        self.btn = Button(root, text=self.bill_data_list[i][j], width=5)
                        self.btn.grid(row=i + 3, column=j, pady=10)
                    else:
                        self.e = Label(root, text=self.bill_data_list[i][j], fg='black').grid(row=i + 3, column=j,
                                                                                              pady=10)

                if i == self.total_rows - 1:
                    if j == 2:
                        Label(root, text='TOTAL = ', fg='black', font=("Helvetica", 9, 'bold'), bg="#CAE9F5",
                              pady=20).grid(row=i + 5, column=j)
                    elif j == 3:
                        Label(root, text=str(BD_M.Bill_Calculation().electric_Bill()), bg="#CAE9F5", fg='black',
                              font=("Helvetica", 9, 'bold'), pady=20).grid(row=i + 5, column=j)
                    elif j == 4:
                        Label(root, text=str(BD_M.Bill_Calculation().gas_Bill()), bg="#CAE9F5", fg='black',
                              font=("Helvetica", 9, 'bold'),
                              pady=20).grid(row=i + 5, column=j)
                    elif j == 5:
                        Label(root, text=str(BD_M.Bill_Calculation().water_bill()), bg="#CAE9F5", fg='black',
                              font=("Helvetica", 9, 'bold'), pady=20).grid(row=i + 5, column=j)
                    elif j == 6:
                        Label(root, text=str(BD_M.Bill_Calculation().total_monthly_bill()), bg="#CAE9F5", fg='black',
                              font=("Helvetica", 9, 'bold'), pady=20).grid(row=i + 5, column=j)

    @staticmethod
    def monthly_values():
        BD_M.Bill_Calculation().monthly_calculation()
        BD_M.Bill_database_create().save_excel()

    def goto_water_bill(self):
        self.root.destroy()
        present_waterBill_gui_frame()

    def goto_gas_bill(self):
        self.root.destroy()
        present_gasBill_gui_frame()

    def goto_electric_bill(self):
        self.root.destroy()
        present_electricBill_gui_frame()

    def delete_bill_row(self):
        row = int(self.delete_bill_row_value.get())
        BD_M.Bill_database_edit().make_it_zero_row(row)
        BD_M.Bill_database_create().save_excel()

    bill_data_list = BD_M.Bill_database_read().load_workbook_values()
    total_rows = len(bill_data_list)
    total_columns = len(bill_data_list[0])


class Electric_Bill:
    def __init__(self, root):
        self.root = root
        Label(root, text="        ", bg="#CAE9F5", fg='black').grid(row=0, column=0)
        Label(root, text="        ", bg="#CAE9F5", fg='black').grid(row=1, column=0)
        Label(root, text="        ", bg="#CAE9F5", fg='black').grid(row=2, column=0)

        # ------------------Title----------------------
        Label(root, text="Electric Bill", bg="#CAE9F5", fg='Black', font=("Helvetica", 15, 'bold')).place(x=60, y=10)

        refresh_button = Button(root, text="Refresh", bg="#ADECDF")
        refresh_button.place(x=360, y=140)

        Back_button = Button(root, text="Back", bg="#ADECDF", command=self.go_back)
        Back_button.place(x=255, y=20)

        Home_button = Button(root, text="Home Page", bg="#ADECDF")
        Home_button.place(x=295, y=20)

        quit_button = Button(root, text="Quit", bg="#ADECDF", command=root.destroy)
        quit_button.place(x=360, y=200)

        delete_bill_row_button = Button(root, text="Delete row :", bg="#ADECDF", command=self.delete_bill_row)
        delete_bill_row_button.place(x=400, y=20)
        self.delete_bill_row_value = Entry(width=5)
        self.delete_bill_row_value.place(x=480, y=23)

        for i in range(self.total_rows):
            for j in range(self.total_columns):
                if j == 0 or j == 1 or j == 3:
                    if i == 0:
                        self.e = Label(root, text=str(self.elecricbill_data_list[i][j]) + "   ", bg="#CAE9F5",
                                       fg='black', font=("Helvetica", 10, 'bold'))
                        self.e.grid(row=i + 3, column=j)
                    else:
                        if j == 1:
                            self.btn = Button(root, text=self.elecricbill_data_list[i][j], width=5)
                            self.btn.grid(row=i + 3, column=j, pady=10)
                        else:
                            self.e = Label(root, text=self.elecricbill_data_list[i][j], fg='black').grid(row=i + 3,
                                                                                                         column=j,
                                                                                                         pady=10)
                if i == self.total_rows - 1:
                    if j == 1:
                        Label(root, text='TOTAL = ', fg='black', font=("Helvetica", 9, 'bold'), bg="#CAE9F5",
                              pady=20).grid(row=i + 5, column=j)

                    elif j == 5:
                        Label(root, text=str(BD_M.Bill_Calculation().electric_Bill()), bg="#CAE9F5", fg='black',
                              font=("Helvetica", 9, 'bold'), pady=20).grid(row=i + 5, column=3)

    def go_back(self):
        self.root.destroy()
        present_BillSection_gui_frame()

    def delete_bill_row(self):
        row = int(self.delete_bill_row_value.get())
        BD_M.Bill_database_edit().delete_electric_bill_row(row)
        BD_M.Bill_database_create().save_excel()

    elecricbill_data_list = BD_M.Bill_database_read().load_workbook_values()
    total_rows = len(elecricbill_data_list)
    total_columns = len(elecricbill_data_list[0])


class Water_bill:
    def __init__(self, root):
        self.root = root
        Label(root, text="        ", bg="#CAE9F5", fg='black').grid(row=0, column=0)
        Label(root, text="        ", bg="#CAE9F5", fg='black').grid(row=1, column=0)
        Label(root, text="        ", bg="#CAE9F5", fg='black').grid(row=2, column=0)

        # ------------------Title----------------------
        Label(root, text="Water Bill", bg="#CAE9F5", fg='Black', font=("Helvetica", 15, 'bold')).place(x=60, y=10)

        refresh_button = Button(root, text="Refresh", bg="#ADECDF")
        refresh_button.place(x=360, y=140)

        Back_button = Button(root, text="Back", bg="#ADECDF", command=self.go_back)
        Back_button.place(x=255, y=20)

        Home_button = Button(root, text="Home Page", bg="#ADECDF", command=self.goto_home_page)
        Home_button.place(x=295, y=20)

        quit_button = Button(root, text="Quit", bg="#ADECDF", command=root.destroy)
        quit_button.place(x=360, y=200)

        delete_bill_row_button = Button(root, text="Clear row :", bg="#ADECDF", command=self.delete_bill_row)
        delete_bill_row_button.place(x=400, y=20)
        self.delete_bill_row_value = Entry(width=5)
        self.delete_bill_row_value.place(x=480, y=23)

        for i in range(self.total_rows):
            for j in range(self.total_columns):
                if j == 0 or j == 1 or j == 5:
                    if i == 0:
                        self.e = Label(root, text=str(self.waterbill_data_list[i][j]) + "   ", bg="#CAE9F5", fg='black',
                                       font=("Helvetica", 10, 'bold'))
                        self.e.grid(row=i + 3, column=j)
                    else:
                        if j == 1:
                            self.btn = Button(root, text=self.waterbill_data_list[i][j], width=5)
                            self.btn.grid(row=i + 3, column=j, pady=10)
                        else:
                            self.e = Label(root, text=self.waterbill_data_list[i][j], fg='black').grid(row=i + 3,
                                                                                                       column=j,
                                                                                                       pady=10)

                if i == self.total_rows - 1:
                    if j == 1:
                        Label(root, text='TOTAL = ', fg='black', font=("Helvetica", 9, 'bold'), bg="#CAE9F5",
                              pady=20).grid(row=i + 5, column=j)
                    elif j == 3:
                        Label(root, text=str(BD_M.Bill_Calculation().water_bill()), bg="#CAE9F5", fg='black',
                              font=("Helvetica", 9, 'bold'), pady=20).grid(row=i + 5, column=5)

    def go_back(self):
        self.root.destroy()
        present_BillSection_gui_frame()

    def delete_bill_row(self):
        row = int(self.delete_bill_row_value.get())
        BD_M.Bill_database_edit().delete_water_bill(row)
        BD_M.Bill_database_create().save_excel()

    def goto_home_page(self):
        self.root.destroy()

    waterbill_data_list = BD_M.Bill_database_read().load_workbook_values()
    total_rows = len(waterbill_data_list)
    total_columns = len(waterbill_data_list[0])


class Gas_Bill:
    def __init__(self, root):
        self.root = root
        Label(root, text="        ", bg="#CAE9F5", fg='black').grid(row=0, column=0)
        Label(root, text="        ", bg="#CAE9F5", fg='black').grid(row=1, column=0)
        Label(root, text="        ", bg="#CAE9F5", fg='black').grid(row=2, column=0)

        # ------------------Title----------------------
        Label(root, text="Gas Bill", bg="#CAE9F5", fg='Black', font=("Helvetica", 15, 'bold')).place(x=60, y=10)

        refresh_button = Button(root, text="Refresh", bg="#ADECDF")
        refresh_button.place(x=360, y=140)

        Back_button = Button(root, text="Back", bg="#ADECDF", command=self.go_back)
        Back_button.place(x=255, y=20)

        Home_button = Button(root, text="Home Page", bg="#ADECDF", command=self.goto_home_page)
        Home_button.place(x=295, y=20)

        quit_button = Button(root, text="Quit", bg="#ADECDF", command=root.destroy)
        quit_button.place(x=360, y=200)

        delete_bill_row_button = Button(root, text="Delete row :", bg="#ADECDF", command=self.delete_bill_row)
        delete_bill_row_button.place(x=400, y=20)
        self.delete_bill_row_value = Entry(width=5)
        self.delete_bill_row_value.place(x=480, y=23)

        for i in range(self.total_rows):
            for j in range(self.total_columns):
                if j == 0 or j == 1 or j == 4:
                    if i == 0:
                        self.e = Label(root, text=str(self.gasbill_data_list[i][j]) + "   ", bg="#CAE9F5", fg='black',
                                       font=("Helvetica", 10, 'bold'))
                        self.e.grid(row=i + 3, column=j)
                    else:
                        if j == 1:
                            self.btn = Button(root, text=self.gasbill_data_list[i][j], width=5)
                            self.btn.grid(row=i + 3, column=j, pady=10)
                        else:
                            self.e = Label(root, text=self.gasbill_data_list[i][j], fg='black').grid(row=i + 3,
                                                                                                     column=j, pady=10)

                if i == self.total_rows - 1:
                    if j == 1:
                        Label(root, text='TOTAL = ', fg='black', bg="#CAE9F5", font=("Helvetica", 9, 'bold'),
                              pady=20).grid(row=i + 5, column=j)
                    elif j == 4:
                        Label(root, text=str(BD_M.Bill_Calculation().gas_Bill()), bg="#CAE9F5", fg='black',
                              font=("Helvetica", 9, 'bold'), pady=20).grid(row=i + 5, column=4)

    def go_back(self):
        self.root.destroy()
        present_BillSection_gui_frame()

    def delete_bill_row(self):
        row = int(self.delete_bill_row_value.get())
        BD_M.Bill_database_edit().delete_gas_bill_row(row)
        BD_M.Bill_database_create().save_excel()

    def goto_home_page(self):
        self.root.destroy()

    gasbill_data_list = BD_M.Bill_database_read().load_workbook_values()
    total_rows = len(gasbill_data_list)
    total_columns = len(gasbill_data_list[0])


def present_BillSection_gui_frame():
    root = Tk()
    Bill_Section(root)

    windowWidth = root.winfo_reqwidth()
    windowHeight = root.winfo_reqheight()

    positionRight = int(root.winfo_screenwidth() / 3 - windowWidth / 2)
    positionDown = int(root.winfo_screenheight() / 3 - windowHeight / 2)

    root.title('Apartment Management System')
    root.geometry("950x500+{}+{}".format(positionRight, positionDown))
    root['background'] = '#CAE9F5'
    root.mainloop()


def present_waterBill_gui_frame():
    root = Tk()
    Water_bill(root)

    windowWidth = root.winfo_reqwidth()
    windowHeight = root.winfo_reqheight()

    positionRight = int(root.winfo_screenwidth() / 3 - windowWidth / 2)
    positionDown = int(root.winfo_screenheight() / 3 - windowHeight / 2)

    root.title('Apartment Management System')
    root.geometry("550x500+{}+{}".format(positionRight, positionDown))
    root['background'] = '#CAE9F5'
    root.mainloop()


def present_gasBill_gui_frame():
    root = Tk()
    Gas_Bill(root)

    windowWidth = root.winfo_reqwidth()
    windowHeight = root.winfo_reqheight()

    positionRight = int(root.winfo_screenwidth() / 3 - windowWidth / 2)
    positionDown = int(root.winfo_screenheight() / 3 - windowHeight / 2)

    root.title('Apartment Management System')
    root.geometry("550x500+{}+{}".format(positionRight, positionDown))
    root['background'] = '#CAE9F5'
    root.mainloop()


def present_electricBill_gui_frame():
    root = Tk()
    Electric_Bill(root)

    windowWidth = root.winfo_reqwidth()
    windowHeight = root.winfo_reqheight()

    positionRight = int(root.winfo_screenwidth() / 3 - windowWidth / 2)
    positionDown = int(root.winfo_screenheight() / 3 - windowHeight / 2)

    root.title('Apartment Management System')
    root.geometry("550x500+{}+{}".format(positionRight, positionDown))
    root['background'] = '#CAE9F5'
    root.mainloop()


present_BillSection_gui_frame()
