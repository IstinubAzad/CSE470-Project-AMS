from tkinter import *
import Controller.Flat_Database as FD_M
import Controller.Bill_Database as BD_M


class Add_Owner:
    def __init__(self, master):
        self.master = master

        Label(master, text="                             ", bg="#CAE9F5", font=("Helvetica", 10)).grid(row=0, column=0)
        Label(master, text="Add flat owner info-", bg="#CAE9F5", font=("Helvetica", 10, 'bold')).place(x=0, y=0)

        Home_button = Button(master, text="Home Page", bg="#ADECDF", command=self.goto_home_page)
        Home_button.grid(row=10, column=0, sticky=E)

        quit_button = Button(master, text="Quit", bg="#ADECDF", command=master.destroy)
        quit_button.grid(row=10, column=2, sticky=E)

        l1 = Label(master, text="Flat number :", bg="#CAE9F5")
        l2 = Label(master, text="Owner name :", bg="#CAE9F5")
        l3 = Label(master, text="Contact Number :", bg="#CAE9F5")
        l4 = Label(master, text="Intercom no. :", bg="#CAE9F5")
        l5 = Label(master, text="Parking :", bg="#CAE9F5")

        l1.grid(row=1, column=0, sticky=W, pady=2)
        l2.grid(row=2, column=0, sticky=W, pady=2)
        l3.grid(row=3, column=0, sticky=W, pady=2)
        l4.grid(row=4, column=0, sticky=W, pady=2)
        l5.grid(row=5, column=0, sticky=W, pady=2)

        self.e1 = Entry()
        self.e2 = Entry()
        self.e3 = Entry()
        self.e4 = Entry()
        self.e5 = Entry()

        self.e1.grid(row=1, column=1, pady=2)
        self.e2.grid(row=2, column=1, pady=2)
        self.e3.grid(row=3, column=1, pady=2)
        self.e4.grid(row=4, column=1, pady=2)
        self.e5.grid(row=5, column=1, pady=2)

        b1 = Button(master, text="Add", bg="#ADECDF", command=self.add_value)
        b1.grid(row=10, column=1, sticky=E)

    def add_value(self):
        data_list = []
        data_list2 = []
        data_list.append(self.e1.get())
        data_list2.append(self.e1.get())
        data_list.append(self.e2.get())
        data_list2.append(self.e2.get())
        data_list.append(self.e3.get())
        data_list.append(self.e4.get())
        data_list.append(self.e5.get())
        FD_M.Flat_database_edit().add_information(data_list)
        FD_M.Flat_database_create().save_excel()

        BD_M.Bill_database_edit().add_information(data_list2)
        BD_M.Bill_database_create().save_excel()

    def goto_home_page(self):
        self.master.destroy()


def present_AddOwner_gui_frame():
    master = Tk()
    Add_Owner(master)

    windowWidth = master.winfo_reqwidth()
    windowHeight = master.winfo_reqheight()

    positionRight = int(master.winfo_screenwidth() / 2 - windowWidth / 2)
    positionDown = int(master.winfo_screenheight() / 2 - windowHeight / 2)

    master.title("Apartment Management System")
    master.geometry("400x200+{}+{}".format(positionRight, positionDown))
    master['background'] = '#CAE9F5'
    mainloop()


present_AddOwner_gui_frame()
