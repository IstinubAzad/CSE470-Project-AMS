from tkinter import *
import Controller.Flat_Database as FD_M
import Controller.Bill_Database as BD_M


class Homepage:
    def __init__(self, window):
        self.window = window
        # ------------------Title----------------------
        Title = Label(window, text="HOME PAGE", fg='Black', bg="#CAE9F5", font=("Helvetica", 40, 'bold'))
        Title.place(x=220, y=10)

        note = Label(window,
                     text="Add titles for both new files 'Flat_Database' & 'Bill_database'. Click the above right buttons to create titles in the database",
                     fg='Black', bg="#CAE9F5", font=("Helvetica", 10))
        note.place(x=10, y=400)

        create_Flat_data_button = Button(window, text="Create title for flat profile", bg="#ADECDF",
                                         command=self.createFlatDatabase)
        create_Flat_data_button.place(x=400, y=200)

        create_bill_data_button = Button(window, text="Create title for bill section", bg="#ADECDF",
                                         command=self.createBilldatabase)
        create_bill_data_button.place(x=400, y=240)

        quit_button = Button(window, text="Quit", bg="#ADECDF", command=window.destroy)
        quit_button.place(x=400, y=300)

        # -----------------Sections---------------------
        lbl = Label(window, text="Select a database to visit in the list below:", bg="#CAE9F5", fg='Black',
                    font=("Helvetica", 14, 'bold'))
        lbl.place(x=10, y=100)

        flat_Profile_btn = Button(window, text="Flat Profile", bg="#ADECDF", fg='blue', command=self.goto_flatProfile)
        flat_Profile_btn.place(x=20, y=140)

        Bill_Section_btn = Button(window, text="Bill Section", bg="#ADECDF", fg='blue', command=self.goto_billSection)
        Bill_Section_btn.place(x=20, y=180)

        Add_bill_btn = Button(window, text="Add Bill", bg="#ADECDF", fg='blue', command=self.goto_AddBill)
        Add_bill_btn.place(x=20, y=220)

        Add_flat_members_btn = Button(window, text="Add flat members", bg="#ADECDF", fg='blue',
                                      command=self.goto_AddMember)
        Add_flat_members_btn.place(x=20, y=260)

        Service_Charge_btn = Button(window, text="Service Charge", bg="#ADECDF", fg='blue')
        Service_Charge_btn.place(x=20, y=300)

        Maintenance_btn = Button(window, text="Maintenance", bg="#ADECDF", fg='blue')
        Maintenance_btn.place(x=20, y=340)

    @staticmethod
    def createFlatDatabase():
        FD_M.Flat_database_create().add_title()
        FD_M.Flat_database_create().save_excel()

    @staticmethod
    def createBilldatabase():
        BD_M.Bill_database_create().add_title()
        BD_M.Bill_database_create().save_excel()

    def goto_flatProfile(self):
        self.window.destroy()

    def goto_billSection(self):
        self.window.destroy()

    def goto_AddMember(self):
        self.window.destroy()

    def goto_AddBill(self):
        self.window.destroy()


def present_Homepage_gui_frame():
    window = Tk()
    Homepage(window)
    windowWidth = window.winfo_reqwidth()
    windowHeight = window.winfo_reqheight()

    positionRight = int(window.winfo_screenwidth() / 3 - windowWidth / 2)
    positionDown = int(window.winfo_screenheight() / 3 - windowHeight / 2)

    window.title('Apartment Management System')
    window.geometry("840x450+{}+{}".format(positionRight, positionDown))
    window['background'] = '#CAE9F5'
    window.mainloop()


present_Homepage_gui_frame()
