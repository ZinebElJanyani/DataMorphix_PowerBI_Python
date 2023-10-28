from tkinter import messagebox


from tkinter import *




class loginInterface():
    def __init__(self):
        self.root = Tk()
        self.root.title('Bienvenue !')
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        window_width = screen_width // 2
        window_height = screen_height // 2
        x = (screen_width - window_width) // 2
        y = (screen_height - window_height) // 2

        self.root.geometry('925x500+300+200')  # Set the initial window size
        self.root.minsize(1000, 800)  # Set the minimum window size
        self.root.maxsize(1600, 800)  # Set the maximum window size
        self.root.resizable(False,False);
        self.root.configure(bg="#fff")
        self.widgets(self.root)


    def signIn(self):
        usename = self.user.get()
        password = self.pwd.get()
        if usename == "zineb" and password == "123":
            messagebox.showinfo("Success", "Connexion réussie :)")
            self.root.destroy()



    def widgets(self,r):
        self.img = PhotoImage(file='login.png')
        lbl = Label(r,bg="white",image=self.img)
        lbl.place(x=50,y=50)

        frame = Frame(r,width=350,height=600,bg="white")
        frame.place(x=480,y=70)

        heading = Label(frame,text='Authentification',fg='#57a1f8',bg='white',font=('Microsoft YaHei UI Light',23,'bold'))
        heading.place(x=100,y=5)

        def on_enter(e):
            self.user.delete(0,'end')
        def on_leave(e):
            name=self.user.get()
            if name=="":
                self.user.insert(0,"Email")
        self.user = Entry(frame,fg='gray',border=0,bg='white',font=('Microsoft  YaHei UI Light',11))
        self.user.configure(width=50)
        self.user.place(x=100,y=130)
        self.user.insert(0,'Email')
        self.user.bind('<FocusIn>',on_enter)
        self.user.bind('<FocusOut>', on_leave)
        Frame(frame,width=295,height=2,bg='black').place(x=90,y=160)

        def on_enter2(e):
            self.pwd.delete(0,'end')
        def on_leave2(e):
            cd=self.pwd.get()
            if cd=="":
                self.pwd.insert(0,"mot de passe")
        self.pwd = Entry(frame, fg='gray', border=0, bg='white', font=('Microsoft  YaHei UI Light', 11))
        self.pwd.configure(width=50)
        self.pwd.place(x=100, y=200)
        self.pwd.insert(0, 'mot de passe')
        self.pwd.bind('<FocusIn>', on_enter2)
        self.pwd.bind('<FocusOut>', on_leave2)
        Frame(frame, width=295, height=2, bg='black').place(x=90, y=230)

        Button(frame,width=37,pady=7,text='Sign in', bg='#57a1f8',fg='white',border=0, command=self.signIn).place(x=90,y=290)

        label = Label(frame,text="Vous n'avez pas un compte?",fg='black',bg='white',font=('Microsoft YaHei UI Light',9))
        label.place(x=90,y=340)

        login = Button(frame,width=6,text='Sign up',border = 0,bg='white',cursor='hand2',fg='#57a1f8')
        login.place(x=260, y=340)



        self.root.mainloop(  )












