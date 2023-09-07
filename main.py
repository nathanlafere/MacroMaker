from tkinter import messagebox, ttk
from tkinter import *
import keyboard
import pyautogui
import os
import pickle

class Main:
    def __init__(self, root):
        # [0] = type of event [1] = coordinates or values [2] = delay in milliseconds
        self.event_list = [("Mouse left click",(105,20),15),("Mouse right click",(12,44),4),("copy value",("path","cel coordinates"),0),("cop3y value",("path","cel coordinates"),0),("copy va12lue",("path","cel coordinates"),0),("copy va1lue",("path","cel coordinates"),0),("copy valu1e",("path","cel coordinates"),0),("copy 1value",("path","cel coordinates"),0),("cop1y value",("path","cel coordinates"),0)]
        #setting window size
        root.configure(padx=7,pady=7,width=413,height=413)
        #placing the frames
        self.new_event_frame = Frame(root,background="gray81",borderwidth=4, relief=GROOVE)
        self.new_event_frame.place(x=200,y=0,height=340,width=200)
        
        self.event_type = ttk.Combobox(self.new_event_frame, state="readonly",values=["Mouse click", "Copy or insert XLSX values", "Press keys", "Definir condição de fulga"],width=27)
        self.event_type.place(anchor=CENTER,relx=0.497,rely=0.04) # criar um event handler que altera os itens abaixo do combobox para os inputs adequados para cada opção
        self.event_type.bind("<<ComboboxSelected>>",self.change_new_event_frame)
        self.inside_new_event = Frame(self.new_event_frame)
        self.inside_new_event.place(height=298,width=183,relx=0.019,rely=0.09)
        config_frame = Frame(root,background="gray81",borderwidth=4, relief=GROOVE)
        config_frame.place(x=200,y=340,height=60,width=200)
        play_btn = Button(config_frame, font=("Onyx",7,"bold"),bg="gray55",fg="red3",text="►",height=2, width=3)
        play_btn.place(anchor=CENTER,relx=0.15,rely=.5)
        load_btn2 = Button(config_frame, font=("Onyx",7,"bold"),bg="gray55",text="⚜",height=2, width=3)
        load_btn2.place(anchor=CENTER,relx=0.32,rely=.5)
        self.queue_frame = Frame(root,background="gray81",borderwidth=4, relief=GROOVE)
        self.queue_canvas = Canvas(self.queue_frame)
        self.queue_canvas.place(height=390,width=180)
        self.queue_frame.place(x=0,y=0,height=400,width=200)
        queue_scrollbar = Scrollbar(self.queue_frame, orient=VERTICAL,command=self.queue_canvas.yview)
        queue_scrollbar.place(anchor="se",relx=1,rely=1,height=392,width=13)
        self.queue_canvas.configure(yscrollcommand=queue_scrollbar.set)
        self.queue_canvas.bind('<Configure>',  lambda e: self.queue_canvas.configure(scrollregion=self.queue_canvas.bbox("all")))
        self.queue_canvas.bind_all("<MouseWheel>", self._on_mousewheel)
        for event in self.event_list:
            self.event_display(event)
        
    def _on_mousewheel(self,event):
        self.queue_canvas.yview_scroll(int(-1*(event.delta/120)), "units")
    
    def change_new_event_frame(self,event):
        self.clear_frame(self.inside_new_event)
        new_event_info = None
        match event.widget.get():
            case "Mouse click":
                left_or_right = ttk.Combobox(self.inside_new_event,state= "readonly",values=["left click","right click"], width=25)
                left_or_right.current(0)
                left_or_right.grid(columnspan=8,pady=5, padx=5)
                
                delay_text = Label(self.inside_new_event, text="Delay :", justify='left')
                delay_text.grid(row=1, pady=5, columnspan=4, sticky=W)
                delay_entry = Entry(self.inside_new_event,width=11)
                delay_entry.grid(row=1, pady=5, padx=41, columnspan=4, sticky=W)
                
                first_entry = self.create_input("X :", 0, 1, 2)
                second_entry = self.create_input("Y :", 2, 3, 2)
                checkbox_cvar = IntVar() 
                checkbox = self.create_checkbox(self.inside_new_event,checkbox_cvar,event)
                checkbox.grid(row=3,column=0,columnspan=4,sticky=W)
                new_event_info = (left_or_right,first_entry,second_entry,checkbox_cvar)
            case "Press keys":
                keys_list = list(keyboard._os_keyboard.from_name.keys())
                keys_combobox = ttk.Combobox(self.inside_new_event,values=keys_list, width=25)
                keys_combobox.grid(row=0,column=0,padx=5,sticky=N,pady=5)
                delay_text = Label(self.inside_new_event, text="Delay :", justify='left')
                delay_text.grid(row=1, pady=5, columnspan=4, sticky=W)
                delay_entry = Entry(self.inside_new_event,width=11)
                delay_entry.grid(row=1, pady=5, padx=41, columnspan=4, sticky=W)
                new_event_info = (keys_combobox,delay_entry)
            case "Copy or insert XLSX values":
                copy_or_insert = ttk.Combobox(self.inside_new_event,state= "readonly",values=["Copy","Insert"], width=25)
                copy_or_insert.current(0)
                copy_or_insert.grid(columnspan=8,pady=5, padx=5)
                first_entry = self.create_input("C :", 0, 1, 1)
                second_entry = self.create_input("R :", 2, 3, 1)
                checkbox_cvar = IntVar() 
                checkbox = self.create_checkbox(self.inside_new_event,checkbox_cvar,event)
                checkbox.grid(row=2,column=0,columnspan=4,padx=0,sticky=W)
                new_event_info = (copy_or_insert,first_entry,second_entry,checkbox_cvar)
        confirm_button = Button(self.inside_new_event, text="Confirm", font=("Helvetica",8,"bold"),bg="gray55",fg="gray15", width=10 ,relief=GROOVE, command=lambda: self.confirm_new_event(new_event_info))
        confirm_button.place(relx=0.5,rely=0.94,anchor="center")
    
    def confirm_new_event(self,new_event_info):
        new_event = ()
        match self.event_type.get():
            case "Mouse click":
                if (new_event_info[1].get() != '' and new_event_info[2].get() != ''):
                    self.empty_entry(self.inside_new_event.grid_slaves()[-3])
                    if new_event_info[3].get() == 0:
                        new_event = (f'Mouse {new_event_info[0].get()}',(new_event_info[1].get(),new_event_info[2].get()),self.inside_new_event.grid_slaves()[-3].get())
                    else:
                        self.empty_entry(self.inside_new_event.grid_slaves()[0])
                        self.empty_entry(self.inside_new_event.grid_slaves()[2])
                        new_event = (f'Mouse {new_event_info[0].get()}',(new_event_info[1].get(),new_event_info[2].get(),self.inside_new_event.grid_slaves()[2].get(),self.inside_new_event.grid_slaves()[0].get()),self.inside_new_event.grid_slaves()[-3].get())
            case "Press keys":
                if new_event_info[0].get() != '':
                    self.empty_entry(new_event_info[1])
                    new_event = ("Press key",new_event_info[0].get(),new_event_info[1].get())
            case "Copy or insert XLSX values":
                if (new_event_info[1].get() != '' and new_event_info[2].get() != ''):
                    if new_event_info[3].get() == 0:
                        new_event = (new_event_info[0].get(),(new_event_info[1].get(),new_event_info[2].get()))
                    else:
                        self.empty_entry(self.inside_new_event.grid_slaves()[0])
                        self.empty_entry(self.inside_new_event.grid_slaves()[2])
                        new_event = (new_event_info[0].get(),(new_event_info[1].get(),new_event_info[2].get(),self.inside_new_event.grid_slaves()[2].get(),self.inside_new_event.grid_slaves()[0].get()))

        if new_event:
            self.event_list.append(new_event)
            self.event_display(new_event)
            self.queue_canvas.configure(scrollregion=self.queue_canvas.bbox("all"))
        else:
            messagebox.showinfo(title="Error", message="Invalid entry, please try again")

    def empty_entry(self, entry):
        if entry.get() == '':
            entry.insert(0,0)
    
    def create_variation_input(self,event,variable):
        match event.widget.get():
            case "Mouse click":
                if variable.get() == 1:
                    self.create_input("X :", 0, 1, 4)
                    self.create_input("Y :", 2, 3, 4)
                else:
                    for label in self.inside_new_event.grid_slaves():
                        if int(label.grid_info()["row"]) > 3:
                            label.destroy()
            case "Copy or insert XLSX values":
                if variable.get() == 1:
                    self.create_input("C :", 0, 1, 3)
                    self.create_input("R :", 2, 3, 3)
                else:
                    for label in self.inside_new_event.grid_slaves():
                        if int(label.grid_info()["row"]) > 2:
                            label.destroy()
        
    def create_input(self, text, column1, column2, row):
        text = Label(self.inside_new_event, text=text, justify='left')
        text.grid(row=row, column=column1, pady=5)
        entry = Entry(self.inside_new_event,width=11)
        entry.grid(row=row,column=column2, pady=5)
        return entry
        
    def event_display(self,event):
        event_frame = Frame(self.queue_canvas,background="gray63",borderwidth=4, height=70,width=179,relief=GROOVE)
        self.queue_canvas.create_window((0,0+self.event_list.index(event)*70),window=event_frame,anchor="nw")
        clear_btn = Button(event_frame, font=("Helvetica",6,"bold"),bg="gray55",fg="gray20",text="X",height=1, width=1, command=lambda:self.clear_event(event))
        clear_btn.place(anchor='e',relx=0.997,rely=0.15)
        up_btn = Button(event_frame, font=("Onyx",6),bg="gray55",fg="red3",text="▲",height=1, width=1, command=lambda:self.change_order(event,-1))
        up_btn.place(anchor='e',relx=0.821,rely=0.15)
        down_btn = Button(event_frame, font=("Onyx",6),bg="gray55",fg="red3",text="▼",height=1, width=1, command=lambda:self.change_order(event,1))
        down_btn.place(anchor='e',relx=0.909,rely=0.15)
        type_text = Label(event_frame,text=event[0].capitalize(),bg="gray63",border=True, font=("Nyala",9,"bold"))

        if event[0] in ["Mouse right click", "Mouse left click", "Press keys"]:
            delay_text = Label(event_frame,text=f'Delay={event[2]}',bg="gray63",border=True, font=("Nyala",8))
            delay_text.place(relx=0.01,rely=0.65)
        if "click" in event[0]:
            values_text = Label(event_frame,text=f'X={event[1][0]} Y={event[1][1]}',bg="gray63",border=True, font=("Nyala",8))
        elif event[0] in ["Copy","Insert"]:
            values_text = Label(event_frame,text=f'{event[0]} value in row {event[1][0]} column {event[1][1]}',bg="gray63",border=True, font=("Nyala",8))
        elif "Press" in event[0]:
            values_text = Label(event_frame,text=f'Press key {event[1].capitalize()}',bg="gray63",border=True, font=("Nyala",8))
        else:
            values_text = Label(event_frame,text="Error",bg="gray63",border=True, font=("Nyala",8))


        type_text.place(relx=0.01)
        values_text.place(relx=0.01,rely=0.35)
        
    def clear_frame(self,frame):
        for widget in frame.winfo_children():
            widget.destroy()
    def clear_event(self,event):
        self.event_list.remove(event)
        self.reload_event()
        
    def reload_event(self):
        self.clear_frame(self.queue_canvas)
        for event in self.event_list:
            self.event_display(event)
        self.queue_canvas.configure(scrollregion=self.queue_canvas.bbox("all"))
    
    def change_order(self,event,direction):
        if (self.event_list.index(event)+direction >= 0 and self.event_list.index(event)+direction < len(self.event_list)):
            older_index = self.event_list.index(event)
            new_index = self.event_list.index(event)+direction
            older_value = self.event_list[new_index]
            self.event_list[new_index] = self.event_list[older_index]; self.event_list[older_index] = older_value
            self.reload_event()
        
    def create_checkbox(self,frame,variable,event):
        return Checkbutton(frame, text="Variation per loop", variable=variable, onvalue=1, offvalue=0,command=lambda: self.create_variation_input(event,variable))




if __name__ == "__main__":
    root = Tk(className=" Macro Maker")
    main = Main(root)
    root.mainloop()

