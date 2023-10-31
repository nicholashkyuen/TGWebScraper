import tkinter as tk
from tkinter import ttk
from tkinter import Canvas
from tkinter import messagebox
import sys
import waybilltoexcel as funcs
from threading import Thread
import time

class App:
    entry = []
    waybill = []


    def __init__(self,master):
        
        self.win = master
        self.win.title("貨倉組 輔助報關程式 v.1.2")
        self.win.resizable(False,False)
        self.win.geometry("500x500")
        self.win.rowconfigure(3,minsize=400,weight=1)
        self.win.protocol("WM_DELETE_WINDOW",self.quit_app)
        
        # Create a new frame for the label and line

        self.top_frame = tk.Frame(self.win)
        self.top_frame.grid(row=0,sticky="ew")
        # self.top_frame.pack(side=tk.LEFT, fill=tk.X,expand=0,anchor="nw")

        self.line_frame = tk.Frame(self.win,height=1, bd=1, relief=tk.SUNKEN, bg='black',width=500)
        self.line_frame.grid(row=1,sticky="ew")
        # self.line_frame.pack(after=self.top_frame,side=tk.LEFT,fill=tk.X,expand=1)

        self.header_frame = tk.Frame(self.win)
        self.header_frame.grid(row=2,sticky="ew")
        # self.header_frame.pack(fill=tk.BOTH,expand=0)

        self.body_frame = tk.Frame(self.win,height=400)
        self.body_frame.grid(row=3,sticky="new",pady=5)
        # self.body_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=1)

        self.body_canvas = Canvas(self.body_frame,highlightthickness=0)
        self.body_canvas.grid(sticky="nsew")
        self.body_frame.rowconfigure(0,minsize=395,weight=1)
        self.body_frame.columnconfigure(0,weight=1)


        scrollbar = ttk.Scrollbar(self.body_frame, orient=tk.VERTICAL, command=self.body_canvas.yview)
        scrollbar.grid(row=0,column=1,sticky="ns")

        
        
        self.body_frame1 = tk.Frame(self.body_canvas,width=500)
        
        self.add_row(True)
        self.body_canvas.configure(yscrollcommand=scrollbar.set)
        self.body_frame1.bind_all("<MouseWheel>",lambda e:self.yview_scroll(e))
        self.body_canvas.bind_all("<MouseWheel>",lambda e: self.yview_scroll(e))

        self.body_canvas.create_window((0,0),window=self.body_frame1, anchor="nw")

        self.bottom_frame = tk.Frame(self.win)
        self.bottom_frame.grid(row=4,sticky="nsew")
        # self.bottom_frame.pack()
        self.bottom_frame.grid_columnconfigure(0,weight=1)
        self.bottom_frame.grid_columnconfigure(1,weight=1)
        
        self.label = tk.Label(self.top_frame, text="Please enter the required information below.",font=('Helvetica',12))
        self.label.grid(row=0,column=0,sticky=tk.W)

        
        
        self.label1 = tk.Label(self.header_frame, text="Forwarder",font=('Helvetica',12))
        self.label1.grid(row=0,column=0,padx=78,ipadx=5,sticky=tk.W)

        self.label2 = tk.Label(self.header_frame, text="Air Waybill",font=('Helvetica',12))
        self.label2.grid(row=0,column=1,ipadx=20)

        runbtn = tk.Button(self.bottom_frame, text="Run",font=('Helvetica',12),command=self.startdl)
        runbtn.grid(row=0,column=1,padx=5,pady=10,ipadx=8,sticky=tk.W)
        closebtn = tk.Button(self.bottom_frame, text="Close",font=('Helvetica',12),command=self.quit_app)
        closebtn.grid(row=0,column=0,padx=5,pady=10,sticky=tk.E)

        self.progress_frame = ttk.Frame(self.win)

        # configrue the grid to place the progress bar is at the center
        self.progress_frame.columnconfigure(0, weight=1)
        self.progress_frame.rowconfigure(0, weight=1)
        self.progress_frame.rowconfigure(1, weight=1)
        # progressbar
        self.pb = ttk.Progressbar(
            self.progress_frame, orient=tk.HORIZONTAL, mode='determinate',length=280)
        self.pb.grid(row=0, column=0, sticky="SEW", padx=20, pady=10)

        value_label = ttk.Label(self.progress_frame, text= "Loading...", font=('Helvetica',12))
        value_label.grid(column=0, row=1, columnspan=2, sticky=tk.N)
        

    def yview_scroll(self,event):
    
        self.body_canvas.configure(scrollregion = self.body_canvas.bbox("all"))
        
        self.body_canvas.yview_scroll(int(-(event.delta / 120)), "units")


    def add_row(self,top=False):
        y = len(App.entry)
        removebtn = tk.Button(self.body_frame1, font=('Helvetica',12))
        removebtn.config(text="Clear" if top else "Remove",width=7)  
        removebtn.config(command=lambda row=y: self.clear_line(row)  if top else \
                         self.remove_row(row))
        removebtn.grid(row=y,column=0,padx=5,sticky=tk.W)
        
        forwarder_combo = ttk.Combobox(self.body_frame1,width=5,font=('Helvetica',12),values=\
                                       ["SF","UPS","DHL","Fedex","TNT"],state="readonly")
        forwarder_combo.grid(row=y,column=1,padx=5,pady=5, ipady=5,sticky=tk.W)
        forwarder_combo.unbind_class("TCombobox","<MouseWheel>")

        waybill_entry = tk.Entry(self.body_frame1,width=26,font=('Helvetica',12))
        waybill_entry.grid(row=y,column=2,padx=10,pady=5, ipady=5,sticky=tk.W)

        addbtn = tk.Button(self.body_frame1, text="Add", relief="raised",font=('Helvetica',12),command=self.add_row)
        addbtn.grid(row=y,column=3, ipadx=8,sticky=tk.W)
        self.body_canvas.configure(scrollregion = self.body_canvas.bbox("all"))
        App.entry.append([y,forwarder_combo,waybill_entry])


    def remove_row(self,row):
        for i in range(len(App.entry)):
            if App.entry[i][0] == row:
                del App.entry [i]
                break
        for item in self.body_frame1.grid_slaves(row=row):
            item.grid_forget()
    
    
    def quit_app(self):
        self.win.destroy()
        sys.exit()


    def startdl(self):
        for item in App.entry:
            if item[1].get()=='' or item[2].get()=='':
                self.error_msg()
                return
            App.waybill.append([item[1].get().upper().strip(),item[2].get().strip()])
        self.progress_frame.grid(row=0, column=0, rowspan=4,columnspan=4,sticky=tk.NSEW)  
        
        Thread(target=self.run_app).start()      
        

    def run_app(self):
        
        step = 85/len(App.waybill)
        self.pb['value'] += 10 
        self.a = funcs.setup(App.waybill)

        for thing in App.waybill:
            forwarder = thing[0]
            if forwarder.lower() == 'ups':
                self.no_info_error(thing[0],thing[1]) if funcs.case_ups(self.a,thing[1]) else None
            elif forwarder.lower() == 'dhl':
                self.no_info_error(thing[0],thing[1]) if funcs.case_dhl(self.a,thing[1]) else None
            elif forwarder.lower() == 'sf':
                self.no_info_error(thing[0],thing[1]) if funcs.case_sf(self.a,thing[1]) else None
            else:
                self.no_info_error(thing[0],thing[1]) if funcs.case_fedex(thing[0],thing[1]) else None
            self.pb['value'] += step
            
            
        for line in App.waybill:
            if line[0] in ["UPS", "DHL", "SF"]:
                self.a.quit()
                break

        try:
            funcs.output_func()
        
        except PermissionError:
            self.error_popup("The file is failed to save!")
    
        except Exception:
            pass             
        self.pb['value'] += 5
        time.sleep(0.5)
        App.waybill.clear()
        self.show_done
        self.progress_frame.grid_forget()
        self.pb['value'] = 0


    def show_done(self):
        messagebox.showinfo("Message", "Done!")


    def error_msg(self):
        messagebox.showerror("Warning", "Please fill in all required information!")


    def error_popup(self, info):
        messagebox.showerror("Warning", "{}".format(info))


    def clear_line(self,row):
        App.entry[row][2].delete(0,tk.END)
        App.entry[row][1].set('')


    def return_waybill(self):
        return App.waybill


    def no_info_error(self,forwarder,waybillno):
        messagebox.showerror("Warning", "Error: [{}] waybill - {} does not exist / has no information, please retry.".format(forwarder,waybillno))
        

    
