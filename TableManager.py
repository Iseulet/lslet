import os
import table as tb
import tkinter as tk
from tkinter import *
from tkinter import Tk, ttk, Label, LabelFrame, Scrollbar, Listbox, filedialog
# from tkinter.messagebox import showinfo, showerror, showwarning


from comfunc import search_file
from ExcelFunc import *

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title('Table Manager')
        self.iconbitmap (search_file ('rin.ico'))


class MainFrame(ttk.Frame):
    def __init__(self, container):
        super().__init__(container)
        self.initUI()

        toolinfo = 'tablemanager_info.txt'
        root_path = search_file ('tablemanager_info.txt')
        if root_path == None:
            fld_select = filedialog.askdirectory (title = '폴더 경로를 지정하세요')
            toolpath = os.path.dirname(__file__)
            with open (toolpath + '/'+ toolinfo,'w', encoding='utf8') as info_file :    
                info_file.write (fld_select)
                self.default_path = fld_select
        else :
            with open (root_path, 'r', encoding='utf8') as info_file : 
                self.default_path = info_file.readline()
                self.csv_path = info_file.readline()
        
        print (self.default_path)
        print (self.csv_path)


    def initUI (self):
        options = {'padx': 2, 'pady': 2}

        frame1 = LabelFrame(self, text='WorkGroup')
        frame1.pack(fill='both')

        frame2 = LabelFrame(self, text='Excel')
        frame2.pack(side='left')

        frame3 = LabelFrame(self, text='Table')
        frame3.pack(side='left')


        self.WGroup(frame1)
        self.list_file_manage (frame2)
        self.list_tbl_manage (frame3)
        self.update()

        self.pack(**options)

    def update (self):
        self.cbox_wgroup.current(0)
        self.load_lst_tbl(self.cbox_wgroup)
        self.renew_lst_tbl(self.cbox_wgroup)



    #frame1
    # about wgroup
    def load_lst_tbl (self, cbox):
        cbox = cbox.get()
        for c in self.wg_col :
            return self.wg_dtbl[c]
    def renew_lst_tbl (self, cbox):
        sel = cbox.get()
        self.lst_excel.delete (0, END)
        for f in self.wg_dtbl[sel]:
            self.lst_excel.insert(END, f) 
    def cmd_reload (self):
        wgroup = tb.excel ('workgroup.xlsx')
        self.load_lst_tbl(self.cbox_wgroup)
        self.renew_lst_tbl(self.cbox_wgroup)
    def cmd_wgroup_open (self):
        wgroup = tb.excel ('workgroup.xlsx')
        wgroup.openxl()

    def WGroup (self, frame):
        # workgroup
        wgroupfile = 'workgroup.xlsx'
        wgroup_xl = tb.excel (wgroupfile)

        self.wg_col, self.wg_dtbl = wgroup_xl.readxltbl ('workgroup')
        
        self.cbox_wgroup = ttk.Combobox(frame, width = 10, values = self.wg_col, state = 'readonly') #### 이거 통째로 함수로 하던중
        self.cbox_wgroup.pack(side = 'left', padx = 2, pady = 2)
       
        btn_wgroup_reload = ttk.Button(frame, text='Reload')
        btn_wgroup_reload['command'] = self.cmd_reload
        btn_wgroup_reload.pack(side= 'right' )

        btn_wgroup_open = ttk.Button(frame, text='Open')
        btn_wgroup_open['command'] = self.cmd_wgroup_open
        btn_wgroup_open.pack(side= 'right' )




    #frame2
    # about list file
    def list_file_manage (self, frame):
        options = {'padx': 2, 'pady': 2}
        options_btn = {'padx': 2, 'pady': 2, 'sticky' : W+E+N+S}

        frame_btn_upper = Frame(frame)
        frame_btn_upper.pack(fill = 'both', expand=True)
        frame_lst = Frame(frame)
        frame_lst.pack(fill='both')
        frame_btn_lower = Frame(frame)
        frame_btn_lower.pack(fill = 'both', expand=True)

        #list
        scrollbar_excel = Scrollbar (frame_lst)
        scrollbar_excel.pack(side = 'right', fill = 'y' )
        self.lst_excel = Listbox (frame_lst, selectmode = "browse", height = 15, width = 30, yscrollcommand = scrollbar_excel.set)
        self.lst_excel.pack(side = 'left', fill = 'both', expand = True, **options)
        scrollbar_excel.config (command = self.lst_excel.yview)

        #btn
        btn_open = ttk.Button(frame_btn_upper, text='Open Excel')
        btn_open_a = ttk.Button(frame_btn_upper, text='Open All Excel')
        btn_export = ttk.Button(frame_btn_lower, text='Export Excel')
        btn_export_a = ttk.Button(frame_btn_lower, text='Export All Excel')

        btn_open_a.pack(side = 'right',**options)
        btn_open.pack(side = 'right', **options)
        btn_export.pack(side = 'right',**options)        
        btn_export_a.pack(side = 'right',**options)
        # btn_validate.pack(**options)

        btn_open['command'] = self.cmd_open_excel
        btn_open_a ['command'] = self.cmd_open_all_excel
        btn_export['command'] = self.cmd_export_excel
        btn_export_a['command'] = self.cmd_export_all_excel
        # btn_validate['command'] = self.cmd_validate

        # btn_open.grid (row = 0, column= 0, **options_btn)
    def cmd_open_excel (self):
        for i in self.lst_excel.curselection():
            excel = tb.excel(self.lst_excel.get(i))
            excel.openxl()

    def cmd_open_all_excel (self):
        for file in self.lst_excel.get(0, END):
            print (file)
            excel = tb.excel(file)
            excel.openxl()
            
    def cmd_export_excel (self):
        for i in self.lst_excel.curselection():
            wb = self.lst_excel.get(i)
            excel = tb.excel(wb)
            tables = excel.tablelistxl()
            for table in tables:
                tbl_df = tb.sheettable(wb, table, table)
                tbl_df.export(self.csv_path)

    def cmd_export_all_excel (self):
        for wb in self.lst_excel.get(0, END):
            excel = tb.excel(wb)
            tables = excel.tablelistxl()
            for table in tables:
                tbl_df = tb.sheettable(wb, table, table)
                tbl_df.export(self.csv_path)



    def list_tbl_manage (self, frame):
        options = {'padx': 2, 'pady': 2}

        frame_btn_upper = Frame(frame)
        frame_btn_upper.pack(fill = 'both', expand=True)
        frame_lst = Frame(frame)
        frame_lst.pack(fill='both')
        frame_btn_lower = Frame(frame)
        frame_btn_lower.pack(fill = 'both', expand=True)

        scrollbar_tbl = Scrollbar (frame_lst)
        scrollbar_tbl.pack(side = 'right', fill = 'y' )
        self.lst_tbl = Listbox (frame_lst, selectmode = "extended", height = 15, width = 30, yscrollcommand = scrollbar_tbl.set)
        self.lst_tbl.pack(side = 'left', fill = 'both', expand = True, **options)
        scrollbar_tbl.config (command = self.lst_tbl.yview)

        btn_load = ttk.Button(frame_btn_upper, text='Load Table')

        btn_export = ttk.Button(frame_btn_lower, text='Export Table')
        # btn_validate = ttk.Button(frame_btn_lower, text='Check Validate')

        btn_load.pack(side = 'left', **options)
        btn_export.pack(side = 'right',**options)
        # btn_validate.pack(**options)

        btn_load['command'] = self.cmd_table_load
        btn_export['command'] = self.cmd_export_table

        # btn_validate['command'] = self.cmd_validate

    def cmd_table_load(self):
        for i in self.lst_excel.curselection():
            wb = self.lst_excel.get(i)
        excel = tb.excel(wb)
        lst_tbl= excel.tablelistxl()

        self.lst_tbl.delete (0, END)
        for f in lst_tbl:
            self.lst_tbl.insert(END, f)

        self.lst_tbl_name = wb
                    
    def cmd_export_table (self):
        for i in self.lst_tbl.curselection():
            table = self.lst_tbl.get(i)
            tbl_df = tb.sheettable(self.lst_tbl_name, table, table)
            tbl_df.export(self.csv_path)
    def cmd_checkout (self):
        pass
    def cmd_validate (self):
        pass

if __name__ == "__main__":
    app = App()
    frame = MainFrame(app)
    app.resizable (False, False)
    app.mainloop()