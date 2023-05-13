import pandas as pd
from tkinter.filedialog import askopenfilename, asksaveasfilename
import customtkinter as ctk
import tksheet
import win32com.client
import pythoncom


df = pd.DataFrame([[1,2,3,4,5,6,7,8,9,10,11]])
df1 = ['X','Y','Tên HG','CĐ_Đỉnh','CĐ_Đáy','ĐK1','ĐK2','ĐK3','CĐ1','CĐ2','CĐ3']

ctk.set_appearance_mode("dark")
ctk.deactivate_automatic_dpi_awareness()

class ToplevelWindow(ctk.CTkToplevel):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.geometry("300x100")
        self.title("EROR")
        self.label = ctk.CTkLabel(self, text=MN_err)
        self.label.pack(padx=20, pady=20)
        self.lift()


class App(ctk.CTk):
    def __init__(self,*args, **kwargs):
        super().__init__(*args, **kwargs)

        self.title("MN app")
        self.geometry("1450x530")
        self.grid_columnconfigure( 1, weight=1)
        self.lift()

        self.Table_frame = ctk.CTkScrollableFrame(self, height =400 )
        self.Table_frame.grid(row=0, column=0, padx=20, pady=(20,5), columnspan=3,rowspan=6, sticky="ewns")

        self.sheet = tksheet.Sheet(self.Table_frame,width=1200, height= 460, theme = "Dark Blue",  table_bg= "gray13", index_bg= "gray13", header_bg= "gray10", show_x_scrollbar = False, show_y_scrollbar= False)
        self.sheet.default_column_width(width = 106)
        self.sheet.default_row_height(height = 25)
        self.sheet.headers(newheaders = df1)
        
        self.sheet.header_font(newfont = ("Arial", 13, "bold"))
        self.sheet.font(newfont= ("Arial", 11, "normal"))
        self.sheet.align(align = "center", redraw = True)
        self.sheet.grid(padx=20, pady=20)
        self.sheet.set_sheet_data(df.values.tolist(), redraw = True)

        self.button_1 = ctk.CTkButton(self, text="Load Excel file", command=self.load_ex)
        self.button_1.grid(row=0, column=4, padx=(0,20), pady=(20,5), sticky="ewns")

        self.button_2 = ctk.CTkButton(self, text="Load From Cliboard", command=self.load_cl)
        self.button_2.grid(row=1, column=4, padx=(0,20), pady=5, sticky="ewns")

        self.button_3 = ctk.CTkButton(self, text="Show Data", command=self.reload_data)
        self.button_3.grid(row=2, column=4, padx=(0,20), pady=5, sticky="ewns")

        self.button_4 = ctk.CTkButton(self, text="Export to CAD", command=self.Add_HG)
        self.button_4.grid(row=3, column=4, padx=(0,20), pady=5, sticky="ewns")

        self.button_5 = ctk.CTkButton(self, text="Get HG from CAD", command=self.Get_HG)
        self.button_5.grid(row=4, column=4, padx=(0,20), pady=5, sticky="ewns")

        self.button_6 = ctk.CTkButton(self, text="To Excel file", command=self.save_ex)
        self.button_6.grid(row=5, column=4, padx=(0,20), pady=5, sticky="ewns")

        self.button_7 = ctk.CTkButton(self, text="Exit", command=self.quit)
        self.button_7.grid(row=6, column=4, padx=(0,20), pady=5, sticky="ewns")


        self.Table_frame_SW = ctk.CTkFrame(self, height=50)
        self.Table_frame_SW.grid(row=6, column=0, padx=(20,5), pady=5, sticky="ewns")

        self.Table_frame_Progess = ctk.CTkFrame(self, height=50)
        self.Table_frame_Progess.grid(row=6, column=1, padx=(5,20), pady=5,columnspan=2, sticky="ewns")

        self.switch_var = ctk.StringVar(value="dark")
        self.switch = ctk.CTkSwitch(self.Table_frame_SW, text="Dark Mode", command=self.switch_event, variable=self.switch_var, onvalue="dark", offvalue="light", width= 50)

        self.switch.grid(row=1, column=0, padx=20, pady=20)

        self.progressbar_1 = ctk.CTkProgressBar(master=self.Table_frame_Progess, height=10, width=950)

        self.progressbar_1.grid(row=0, column=0, padx=(20, 10), pady=(20, 20))
        self.progressbar_1.configure(mode="determinate")
        self.progressbar_1.set(0)

        self.label_progress = ctk.CTkLabel(self.Table_frame_Progess, text="0/0", fg_color="transparent")
        self.label_progress.grid(row=0, column=2, padx=(20, 10), pady=(20, 20))
        
        self.toplevel_window = None

            
#Hàm khi nút được nhấn
    def switch_event(self):
            switch_var= self.switch.get()
            ctk.set_appearance_mode(switch_var)
            if switch_var == "light":
                self.sheet.set_options(theme = "light blue",table_bg= "gray90", index_bg= "gray90", header_bg= "gray85")
            else:
                self.sheet.set_options(theme = "dark blue",table_bg= "gray13", index_bg= "gray13", header_bg= "gray10")
            print(switch_var)

    def open_toplevel(self):
        if self.toplevel_window is None or not self.toplevel_window.winfo_exists():
            self.toplevel_window = ToplevelWindow(self)  # create window if its None or destroyed
        else:
            self.toplevel_window.focus()
    
    def load_ex(self):
        global df
        file_path = askopenfilename()
        df = pd.read_excel(file_path,'X_CAD', engine='openpyxl')
        df = df.fillna('')
        df.columns =['X','Y', 'THG', 'CD_TOP','CD_BOT','D1','D2','D3','CD1','CD2','CD3']

        for cell_1 in ['X', 'Y', 'CD1', 'CD2', 'CD3', 'CD_TOP', 'CD_BOT']:
            df[cell_1] = df[cell_1].apply(lambda x: format(float(x), ".2f") if x != '' else '')

        for cell_2 in ['D1','D2','D3']:   
            df[cell_2] = df[cell_2].apply(lambda y: format(float(y),".0f") if y != '' else '')

        # for cell_1 in ['X','Y','CD1','CD2','CD3','CD_TOP','CD_BOT']:   
        #     df[cell_1] = df[cell_1].apply(lambda x: format(float(x),".2f"))

        # for cell_2 in ['D1','D2','D3']:   
        #     df[cell_2] = df[cell_2].apply(lambda y: format(float(y),".0f"))
        

    def load_cl(self):
        global df
        df= pd.read_clipboard(header=None).fillna('')
        df.columns =['X','Y', 'THG', 'CD_TOP','CD_BOT','D1','D2','D3','CD1','CD2','CD3']
        for cell_1 in ['X', 'Y', 'CD1', 'CD2', 'CD3', 'CD_TOP', 'CD_BOT']:
            df[cell_1] = df[cell_1].apply(lambda x: format(float(x), ".2f") if x != '' else '')

        for cell_2 in ['D1','D2','D3']:   
            df[cell_2] = df[cell_2].apply(lambda y: format(float(y),".0f") if y != '' else '')

    def quit(self):
        window.destroy()
  
    def reload_data(self):
        global df
        self.sheet.configure(width=1200, height= 25*(len(df)+1))
        self.sheet.set_sheet_data(df.values.tolist(), redraw = True)
        print('clicked')

    def Add_HG(self):
        acad = win32com.client.Dispatch("AutoCAD.Application")
        ms = acad.ActiveDocument.ModelSpace
        global df
        global MN_err
        try:        
            for i, row in df.iterrows():
                x,y,THG,TOP,BOT,D1,D2,D3,CD1,CD2,CD3 = row

                self.label_progress.configure(text=str(i+1)+'/'+str(df.shape[0]))
                self.label_progress.update()

                self.progressbar_1.set((i+1)/df.shape[0])
                self.progressbar_1.update()

                tag_map = {
                "HG": str(THG),
                "TOP": str("{:.2f}".format(float(TOP))),
                "BOT": str("{:.2f}".format(float(BOT))),
                "D1": str('') if D1 == '' else str("%%C{:.0f}".format(float(D1))),
                "D2": str('') if D2 == '' else str("%%C{:.0f}".format(float(D2))),
                "D3": str('') if D3 == '' else str("%%C{:.0f}".format(float(D3))),
                "D1_BOT": str('') if CD1 == '' else str("+{:.2f}".format(float(CD1))),
                "D2_BOT": str('') if CD2 == '' else str("+{:.2f}".format(float(CD2))),
                "D3_BOT": str('') if CD3 == '' else str("+{:.2f}".format(float(CD3)))\
                }

                point = win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8,(float(x),float(y), 0.0))
                new_block = ms.InsertBlock(point, "HG_ALL", 1.0, 1.0, 1.0, 0.0)
        
                for tag in new_block.GetAttributes():
                    if tag.TagString in tag_map:
                        tag.TextString = tag_map[tag.TagString]
        except Exception as MN_err:
            self.open_toplevel()  

    def Get_HG(self):
        acad = win32com.client.Dispatch("AutoCAD.Application")
        global df
        dataHG=[]
        i=0
        for entity in acad.ActiveDocument.ModelSpace:
            i=i+1
            self.label_progress.configure(text=str(i)+'/'+str(acad.ActiveDocument.ModelSpace.Count))
            self.label_progress.update()

            self.progressbar_1.set((i)/acad.ActiveDocument.ModelSpace.Count)
            self.progressbar_1.update()

            if entity.EntityName == 'AcDbBlockReference':
                if entity.EffectiveName == 'HG_ALL':
                    x, y, z = entity.InsertionPoint
                    for attrib in entity.GetAttributes():
                        
                        if attrib.TagString == 'HG':
                            THG = attrib.TextString
                        if attrib.TagString == 'TOP':
                            TOP = attrib.TextString
                        if attrib.TagString == 'BOT':
                            BOT = attrib.TextString
                        if attrib.TagString == 'D1':
                            D1 = attrib.TextString
                        if attrib.TagString == 'D1_BOT':
                            CD1 = attrib.TextString
                        if attrib.TagString == 'D2':
                            D2 = attrib.TextString
                        if attrib.TagString == 'D2_BOT':
                            CD2 = attrib.TextString
                        if attrib.TagString == 'D3':
                            D3 = attrib.TextString
                        if attrib.TagString == 'D3_BOT':
                            CD3 = attrib.TextString
                        
                    dataHG.append(["{:.2f}".format(float(x)), "{:.2f}".format(float(y)),THG,TOP,BOT,D1,D2,D3,CD1,CD2,CD3])
                else:
                    continue
        df = pd.DataFrame(dataHG, columns=['X','Y', 'THG', 'CD_TOP','CD_BOT','D1','D2','D3','CD1','CD2','CD3'])
                
        for DK_C in ['D1','D2','D3']:   
            df[DK_C] = df[DK_C].str.replace('%%C','')
        for CD_C in ['CD1','CD2','CD3']:   
            df[CD_C] = df[CD_C].str.replace('+','')

    def save_ex(self):
        global df
        file_path = asksaveasfilename(defaultextension='.xlsx')
        df.to_excel(file_path,'X_CAD', index=False)

if __name__ == '__main__':
    window = App()
    window.mainloop()