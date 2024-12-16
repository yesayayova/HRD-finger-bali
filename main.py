import pandas as pd
import numpy as np
import OleFileIO_PL # type: ignore
from tkinter import *
from tkinter import filedialog, ttk
import pickle
from sklearn.ensemble import RandomForestClassifier

last_dir = "C:/"
gjc_dom = ""
gjc_inter = ""
dk_inter = ""
web = ""

def error_file():
    top_error_file = Toplevel()
    top_error_file.title("Error!")
    
    width = 400
    height = 80
    screen_width = top_error_file.winfo_screenwidth()
    screen_height = top_error_file.winfo_screenheight()
    x = (screen_width / 2) - (width / 2)
    y = (screen_height / 2) - (height / 2)
    top_error_file.geometry(f"{width}x{height}+{int(x)}+{int(y)}")
    top_error_file.resizable(False, False)

    label_error = Label(top_error_file, text="The file format is not supported!")
    label_error.place(x=120, y=25)

def saved_success():
    top_error_file = Toplevel()
    top_error_file.title("")
    
    width = 200
    height = 80
    screen_width = top_error_file.winfo_screenwidth()
    screen_height = top_error_file.winfo_screenheight()
    x = (screen_width / 2) - (width / 2)
    y = (screen_height / 2) - (height / 2)
    top_error_file.geometry(f"{width}x{height}+{int(x)}+{int(y)}")
    top_error_file.resizable(False, False)

    label_error = Label(top_error_file, text="Successfully Saved!")
    label_error.place(x=47, y=25)

def saved_failed():
    top_error_file = Toplevel()
    top_error_file.title("")
    
    width = 200
    height = 80
    screen_width = top_error_file.winfo_screenwidth()
    screen_height = top_error_file.winfo_screenheight()
    x = (screen_width / 2) - (width / 2)
    y = (screen_height / 2) - (height / 2)
    top_error_file.geometry(f"{width}x{height}+{int(x)}+{int(y)}")
    top_error_file.resizable(False, False)

    label_error = Label(top_error_file, text="Failed to save")
    label_error.place(x=60, y=25)

def process_raptor(path_dk, path_gjc_inter, path_gjc_dom):

        list_path = [path_dk, path_gjc_inter, path_gjc_dom]
        list_df = []
        len_per_outlet = []

        for i, path in enumerate(list_path):
            if path == "":
              continue
            
            df = pd.read_excel(path, skiprows=9)
            df = df.iloc[:,:4]
            df = df.dropna(how='all')
            df = df.dropna(subset=list(df.columns[:3]),how='all')

            ids = []
            names = []
            id = ""
            name = ""

            for index, row in df.iterrows():
                if (type(row[0]) == str) and (type(row[1]) != float) and (type(row[2]) != str) :
                    id = row[0][-4:]
                    name = row[1][3:]
                    # print(id, name)
                ids.append(id)
                names.append(name)

            df['Name'] = names
            df['ID'] = ids

            df = df[~df['Unnamed: 0'].apply(lambda x: isinstance(x, str))]
            df.columns = ['Date In', 'Time In', 'Date Out', 'Time Out', 'Name', 'ID']
            df.fillna(0, inplace=True)
            df.reset_index(inplace=True, drop=True)

            len_per_outlet.append(df.shape[0])

            list_df.append(df)

        df_final = pd.concat(list_df, ignore_index=True)

        list_outlet = []

        for i, len_outlet in enumerate(len_per_outlet):
            if i == 0:
                outlet = ['DK Inter' for i in range(len_outlet)]
            elif i == 1:
                outlet = ['GJC Inter' for i in range(len_outlet)]
            elif i == 2:
                outlet = ['GJC Dom' for i in range(len_outlet)]
            list_outlet += outlet

        df_final['Outlet'] = list_outlet
        df_final['Status In'] = ["Masuk" for i in range(df_final.shape[0])]
        df_final['Status Out'] = ["Keluar" for i in range(df_final.shape[0])]
        df_final = df_final[['ID', 'Name', 'Outlet', 'Date In', 'Time In', 'Status In', 'Date Out', 'Time Out', 'Status Out']]

        def filter_tanggal(tgl):
            tgl = str(tgl)
            if tgl == '0':
                return "00-00-0000"
            
            tgl = tgl[:10]
            tahun = tgl[:4]
            bulan = tgl[5:7]
            hari = tgl[8:]
            return hari+"-"+bulan+"-"+tahun

        # combined_df['Tanggal scan'] = combined_df['Tanggal scan'].apply(filter_tanggal)
        # return combined_df
        df_final['Date In'] = df_final['Date In'].astype(str)
        df_final['Date Out'] = df_final['Date Out'].astype(str)
        df_final['Date In'] = df_final['Date In'].apply(filter_tanggal)
        df_final['Date Out'] = df_final['Date Out'].apply(filter_tanggal)

        # return df_final

        combined_rows = []

        # Iterate through each row in the original dataframe and split into new structure
        for index, row in df_final.iterrows():
            combined_rows.append({
                'ID': row['ID'],
                'Name': row['Name'],
                'Outlet': row['Outlet'],
                'Date': row['Date In'],
                'Time': row['Time In'],
                'Status':row['Status In']
            })
            combined_rows.append({
                'ID': row['ID'],
                'Name': row['Name'],
                'Outlet': row['Outlet'],
                'Date': row['Date Out'],
                'Time': row['Time Out'],
                'Status':row['Status Out']
            })

        combined_df = pd.DataFrame(combined_rows)
        combined_df.columns = ['NIP', 'Nama', 'Mesin', 'Tanggal scan', 'Jam', 'Status']
        combined_df['Tanggal scan'] = combined_df['Tanggal scan'].astype(str)

        return combined_df

def show(frame, df):
    my_tree = ttk.Treeview(frame)

    # Clear old treeview
    my_tree.delete(*my_tree.get_children())

    # Set up new tree
    my_tree['column'] = list(df.columns)
    my_tree['show'] = "headings"
        
    # Set up all column names
    for column in my_tree['column']:
        my_tree.heading(column, text=column, anchor=W)
        if column == "Nama":      
            my_tree.column(column, width=150)
        else:
            my_tree.column(column, width=60)

    # Set up all rows
    df_rows = df.to_numpy().tolist()
    
    for row in df_rows:
        my_tree.insert("", "end", values=row)

    my_tree.pack(expand=True, fill='both')
    my_tree.place(x=0, y=0, width=440, height=575)
        
    scroll_y = Scrollbar(my_tree, orient='vertical', command=my_tree.yview)
    scroll_y.place(relx=1, rely=0, relheight=1, anchor='ne')

    scroll_x = Scrollbar(my_tree, orient='horizontal', command=my_tree.xview)
    scroll_x.pack(side='bottom', fill='x')

    my_tree.configure(xscrollcommand=scroll_x.set, yscrollcommand=scroll_y.set)

def process_report(path_report):
    if path_report == "":
       return 0
    try:
      ole = OleFileIO_PL.OleFileIO(path_report)

      if ole.exists('Workbook'):
          d = ole.openstream('Workbook')
          df=pd.read_excel(d,engine='xlrd')
    except:
      df=pd.read_excel(path_report)
      
    df['Mesin'] = ['REPORT SCAN' for i in range(df.shape[0])]

    def edit_tanggal(tgl):
        year = tgl[:4]
        month = tgl[5:7]
        day=tgl[8:10]
        return day+"-"+month+"-"+year

    def edit_scan(scan):
        if scan=='Absensi Masuk':
            return 'Masuk'
        else:
            return 'Keluar'
    df['Tanggal Absensi'] = df['Tanggal Absensi'].astype(str)
    df['Tanggal Absensi'] = df['Tanggal Absensi'].apply(edit_tanggal)
    df['Tipe Absensi'] = df['Tipe Absensi'].apply(edit_scan)

    data = df[['PIN', 'Nama Karyawan', 'Mesin','Tanggal Absensi', 'Jam Absensi', 'Tipe Absensi']]
    data.columns = ['NIP', 'Nama', 'Mesin', 'Tanggal scan', 'Jam', 'Status']

    data['Tanggal scan'] = data['Tanggal scan'].astype(str)
    data['Jam'] =  data['Jam'].astype(str)
    data['datetime'] = data['Tanggal scan'] + ' ' + data['Jam']
    
    dates = []
    for date in data['datetime']:
      x = pd.to_datetime([date])
      dates.append(x[0])
    data['Datetime Scan'] = dates

    data['Hour'] = data['Datetime Scan'].dt.hour
    data['Day of Week'] = data['Datetime Scan'].dt.dayofweek

    def labeling(label):
      if label == "Masuk":
        return 0
      else:
        return 1

    data['Status'] = data['Status'].apply(labeling)

    # data['Prev Scan Type'] = data['Status'].shift(1)
    # data['Next Scan Type'] = data['Status'].shift(-1)
    data['Prev Scan Time'] = pd.to_datetime(data['Jam'].shift(1), errors='coerce').dt.hour
    data['Next Scan Time'] = pd.to_datetime(data['Jam'].shift(-1), errors='coerce').dt.hour

    data.fillna(method='bfill', inplace=True)
    data.fillna(method='ffill', inplace=True)

    X = data[['NIP', 'Prev Scan Time', 'Next Scan Time','Hour', 'Day of Week']]
    y = data['Status']

    loaded_model = pickle.load(open('rf_model_2.sav', 'rb'))
    y_pred = loaded_model.predict(X)
    data["Status"] = y_pred

    data = data[['NIP', 'Nama', 'Mesin', 'Tanggal scan', 'Jam', 'Status']]

    def labeling_2(label):
      if label == 0:
        return 'Masuk'
      else:
        return 'Keluar'

    data['Status'] = data['Status'].apply(labeling_2)
    return data

def main():
    global last_dir, gjc_dom, gjc_inter, dk_inter, web

    def dummy():
       pass

    root = Tk() 
    root.title('Mutasi Reader - Bogajaya')

    width = 800
    height = 600

    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()

    x = (screen_width / 2) - (width / 2)
    y = (screen_height / 2) - (height / 2)

    root.geometry(f"{width}x{height}+{int(x)}+{int(y)}")
    root.resizable(False, False)

    # my_menu = Menu(root)

    # file_menu = Menu(my_menu, tearoff=False)
    # my_menu.add_cascade(label='File', menu=file_menu)
    # file_menu.add_command(label='Open File ...        ', command=dummy)
    # file_menu.add_separator()
    # file_menu.add_command(label='Save                 ', command=save)
    # file_menu.add_separator()
    # file_menu.add_command(label='Exit                 ', command=root.quit)
    
    # help_menu = Menu(my_menu, tearoff=False)
    # my_menu.add_cascade(label='About', menu=help_menu)
    # help_menu.add_command(label='Mutasi Reader Version 1.0.0', command=dummy)
    # root.config(menu=my_menu)

    input_paned = PanedWindow(bd=2, relief="groove")
    input_paned.place(x=10, y=15, width=335, height=575)

    label_gjc_dom = Label(input_paned, text='GJC Dom')
    label_gjc_dom.place(x=10, y=10)
    entry_gjc_dom = Entry(input_paned, width=38)
    entry_gjc_dom.place(x=12, y=35)
    openfile_btn_gjc_dom = Button(input_paned, 
                        text='Open File', 
                        width=8, 
                        command=lambda :openfile_gjc_dom(), 
                        relief="ridge", 
                        borderwidth=1, 
                        border=1)
    openfile_btn_gjc_dom.place(x=250, y=32)

    def openfile_gjc_dom():
      global last_dir, gjc_dom

      filename = filedialog.askopenfilename(initialdir=last_dir,
                                              title="Open File",
                                              filetypes=(('Excel Files', '*.xl*'), ('All Files', '*.*')))
      gjc_dom = r"{}".format(filename)
      entry_gjc_dom.delete(0, END)
      entry_gjc_dom.insert(0, gjc_dom.split('/')[-1])
      last_dir = filename.rsplit('/',1)[0]

    label_gjc_inter = Label(input_paned, text='GJC Inter')
    label_gjc_inter.place(x=10, y=65)
    entry_gjc_inter = Entry(input_paned, width=38)
    entry_gjc_inter.place(x=12, y=90)
    openfile_btn_gjc_inter = Button(input_paned, 
                        text='Open File', 
                        width=8, 
                        command=lambda :openfile_gjc_inter(), 
                        relief="ridge", 
                        borderwidth=1, 
                        border=1)
    openfile_btn_gjc_inter.place(x=250, y=87)

    def openfile_gjc_inter():
      global last_dir, gjc_inter

      filename = filedialog.askopenfilename(initialdir=last_dir,
                                              title="Open File",
                                              filetypes=(('Excel Files', '*.xl*'), ('All Files', '*.*')))
      gjc_inter = r"{}".format(filename)
      entry_gjc_inter.delete(0, END)
      entry_gjc_inter.insert(0, gjc_inter.split('/')[-1])
      last_dir = filename.rsplit('/',1)[0]

    label_dk_inter = Label(input_paned, text='DK Inter')
    label_dk_inter.place(x=10, y=120)
    entry_dk_inter = Entry(input_paned, width=38)
    entry_dk_inter.place(x=12, y=145)
    openfile_btn_dk_inter = Button(input_paned, 
                        text='Open File', 
                        width=8, 
                        command=lambda :openfile_dk_inter(), 
                        relief="ridge", 
                        borderwidth=1, 
                        border=1)
    openfile_btn_dk_inter.place(x=250, y=142)

    def openfile_dk_inter():
      global last_dir, dk_inter

      filename = filedialog.askopenfilename(initialdir=last_dir,
                                              title="Open File",
                                              filetypes=(('Excel Files', '*.xl*'), ('All Files', '*.*')))
      dk_inter = r"{}".format(filename)
      entry_dk_inter.delete(0, END)
      entry_dk_inter.insert(0, dk_inter.split('/')[-1])
      last_dir = filename.rsplit('/',1)[0]

    label_web = Label(input_paned, text='Report Scan')
    label_web.place(x=10, y=175)
    entry_web = Entry(input_paned, width=38)
    entry_web.place(x=12, y=200)
    openfile_btn_web = Button(input_paned, 
                        text='Open File', 
                        width=8, 
                        command=lambda :openfile_web(), 
                        relief="ridge", 
                        borderwidth=1, 
                        border=1)
    openfile_btn_web.place(x=250, y=197)

    def openfile_web():
      global last_dir, web

      filename = filedialog.askopenfilename(initialdir=last_dir,
                                              title="Open File",
                                              filetypes=(('Excel Files', '*.xl*'), ('All Files', '*.*')))
      web = r"{}".format(filename)
      entry_web.delete(0, END)
      entry_web.insert(0, web.split('/')[-1])
      last_dir = filename.rsplit('/',1)[0]

    save_btn = Button(input_paned,
                      text="Process",
                      width=14,
                      command=lambda : process(dk_inter, gjc_inter, gjc_dom, web),
                      relief="ridge",
                      borderwidth=1,
                      border=1)
    save_btn.place(x=100, y=255)

    def process(path_dk, path_gjc_inter, path_gjc_dom, path_web):
        global last_dir

        raptor = process_raptor(path_dk, path_gjc_inter, path_gjc_dom)
        web = process_report(path_web)

        # if web == 0:
        #     df_akhir = raptor
        # else:
        df_akhir = pd.concat([raptor, web])
        show(result_paned, df_akhir)

        try:
            save_filename = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                                        initialdir=last_dir,
                                                        title="Save",
                                                        filetypes=(('Microsoft Excel', "*.xlsx"), ("All Files", "*.*")))

            if save_filename:
                df_akhir.to_excel(save_filename, index=False)
            saved_success()
        except:
            saved_failed()

    result_paned = PanedWindow(bd=2, relief="groove")
    result_paned.place(x=350, y=15, width=440, height=575)
    hasil_frame = Frame(result_paned, width=425, height=550, background='#cccbca', relief="groove", border=1)
    hasil_frame.place(x=5, y=15)


    root.mainloop()

if __name__ == "__main__":
    main()