"""
This is a simple desktop application for vlookup,
concurrence check, and PBS coverage analysis.
"""
import tkinter as tk
from tkinter import ttk
import os
import pathlib
import numpy as np
from tkinter import filedialog as fd
from tkinter import messagebox as mbx
import pandas as pd
import datetime
from datetime import date, datetime
from datetime import timedelta
import calendar as cl
from tkcalendar import DateEntry
# from tqdm.tk import tqdm, trange
from babel import numbers
import openpyxl

class LookupPBS(tk.Tk):
    username = os.getlogin()  # Fetch username
    pbs_file = f'/Users/{username}/Documents/pbs_coverage.xlsx' #save exported pbs analysis in this directory

    # initialize the app GUI
    def __init__(self):
        super().__init__()

        self.geometry('900x600')
        self.title('VLookup and PBS Achievement Analysis')

        # create tool bar
        self.toolbar = tk.Frame(self,bg='#00b300')
        self.import_btn = tk.Button(self.toolbar, text='Import PBS file', command=self.import_file)
        self.import_btn.pack(side='left', padx=2, pady=2)

        self.PBS_Coverage_btn = tk.Button(self.toolbar, text='PBS Coverage',command=self.pbs_analysis)
        self.PBS_Coverage_btn.pack(side='left', padx=2, pady=2)

        # create load button
        self.load_btn = tk.Button(self.toolbar, text='Import Lookup files', command=self.open_file)
        self.load_btn.pack(side='left', padx=2, pady=2)
        # create date choosers
        self.enddate = tk.StringVar()
        self.Enddate_box = DateEntry(self.toolbar, width=16, background="magenta3", textvariable=self.enddate, foreground="white", bd=2)
        self.Enddate_box.pack(side='right', padx=2, pady=2)
        self.Enddate_lbl = tk.Label(self.toolbar, text='Enddate:')
        self.Enddate_lbl.pack(side='right', padx=1, pady=2)

        self.startdate = tk.StringVar()
        self.Startdate_box = DateEntry(self.toolbar, width=16, background="magenta3", textvariable=self.startdate, foreground="white", bd=2)
        self.Startdate_lbl = tk.Label(self.toolbar, text='Startdate(Optional):')
        self.Startdate_box.pack(side='right', padx=2, pady=2)
        self.Startdate_lbl.pack(side='right', padx=1, pady=2)

        self.toolbar.pack(side='top',fill='x')

        self.info_label = tk.Label(self, text='', fg='red')
        self.info_label.place(relx=0.03, rely=0.06)
        # create frame 1
        frame1 = tk.LabelFrame(self, text='Merge By')
        frame1.place(rely='0.1', relx='0.03', relwidth='0.21', relheight='0.75')

        # create frame 2
        frame2 = tk.LabelFrame(self, text='Merged without concurrence check')
        frame2.place(rely='0.1', relx='0.27', relwidth='0.21', relheight='0.75')

        # create frame 3
        frame3 = tk.LabelFrame(self, text='Merged with concurrence check')
        frame3.place(relx='0.61', rely='0.1', relwidth='0.21', relheight='0.75')

        # create treeview 1
        self.tree1 = ttk.Treeview(frame1)
        # create treeview 2
        self.tree2 = ttk.Treeview(frame2)
        # create treeview 3
        self.tree3 = ttk.Treeview(frame3)
        # self.tree3.place(relwidth='1', relheight='1')

        # create scroll bars1
        yscroll1 = tk.Scrollbar(frame1, orient='vertical', command=self.tree1.yview)
        xscroll1 = tk.Scrollbar(frame1, orient='horizontal', command=self.tree1.xview)
        self.tree1.configure(xscrollcommand=xscroll1.set, yscrollcommand=yscroll1.set)
        yscroll1.pack(side='right', fill='y')
        xscroll1.pack(side='bottom', fill='x')

        # create scroll bars2
        yscroll2 = tk.Scrollbar(frame2, orient='vertical', command=self.tree2.yview)
        xscroll2 = tk.Scrollbar(frame2, orient='horizontal', command=self.tree2.xview)
        self.tree2.configure(xscrollcommand=xscroll2.set, yscrollcommand=yscroll2.set)
        yscroll2.pack(side='right', fill='y')
        xscroll2.pack(side='bottom', fill='x')

        # create scroll bars3
        yscroll3 = tk.Scrollbar(frame3, orient='vertical', command=self.tree3.yview)
        xscroll3 = tk.Scrollbar(frame3, orient='horizontal', command=self.tree3.xview)
        self.tree3.configure(xscrollcommand=xscroll3.set, yscrollcommand=yscroll3.set)
        yscroll3.pack(side='right', fill='y')
        xscroll3.pack(side='bottom', fill='x')

        # create merge button
        merge_btn = tk.Button(self, text='Merge Files', fg='blue', command=self.merge_files)
        merge_btn.place(relx=0.03, rely=0.86)

        # create move buttons1
        up_btn1 = tk.Button(self, text='Up', fg='blue', command=self.move_up_1)
        up_btn1.place(relx=0.48, rely=0.33)
        down_btn1 = tk.Button(self, text='Down', fg='blue', command=self.move_down_1)
        down_btn1.place(relx=0.48, rely=0.43)
        # create move buttons2
        up_btn2 = tk.Button(self, text='Up', fg='blue', command=self.move_up_2)
        up_btn2.place(relx=0.82, rely=0.33)
        down_btn2 = tk.Button(self, text='Down', fg='blue', command=self.move_down_2)
        down_btn2.place(relx=0.82, rely=0.43)

        # create remove button1
        remove_btn1 = tk.Button(self, text='Remove', fg='blue', command=self.remove_1)
        remove_btn1.place(relx=0.48, rely=0.53)
        # create remove button2
        remove_btn1 = tk.Button(self, text='Remove', fg='blue', command=self.remove_2)
        remove_btn1.place(relx=0.82, rely=0.53)

        # create export button 1
        export_btn1 = tk.Button(self, text='Export this file', fg='blue', command=self.export_merged_file_1)
        export_btn1.place(relx=0.27, rely=0.86)

        # create export button 2
        export_btn2 = tk.Button(self, text='Export this file', fg='blue', command=self.export_merged_file_2)
        export_btn2.place(relx=0.61, rely=0.86)

    # save vlookup and concurrence files in this directory
    merged_file_1 = f'/Users/{username}/Documents/merged_file_1.xlsx'
    merged_file_2 = f'/Users/{username}/Documents/merged_file_2.xlsx'

    # method for vlookup file import
    def open_file(self):
        filename = fd.askopenfilename() #launch a file browser dialog window
        ext = pathlib.Path(filename).suffix #extracts file extension
        if ext != '.xlsx':
            mbx.showerror('Error!', 'Invalid file. Save the files in one .xlsx workbook,\n'
                                    'in Sheet1 and Sheet2, then try again')
        else:
            try:
                self.df1 = pd.read_excel(filename, sheet_name='Sheet1')
                self.df2 = pd.read_excel(filename, sheet_name='Sheet2')

                # clear old trees
                self.clear_trees()

                # set up new treeview
                self.tree1['column'] = 'columns'
                self.tree1['show'] = 'headings'
                self.tree1.heading(self.tree1['column'], text=self.tree1['column'])

                self.tree2['column'] = 'Merged_column_headers_1'
                self.tree2['show'] = 'headings'
                self.tree2.heading(self.tree2['column'], text=self.tree2['column'])

                self.tree3['column'] = 'Merged_column_headers_2'
                self.tree3['show'] = 'headings'
                self.tree3.heading(self.tree3['column'], text=self.tree3['column'])

                self.ws1_cols = list(self.df1.columns) #fetch column headers as a list

                # insert column headers into the treeview widget
                for column in self.ws1_cols:
                    self.tree1.insert("", "end", values=column)

                    # pack the two
                self.tree1.place(relwidth=1, relheight=0.98)
                self.tree2.place(relwidth=1, relheight=0.98)
                self.tree3.place(relwidth=1, relheight=0.98)
                self.info_label.configure(text="File loaded successfully!", fg='green')
            except ValueError:
                self.info_label.configure(text="File couldn't be opened... try again!")
            except FileNotFoundError:
                self.info_label.configure(text="File was not found... try again!")

        # clear trees
    def clear_trees(self):
        self.tree1.delete(*self.tree1.get_children())
        self.tree2.delete(*self.tree2.get_children())
        self.tree3.delete(*self.tree3.get_children())
    # method removes the seleted column(s)
    def remove_1(self):
        items = self.tree2.selection()
        for item in items:
            self.tree2.delete(item)

    def remove_2(self):
        items = self.tree3.selection()
        for item in items:
            self.tree3.delete(item)
    # move the the selected column(s)
    def move_up_1(self):
        items = self.tree2.selection()
        for item in items:
            self.tree2.move(item, self.tree2.parent(item), self.tree2.index(item) - 1)

    def move_up_2(self):
        items = self.tree3.selection()
        for item in items:
            self.tree3.move(item, self.tree3.parent(item), self.tree3.index(item) - 1)

    def move_down_1(self):
        items = self.tree2.selection()
        for item in reversed(items):
            self.tree2.move(item, self.tree2.parent(item), self.tree2.index(item) + 1)

    def move_down_2(self):
        items = self.tree3.selection()
        for item in reversed(items):
            self.tree3.move(item, self.tree3.parent(item), self.tree3.index(item) + 1)

    def merge_files(self):
        try:
            selected = self.tree1.focus() #fetch the selected column used for merging.
            merge_by = self.tree1.item(selected)['values']
            self.merged = self.df1.merge(self.df2, on=merge_by, how='left')

            # prepare normal vlookup section
            # fetch columns not common in both files and those in the lookup file
            needed_col = []
            for col in self.merged.columns:
                if col[-2:] not in ['_x', '_y'] or col[-2:] == '_y':
                    needed_col.append(col)
            self.needed_file = self.merged[needed_col] #filter merged file

            # fetch column namse with _y suffix removed
            striped_col = []
            for col in self.needed_file.columns:
                if col[-2:] != '_y':
                    striped_col.append(col)
                else:
                    striped_col.append(col[:-2])

            self.needed_file.columns = striped_col
            for column in self.needed_file.columns:
                self.tree2.insert("", "end", values=column)

            # prepare concurrence section
            modified_cols = []
            y_cols = []
            for col in self.merged.columns:
                if col[-2:] not in ['_x', '_y'] or col[-2:] == '_x': #column names that ends with _x or not with _y
                    modified_cols.append(col)
                else:
                    y_cols.append(col)
            for col in modified_cols:
                for coly in y_cols:
                    # compares each column's cell value in modified_cols with the corresponding column in y_cols
                    if col[-2:] == '_x' and col[:-2] == coly[:-2]:
                        xindex = modified_cols.index(col) #fetch the column position in the file
                        # insert column coly at xindex + 1 (juxtapose corresponding columns).
                        modified_cols.insert(xindex + 1, coly)

            self.allfile = self.merged[modified_cols]
            # check concurrence
            for col in self.allfile.columns:
                if col[-2:] == '_y':
                    idx = self.allfile.columns.get_loc(col) #get previous column name
                    val = self.allfile[col] #get previous column name value
                    precol = self.allfile.columns[idx - 1] #get next column name
                    preval = self.allfile[precol] #get next column name value
                    self.allfile.insert(loc=idx, column=f'sameValue{idx}',
                                        value=np.where(preval == val, 'True', 'False'))

            for column in self.allfile.columns:
                self.tree3.insert("", "end", values=column)

        except:
            self.info_label.configure(text='File error! Ensure you have imported a proper file, then select a unique column name existing in both files before merging')

    def export_merged_file_1(self):
        try:
            columns = []
            for line in self.tree2.get_children():
                for value in self.tree2.item(line)['values']:
                    columns.append(value)
            file = self.needed_file[columns]
            file.to_excel(LookupPBS.merged_file_1, index=False,engine='openpyxl')
            os.system(f'start excel.exe "{LookupPBS.merged_file_1}"')
        except:
            self.info_label.configure(text='File error! Ensure you imported the proper files, no spaces in column names, '
                                           'then merge with a unique column name existing in both files')

    def export_merged_file_2(self):
        try:
            columns = []
            for line in self.tree3.get_children():
                for value in self.tree3.item(line)['values']:
                    columns.append(value)
            file = self.allfile[columns]
            file.to_excel(LookupPBS.merged_file_2, index=False,engine='openpyxl')
            os.system(f'start excel.exe "{LookupPBS.merged_file_2}"')
        except:
            self.info_label.configure(text='File error! Ensure you imported the proper files, no spaces in column names, '
                                           'then merge with a unique column name existing in both files')
    # import pbs file
    def import_file(self):
        filename = fd.askopenfilename()
        ext = pathlib.Path(filename).suffix
        if ext != '.xlsx':
            mbx.showerror('Error!', 'Invalid file. Save the files as csv then try again')
        else:
            try:
                self.file_df = pd.read_excel(filename)
                mbx.showinfo('File Import', 'File successfully imported')
            except Exception as e:
                mbx.showerror('Import Error!', f'{e}')

    # error message displayed when no pbs linelist was imported
    def export_error_message(self):
        mbx.showerror('Export Error!', 'Import PBS linelist and try again!')

    # fetch selected dates
    def selected_dates(self):
        startdate = datetime.strptime(self.startdate.get(), '%m/%d/%y')
        enddate = datetime.strptime(self.enddate.get(),'%m/%d/%y')
        dates = (str(startdate)[:11], str(enddate)[:11])
        return dates

    def pbs_analysis(self):
        try:
            linelist = self.file_df
            linelist['ARTStartDate'] = pd.to_datetime(linelist['ARTStartDate'], format='%d/%m/%Y', exact=False)
            linelist['LastPickupDate'] = pd.to_datetime(linelist['LastPickupDate'], format='%d/%m/%Y', exact=False)
            linelist['DateOfCurrentVL'] = pd.to_datetime(linelist['DateOfCurrentVL'], format='%d/%m/%Y', exact=False)
            linelist['Last_VL_Sample_Date'] = pd.to_datetime(linelist['Last_VL_Sample_Date'],
                                                             format='%d/%m/%Y', exact=False)
            linelist['NextAppmt'] = pd.to_datetime(linelist['NextAppmt'], format='%d/%m/%Y', exact=False)
            linelist['CurrentVL'] = pd.to_numeric(linelist['CurrentVL'], errors='coerce')

            # filter tx_curr
            filter_tx_curr = linelist.loc[(linelist['CurrentARTStatus_28Days'] == 'Active')]
            tx_curr_frame = pd.DataFrame(filter_tx_curr)
            tx_curr_frame_filtered = tx_curr_frame[['FacilityName', 'CurrentARTStatus_28Days']].groupby('FacilityName').count()
            tx_curr_frame_filtered.columns = ['Tx_curr']

            # filter tx_curr with_pbs
            filter_tx_curr_with_pbs = linelist.loc[(linelist['CurrentARTStatus_28Days'] == 'Active')
                                                   & (linelist['Biometrics_Captured'] == 'Yes')]
            tx_curr_with_pbs_frame = pd.DataFrame(filter_tx_curr_with_pbs)
            tx_curr_with_pbs_frame_filtered = tx_curr_with_pbs_frame[['FacilityName', 'Biometrics_Captured']].fillna(0)
            tx_curr_with_pbs_frame_filtered = tx_curr_with_pbs_frame_filtered.groupby('FacilityName').count()
            tx_curr_with_pbs_frame_filtered.columns = ['Tx_curr with PBS']

            tx_curr_with_pbs_joined = tx_curr_frame_filtered.join(tx_curr_with_pbs_frame_filtered).fillna(0).astype(int)
            tx_curr_with_pbs_joined['% PBS Captured'] = (tx_curr_with_pbs_joined['Tx_curr with PBS'] /
                                                         tx_curr_with_pbs_joined['Tx_curr']) * 100
            tx_curr_with_pbs_joined['% PBS Captured'] = tx_curr_with_pbs_joined['% PBS Captured'].round(1)
            tx_curr_with_pbs_joined

            # filter tx_new for the selected tx_new_period
            tx_new_period = (linelist['ARTStartDate'] >= self.selected_dates()[0]) \
                     & (linelist['ARTStartDate'] <= self.selected_dates()[1])
            filter_artstartdate = linelist.loc[(tx_new_period)]
            artstartdate_frame = pd.DataFrame(filter_artstartdate)[['FacilityName', 'ARTStartDate']]
            artstartdate_frame.columns = ['FacilityName', 'Tx_New between {} & {}'.format(self.selected_dates()[0],
                                                                                          self.selected_dates()[1])]
            artstartdate_frame = artstartdate_frame.groupby('FacilityName').count()
            tx_curr_pbs_art_joined = tx_curr_with_pbs_joined.join(artstartdate_frame).fillna(0).astype(int)

            # filter pbs for tx_new
            filter_pbs_for_tx_new = linelist.loc[(tx_new_period) & (linelist['Biometrics_Captured'] == 'Yes')]
            pbs_for_tx_new_df = pd.DataFrame(filter_pbs_for_tx_new)[['FacilityName', 'Biometrics_Captured']]
            pbs_for_tx_new_df.columns = ['FacilityName','Total Tx_New Captured on PBS']
            pbs_for_tx_new_df = pbs_for_tx_new_df.groupby('FacilityName').count()
            tx_curr_tx_new_pbs_joined = tx_curr_pbs_art_joined.join(pbs_for_tx_new_df).fillna(0).astype(int)
            # filter pbs without pbs
            filter_tx_new_without_pbs = linelist.loc[(tx_new_period) & (linelist['Biometrics_Captured'] != 'Yes')]
            tx_new_without_pbs_df = pd.DataFrame(filter_tx_new_without_pbs)[['FacilityName', 'Biometrics_Captured']].fillna('No')
            tx_new_without_pbs_df.columns = ['FacilityName','Tx_New without PBS']
            tx_new_without_pbs_df = tx_new_without_pbs_df.groupby('FacilityName').count()
            tx_curr_tx_new_without_pbs_joined = tx_curr_tx_new_pbs_joined.join(tx_new_without_pbs_df).fillna(0).astype(int)

            # filter refill period
            refill_period = (linelist['LastPickupDate'] >= self.selected_dates()[0]) \
                     & (linelist['LastPickupDate'] <= self.selected_dates()[1])
            filter_refills_for_the_period = linelist.loc[(refill_period)]
            refills_for_the_period_df = pd.DataFrame(filter_refills_for_the_period)[['FacilityName','Biometrics_Captured']].fillna('Yes')
            refills_for_the_period_df.columns = ['FacilityName','Total refill']
            refills_for_the_period_df = refills_for_the_period_df.groupby('FacilityName').count()
            tx_curr_tx_new_refill_joined = tx_curr_tx_new_without_pbs_joined.join(refills_for_the_period_df).fillna(0).astype(int)
            # filter refills with pbs
            filter_refills_with_pbs = linelist.loc[(refill_period) & (linelist['Biometrics_Captured'] == 'Yes')]
            refills_with_pbs_df = pd.DataFrame(filter_refills_with_pbs)[['FacilityName','Biometrics_Captured']]
            refills_with_pbs_df.columns = ['FacilityName','Total refill with PBS']
            refills_with_pbs_df = refills_with_pbs_df.groupby('FacilityName').count()
            tx_curr_tx_new_refill_with_pbs_joined = tx_curr_tx_new_refill_joined.join(refills_with_pbs_df).fillna(0).astype(int)

            filter_refills_without_pbs = linelist.loc[(refill_period) & (linelist['Biometrics_Captured'] != 'Yes')]
            refills_without_pbs_df = pd.DataFrame(filter_refills_without_pbs)[['FacilityName','Biometrics_Captured']].fillna('No')
            refills_without_pbs_df.columns = ['FacilityName', 'refill with no PBS']
            refills_without_pbs_df = refills_without_pbs_df.groupby('FacilityName').count()
            tx_curr_tx_new_refill_without_pbs_joined = tx_curr_tx_new_refill_with_pbs_joined.join(refills_without_pbs_df).fillna(0).astype(int)
            tx_curr_tx_new_refill_without_pbs_joined['Total Missed opportunity'] = (tx_curr_tx_new_refill_without_pbs_joined['refill with no PBS'] + tx_curr_tx_new_refill_without_pbs_joined['Tx_New without PBS'])
            tx_curr_tx_new_refill_without_pbs_joined.to_excel(LookupPBS.pbs_file,engine='openpyxl')
            os.system(f'start excel.exe "{LookupPBS.pbs_file}"')
        except Exception as e:
            mbx.showerror('Import Error!', f'{e}')
# run app
if __name__ == '__main__':
    app = LookupPBS()
    app.mainloop()





