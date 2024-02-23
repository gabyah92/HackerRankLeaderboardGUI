import threading
from tkinter import ttk, messagebox
import tkinter as tk
import tkinter.font as tkFont
import pandas as pd
import warnings
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
import requests

prog_text = ''


def generateExcelSheet(name, df):
    # Sort the DataFrame by 'Score' column in descending order
    df = df.sort_values(by='Score', ascending=False)
    # Create an Excel writer using openpyxl
    writer = pd.ExcelWriter(f'Leaderboards/{name}.xlsx', engine='openpyxl')
    df.to_excel(writer, index=False, sheet_name='Sheet1')
    worksheet = writer.sheets['Sheet1']
    # Define the cell styles
    font_header = Font(name='Arial', size=18, bold=True)
    font_body = Font(name='Arial', size=14, bold=True)
    fill_header = PatternFill(start_color='00ADEAEA', end_color='00ADEAEA', fill_type='solid')
    fill_body = PatternFill(start_color='00C7ECEC', end_color='00C7ECEC', fill_type='solid')
    align_center = Alignment(horizontal='center', vertical='center')
    border = Border(bottom=Side(style='medium'))
    df.insert(0, 'Rank', range(1, len(df) + 1))
    # Set the width of the 'Rank' column
    worksheet.column_dimensions['A'].width = 22
    worksheet.column_dimensions['B'].width = 30
    worksheet.column_dimensions['C'].width = 30
    # Apply the styling to the 'Rank' column
    for cell in worksheet['A']:
        cell.fill = fill_body
        cell.font = font_body
        cell.alignment = align_center
        cell.border = border
    # Apply formatting to the header row
    for col_num, value in enumerate(df.columns.values):
        worksheet.cell(row=1, column=col_num + 1).value = value
        worksheet.cell(row=1, column=col_num + 1).font = font_header
        worksheet.cell(row=1, column=col_num + 1).fill = fill_header
        worksheet.cell(row=1, column=col_num + 1).alignment = align_center
        worksheet.cell(row=1, column=col_num + 1).border = border
    # Apply formatting to the body cells
    for row_num, row in enumerate(df.values):
        for col_num, value in enumerate(row):
            worksheet.cell(row=row_num + 2, column=col_num + 1).value = value
            worksheet.cell(row=row_num + 2, column=col_num + 1).font = font_body
            worksheet.cell(row=row_num + 2, column=col_num + 1).fill = fill_body
            worksheet.cell(row=row_num + 2, column=col_num + 1).alignment = align_center
            worksheet.cell(row=row_num + 2, column=col_num + 1).border = border
    # Save the Excel file
    writer.close()

def generateCSVFile(name, df):
    # Sort the DataFrame by 'Score' column in descending order
    df = df.sort_values(by='Score', ascending=False)
    df.to_csv(f'LeaderboardsCSV/{name}.csv', index=False)

def getAll(tracker_names):
    try:
        global prog_text
        root.attributes('-disabled', True)
        progress_window = tk.Toplevel(root)
        progress_window.iconbitmap('venv/logo.ico')
        progress_window.title("Please Wait...")
        progress_window["borderwidth"] = "5px"
        progress_window["relief"] = "groove"
        progress_window.geometry("800x400")
        progress_window.resizable(False, False)
        progress_window['background'] = '#404445'
        progress_text = tk.Text(progress_window, height=30, width=80)
        progress_text['background'] = "grey"
        progress_text['fg'] = 'white'
        ft = tkFont.Font(family='Times', size=20, weight='bold')
        progress_text['font'] = ft
        print(f'progress text: {prog_text}'+ '\n')
        progress_text.config(state=tk.NORMAL)
        progress_text.insert(tk.END, prog_text+ '\n')
        progress_text.see(tk.END)
        progress_text.config(state=tk.DISABLED)
        progress_text.pack(pady=80)

        style = ttk.Style()
        style.theme_use('clam')  # Use the 'clam' theme as a base

        # Configure the style of the progress bar
        style.configure("TProgressbar",
                        thickness=20,  # Customize the thickness of the progress bar
                        troughcolor='lightgrey',  # Set the background color
                        background='#FF6C40',  # Set the color of the progress bar
                        )
        progress = ttk.Progressbar(progress_window, mode='determinate', style="TProgressbar")
        width = 800
        height = 400
        screenwidth = root.winfo_screenwidth()
        screenheight = root.winfo_screenheight()
        alignstr = '%dx%d+%d+%d' % (width, height, (screenwidth - width) / 2, (screenheight - height) / 2)
        progress_window.geometry(alignstr)
        progress.pack(padx=10, ipady=20)
        progress.place(x=50, y=10, width=700, height=50)
        def cleanup():
            root.attributes('-disabled', False)
            progress_window.wm_protocol(name='WM_DELETE_WINDOW')
            progress_window.destroy()
            root.state = 'normal'

        def generate_sheets_thread():
            global prog_text
            total_sheets = len(tracker_names)
            finished_sheets = 0
            progress_percent = 0
            warnings.filterwarnings('ignore')

            def check_close():
                nonlocal progress_window
                if progress_window.winfo_viewable():
                    progress_window.after(1000, check_close)
                else:
                    messagebox.showinfo("Process Interrupted", "Sheets generation process was interrupted.")
                    cleanup()
                    return
            d3 = dict()
            for tracker_name in tracker_names:
                data = []
                for offset in range(0, 1000, 100):
                    url = f'https://www.hackerrank.com/rest/contests/{tracker_name}/leaderboard?offset={offset}&limit=100'
                    headers = {
                        "User-agent": "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/47.0.2526.80 Safari/537.36"}
                    response = requests.get(url, headers=headers)
                    try:
                        response.raise_for_status()
                    except :
                        messagebox.showinfo("Process Interrupted", "Invalid URL || NO Internet!")
                        cleanup()
                        return
                    try:
                        json_data = response.json()
                    except:
                        print('Error'+str(response.status_code))
                        messagebox.showinfo("Invalid Response Code", "Something went Wrong1!.")
                        cleanup()
                        return
                    try :
                        if len(json_data['models']) == 0:
                            break
                    except :
                        break
                    for item in json_data['models']:
                        name = item['hacker']
                        score = item['score']
                        if name in d3:
                            d3.update( { name : d3.get(name)+score } )
                        else:
                            d3.update( {name : score} )
                        data.append({'Name': name, 'Score': score})

                try:
                    if len(data) == 0 :
                        messagebox.showinfo("Invalid Response Code", tracker_name + " was empty!")
                        continue
                    df = pd.DataFrame(data)
                    # dfA = pd.concat([df, dfA], ignore_index=True)
                    generateExcelSheet(tracker_name, df)
                except:
                    messagebox.showinfo("Invalid Data", "Something went Wrong2!.")
                    cleanup()
                    return

                finished_sheets += 1
                print(f'Finished : {tracker_name}\n')
                prog_text += f'\nFinished {tracker_name}!\n'
                print(f' Inside prog_text :  {prog_text}')
                progress_text.config(state=tk.NORMAL)
                progress_text.insert(tk.END, prog_text + '\n')
                progress_text.see(tk.END)
                progress_text.config(state=tk.DISABLED)
                progress_percent = int(finished_sheets / total_sheets * 100)
                progress['value'] = progress_percent
                progress_window.update()

            namesf = []
            scoref = []
            for i, uu in d3.items():
                namesf.append(i)
                scoref.append(float(uu))
            dfA = pd.DataFrame({'Name': namesf, 'Score': scoref})
            generateExcelSheet('TotalHackerrankLeaderBoard', dfA)
            if progress_window.winfo_viewable():
                if len(dfA) != 0:
                    messagebox.showinfo("Process Completed", "Sheets generated successfully.")
                cleanup()
                return
        threading.Thread(target=generate_sheets_thread).start()
        prog_text = ''
    except:
        messagebox.showinfo('Something Went Wrong!', 'This is Unexpected... Try again!')
        root.attributes('-disabled', False)
        root.state = 'normal'
        return

def on_closing():
    root.destroy()


def generate_sheets(ids):
    try:
        getAll(ids)
    except:
        return

def GButton_486_command():
    global prog_text
    inp = entry.get(1.0, 'end-1c')
    if inp == '   Enter Comma Separated values of HACKERRANK_CONTEST_ID\'s':
        messagebox.showerror('Error', 'Enter Something!')
        return
    try :
        inp = inp.replace(' ','').replace('\n','').split(',')
        d = list()
        prog_text += 'Generating leaderboards for : ('
        for i in inp:
            d.append(i)
            prog_text += str(i + ",")
        prog_text = prog_text[:-1]+ ')\n'
        print(d)
        generate_sheets(d)
    except :
        messagebox.showerror('Error', 'Something Went Wrong!')
        return
    else:
        import get_marks
        get_marks.main()



# Create the main window
root = tk.Tk()
root.title("Hackerrank Leaderboard")
root.configure(background='#404445')
width = 1142
height = 697
screenwidth = root.winfo_screenwidth()
screenheight = root.winfo_screenheight()
alignstr = '%dx%d+%d+%d' % (width, height, (screenwidth - width) / 2, (screenheight - height) / 2)
root.geometry(alignstr)
root.resizable(width=False, height=False)

# Create the Enter ID label
id_label = tk.Label(root)
id_label["anchor"] = "center"
ft = tkFont.Font(family='Helvetica', size=60, weight = 'bold')
id_label["font"] = ft
id_label["fg"] = "#FF6C40"
id_label["justify"] = "center"
id_label["text"] = "ENTER HACKERRANK ID'S!"
id_label['bg'] = '#404445'
id_label.place(x=15, y=2, width=1100, height=131)


def on_entry_click(event):
    if entry.get("1.0", 'end-1c') == '   Enter Comma Separated values of HACKERRANK_CONTEST_ID\'s':
        entry.delete('1.0', tk.END)


# Create the input field #FF6C40
entry = tk.Text(root)
entry["borderwidth"] = "5px"
entry['background'] = "black"
ft = tkFont.Font(family='Times', size=25, weight="bold")
entry["font"] = ft
entry.insert('1.0', '   Enter Comma Separated values of HACKERRANK_CONTEST_ID\'s')
entry.bind("<FocusIn>", on_entry_click)
entry["fg"] = "#FFE33E"
entry["relief"] = "groove"
entry.place(x=20, y=120, width=1101, height=431)
entry["insertbackground"] = "#FFE33E"


# Create the Generate button
def button_enter(event):
    generate_btn.config(background='black')


def button_leave(event):
    generate_btn.config(background='maroon')


generate_btn = tk.Button(root)
generate_btn.bind('<Enter>', button_enter)
generate_btn.bind('<Leave>', button_leave)
generate_btn["background"] = "maroon"
ft = tkFont.Font(family='Times', size=40, weight='bold')
generate_btn["borderwidth"] = "7px"
generate_btn["font"] = ft
generate_btn["fg"] = "#FFE33E"
generate_btn["justify"] = "center"
generate_btn["relief"] = "groove"
generate_btn["text"] = "Generate Excel Sheets!"
generate_btn["command"] = GButton_486_command
generate_btn.place(x=60, y=570, width=1010, height=99)
root.iconbitmap('venv/logo.ico')
root.wm_protocol(name='WM_DELETE_WINDOW')

# Start the tkinter event loop
root.mainloop()