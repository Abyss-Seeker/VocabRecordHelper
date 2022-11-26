# encoding: UTF-8
# Created by Abyss_Seeker!
# Modules used: built-in: [tkinter, os, json]; others: [requests, openpyxl]
# Glhf

from Translator_crawler import search
import tkinter as tk
from Recorder import add_line, del_line


# From translate_button
def translate(word):
    # Using search function, get the translation
    translation = search(word)
    # Editing output textbox
    output_text.delete('1.0', 'end')
    output_text.insert('1.0', translation)
    # Modify event log label
    # 此处'end-1c'是为了去掉文本框返回时自带的一个\n
    if output_text.get('1.0','end-1c') != 'Word not found. Make sure it is in present tense and not spelled wrong! Try again :)':
        action_label['text'] = 'Translation for word "{}" successly found and presented'.format(word)
    else:
        action_label['text'] = 'Translation for word "{}" not found'.format(word)

# From record_button
def record(word, translation):
    row_num = add_line(word, translation)
    action_label['text'] = 'Row {} added: {} - {}'.format(row_num, word, translation)

def delete():
    line_info = del_line()
    if line_info == 'File doesn\'t exist' or line_info == 'No lines left':
        action_label['text'] = 'Delete line failed: {}'.format(line_info)
    else:
        action_label['text'] = 'Row {} successfully deleted: {} - {}'.format(line_info['row_num'], line_info['word'], line_info['translation'])


# 说明窗口
def help_window():
    help_root = tk.Tk()
    help_root.title('How to use Vocab Record Helper')
    help_label = tk.Label(help_root, text = '''=====Help=====

Translate - shows Chinese translation of word down in the text box

Record - records word + translation down in an excel under the same folder as the application

Delete - deletes the last word recorded in the excel

*Please make sure that while using record and delete, your excel file isn't open
*Translate function requires wifi connection
    ''')
    help_label.pack()

# Show help window
help_window()

# User interface
root = tk.Tk()
root.configure(bg='#272935')
root.geometry('500x450+200+200')
root.resizable(False, False)
root.title('Vocab Record Helper - User Interface')

# Setting title
title_label = tk.Label(root, text='Vocab Record Helper', fg='white', bg='#7B7B7D', font=('Arial',18), height=2, width=40)
title_label.place(x=-30,y=0)

# Initializing input label and entry box
input_label = tk.Label(root, text='Please enter the vocab:', fg='white', bg='#272935', font=('Arial', 14))
input_label.place(x=20, y=80)
input_entry = tk.Entry(root, fg='white', bg='#2B424A', font=('Arial', 14), width=18)
input_entry.insert(0, 'Example')
input_entry.place(x=250, y=80)

# Initializing output label and text box
output_label = tk.Label(root, text='Translation:', fg='white', bg='#272935', font=('Arial', 14))
output_label.place(x=20, y=120)
output_text = tk.Text(root, fg='white', bg='#2B424A', font=('Arial', 14), width=41, height=6)
output_text.insert('1.0','The translation will appear here.')
output_text.place(x=20, y=160)

# Initializing left button, when button clicked, function "translate" will be triggered
translate_button = tk.Button(root, text='Translate', fg='white', bg='#E6BE70', font=('Arial', 14, 'bold'), width = 10, height = 2,
                   command=lambda: translate(input_entry.get()))
translate_button.place(x=20, y=320)

# Initializing middle button, when button clicked, function "record" will be triggered
record_button = tk.Button(root, text='Record', fg='white', bg='#EDA14B', font=('Arial', 14, 'bold'), width = 10, height = 2,
                   command=lambda: record(input_entry.get().lower(), output_text.get('1.0','end-1c')))
record_button.place(x=180, y=320)

# Initializing right button, when button clicked, function "delete" will be triggered
record_button = tk.Button(root, text='Delete', fg='white', bg='#A14939', font=('Arial', 14, 'bold'), width = 10, height = 2,
                   command=delete)
record_button.place(x=340, y=320)

# Add small comments
cmt_label = tk.Label(root, text='*While using Record and Delete, make sure the excel file is closed!', fg='white', bg='#272935', font=('Arial', 10))
cmt_label.place(x=20, y=385)

name_label = tk.Label(root, text='@Abyss_Seeker!', fg='#BAFFBA', bg='#272935', font=('Arial', 6))
name_label.place(x=440,y=435)

# Initializing event log (title and log) at bottom
event_log_label = tk.Label(root, text='Event log: Recent actions will appear below:', fg='white', bg='#272935', font=('Arial', 8))
event_log_label.place(x=20, y=410)
action_label = tk.Label(root, text='No events yet', fg='white', bg='#272935', font=('Arial', 8))
action_label.place(x=20, y=428)

# Help button, shows help window
record_button = tk.Button(root, text='?', fg='#E3C515', bg='#7B7B7D', font=('Arial', 8, 'bold'), width = 2, height = 1,
                   command=help_window)
record_button.place(x=475, y=0)

# Let's rock
root.mainloop()


