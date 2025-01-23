import tkinter as tk
from tkinter import filedialog

def select_file(x):
  file_path = filedialog.askopenfilename()
  if file_path and x=='Input':
    Input_file_label.config(text=f"Selected Input file path: {file_path}")
  elif file_path and x=='Output':
    Output_file_label.config(text=f"Selected Output file path: {file_path}")
def process():
  pass
root = tk.Tk()
root.title("Exel editor")

input_button = tk.Button(root, text="Input file path", command=lambda : select_file('Input'))
input_button.pack(pady=20,padx=200)
output_button = tk.Button(root, text="Output file path", command=lambda : select_file('Output'))
output_button.pack(pady=20,padx=200)

Input_file_label = tk.Label(root, text="Input file path :No file selected")
Input_file_label.pack(pady=20)
Output_file_label = tk.Label(root, text="Output file path :No file selected")
Output_file_label.pack(pady=20)
start_button = tk.Button(root, text="start_process", command=process)
start_button.pack(pady=20,padx=200)
root.mainloop()
