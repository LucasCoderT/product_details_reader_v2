import subprocess, os, platform
import pathlib
import threading
import tkinter as tk
from tkinter import filedialog, ttk, messagebox

import main


def select_file(var):
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    var.set(file_path)

def open_file(file_path):
    output_file_path = pathlib.Path(file_path)
    if output_file_path.exists():
        if platform.system() == 'Darwin':       # macOS
            subprocess.call(('open', output_file_path))
        elif platform.system() == 'Windows':    # Windows
            os.startfile(output_file_path)
        else:                                   # linux variants
            subprocess.call(('xdg-open', output_file_path))

def worker(
        restock_report_path: pathlib.Path,
        inventory_file_path: pathlib.Path,
        feed_visor_processor_path: pathlib.Path,
):
    try:
        result = main.main(
            restock_report_path,
            inventory_file_path,
            feed_visor_processor_path,
            progress_bar=progress_bar
        )
        if result:
            messagebox.showinfo("Process Complete", "File processing is complete.")
            open_file(result)
            window.withdraw()
            window.destroy()
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {e}")
    finally:
        run_button.configure(state=tk.NORMAL)

def run_process():
    run_button.configure(state=tk.DISABLED)
    try:
        restock_report_path = pathlib.Path(file1_var.get())
        inventory_file_path = pathlib.Path(file2_var.get())
        feed_visor_processor_path = pathlib.Path(file3_var.get())

        if restock_report_path.exists() and inventory_file_path.exists() and feed_visor_processor_path.exists():
            # Add your process logic here
            threading.Thread(target=worker, daemon=True,kwargs={
                'restock_report_path': restock_report_path,
                'inventory_file_path': inventory_file_path,
                'feed_visor_processor_path': feed_visor_processor_path
            }).start()
        else:
            messagebox.showerror("File Not Found", "One or more files were not found. Please check the paths.")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {e}")


window = tk.Tk()
window.title("Product Details Processor")
window.resizable(False, False)
window.eval('tk::PlaceWindow . center')

file1_var = tk.StringVar()
file2_var = tk.StringVar()
file3_var = tk.StringVar()

label1 = tk.Label(window, text="Restock Report:")
label1.grid(row=0, column=0, padx=5, pady=5)
entry1 = tk.Entry(window, textvariable=file1_var, width=50)
entry1.grid(row=0, column=1, padx=5, pady=5)
button1 = tk.Button(window, text="Browse", command=lambda: select_file(file1_var))
button1.grid(row=0, column=2, padx=5, pady=5)

label2 = tk.Label(window, text="Inventory:")
label2.grid(row=1, column=0, padx=5, pady=5)
entry2 = tk.Entry(window, textvariable=file2_var, width=50)
entry2.grid(row=1, column=1, padx=5, pady=5)
button2 = tk.Button(window, text="Browse", command=lambda: select_file(file2_var))
button2.grid(row=1, column=2, padx=5, pady=5)

label3 = tk.Label(window, text="Feed visor:")
label3.grid(row=2, column=0, padx=5, pady=5)
entry3 = tk.Entry(window, textvariable=file3_var, width=50)
entry3.grid(row=2, column=1, padx=5, pady=5)
button3 = tk.Button(window, text="Browse.", command=lambda: select_file(file3_var))
button3.grid(row=2, column=2, padx=5, pady=5)

run_button = tk.Button(window, text="Run", command=run_process)
run_button.grid(row=3, column=1, padx=5, pady=10)

progress_bar = ttk.Progressbar(window, mode='determinate', maximum=10)
progress_bar.grid(row=4, column=0, columnspan=3, padx=5, pady=10, sticky='ew')


window.mainloop()