import tkinter as tk
from tkinter import messagebox, filedialog
import os
import tempfile
import csv
import webbrowser

# ---------------- Paths ----------------
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
ICON_PATH = os.path.join(BASE_DIR, "icon.ico")
SPLASH_PATH = os.path.join(BASE_DIR, "splash.png")
EXCEL_FILE = os.path.join(BASE_DIR, "Student_Marks.csv")  # CSV acts as Excel

# ---------------- Splash Screen ----------------
def show_splash(root):
    splash = tk.Toplevel(root)
    splash.overrideredirect(True)
    splash.geometry("420x300+500+250")
    splash.configure(bg="#121212")

    try:
        splash.iconbitmap(ICON_PATH)
    except:
        pass

    try:
        splash_img = tk.PhotoImage(file=SPLASH_PATH)
        tk.Label(splash, image=splash_img, bg="#121212").pack()
        splash.image = splash_img
    except:
        tk.Label(
            splash,
            text="Student Grade Calculator",
            fg="white",
            bg="#121212",
            font=("Segoe UI", 16, "bold")
        ).pack(expand=True)

    splash.after(2500, splash.destroy)

# ---------------- Grade Logic ----------------
def calculate_grade(marks):
    if 90 <= marks <= 100:
        return "A", "Outstanding performance. Keep it up."
    elif 80 <= marks <= 89:
        return "B", "Great job. You are close to excellence."
    elif 70 <= marks <= 79:
        return "C", "Good effort. Keep improving."
    elif 60 <= marks <= 69:
        return "D", "You passed. Try harder next time."
    else:
        return "F", "Do not give up. Learn and try again."

# ---------------- Button Functions ----------------
def submit():
    name = name_entry.get().strip()
    name = name.capitalize()  # Capitalize first letter
    marks_input = marks_entry.get().strip()
    if not name:
        messagebox.showerror("Input Error", "Please enter student name.")
        return
    try:
        marks = int(marks_input)
        if not (0 <= marks <= 100):
            raise ValueError
    except ValueError:
        messagebox.showerror("Input Error", "Marks must be between 0 and 100.")
        return

    grade, message = calculate_grade(marks)
    result_label.config(
        text=(
            "STUDENT RESULT\n\n"
            f"Name   : {name}\n"
            f"Marks : {marks}\n"
            f"Grade : {grade}\n\n"
            f"{message}"
        )
    )

def reset_fields():
    name_entry.delete(0, tk.END)
    marks_entry.delete(0, tk.END)
    result_label.config(text="")
    name_entry.focus()

# ---------------- Download PDF (text-based) ----------------
def download_pdf():
    text = result_label.cget("text")
    if not text.strip():
        messagebox.showwarning("No Data", "No result to export.")
        return
    file_path = filedialog.asksaveasfilename(
        defaultextension=".pdf",
        filetypes=[("PDF Files", "*.pdf")]
    )
    if not file_path:
        return

    temp_txt = file_path.replace(".pdf", ".txt")
    with open(temp_txt, "w", encoding="utf-8") as f:
        f.write(text)
    messagebox.showinfo("Saved", f"PDF saved (text-based) at:\n{file_path}\nOpen to print in any PDF reader.")

# ---------------- Print PDF (opens default viewer) ----------------
def print_pdf():
    text = result_label.cget("text")
    if not text.strip():
        messagebox.showwarning("No Data", "No result to print.")
        return

    file_path = filedialog.asksaveasfilename(
        defaultextension=".pdf",
        filetypes=[("PDF Files", "*.pdf")]
    )
    if not file_path:
        return

    temp_txt = file_path.replace(".pdf", ".txt")
    with open(temp_txt, "w", encoding="utf-8") as f:
        f.write(text)

    webbrowser.open(temp_txt)

# ---------------- Save/Update Excel-like CSV ----------------
def save_to_excel():
    name = name_entry.get().strip()
    name = name.capitalize()  # Capitalize first letter
    marks_input = marks_entry.get().strip()
    
    if not name or not marks_input:
        messagebox.showwarning("Input Error", "Please enter student name and marks before saving.")
        return
    
    try:
        marks = int(marks_input)
        if not (0 <= marks <= 100):
            raise ValueError
    except ValueError:
        messagebox.showerror("Input Error", "Marks must be between 0 and 100.")
        return

    existing_data = []
    if os.path.isfile(EXCEL_FILE):
        with open(EXCEL_FILE, mode='r', newline='', encoding='utf-8') as file:
            reader = csv.reader(file)
            existing_data = list(reader)

    if not existing_data:
        existing_data.append(["Student Name", "Marks"])

    existing_data.append([name, marks])

    with open(EXCEL_FILE, mode='w', newline='', encoding='utf-8') as file:
        writer = csv.writer(file)
        writer.writerows(existing_data)

    messagebox.showinfo("Saved", f"Student data saved/updated in:\n{EXCEL_FILE}")

# ---------------- Main GUI ----------------
root = tk.Tk()
root.withdraw()
show_splash(root)
root.after(2600, root.deiconify)

root.title("Student Grade Calculator")
root.geometry("500x520")
root.configure(bg="#121212")
root.resizable(False, False)

try:
    root.iconbitmap(ICON_PATH)
except:
    pass

card = tk.Frame(root, bg="#1e1e1e", bd=2, relief="groove")
card.place(relx=0.5, rely=0.5, anchor="center", width=440, height=470)

tk.Label(card, text="Student Grade Calculator",
         bg="#1e1e1e", fg="white", font=("Segoe UI",16,"bold")).pack(pady=15)

tk.Label(card, text="Student Name", bg="#1e1e1e", fg="#b0bec5").pack(anchor="w", padx=30)
name_entry = tk.Entry(card, width=32, bg="#2c2c2c", fg="white", insertbackground="white")
name_entry.pack(pady=6)

tk.Label(card, text="Marks (0 - 100)", bg="#1e1e1e", fg="#b0bec5").pack(anchor="w", padx=30, pady=(10,0))
marks_entry = tk.Entry(card, width=32, bg="#2c2c2c", fg="white", insertbackground="white")
marks_entry.pack(pady=6)

# Buttons
btn_frame1 = tk.Frame(card, bg="#1e1e1e")
btn_frame1.pack(pady=15)

tk.Button(btn_frame1, text="Calculate", command=submit, bg="#2962ff", fg="white", width=12).grid(row=0,column=0,padx=8)
tk.Button(btn_frame1, text="Reset", command=reset_fields, bg="#d32f2f", fg="white", width=12).grid(row=0,column=1,padx=8)

btn_frame2 = tk.Frame(card, bg="#1e1e1e")
btn_frame2.pack(pady=10)

tk.Button(btn_frame2, text="Download PDF", command=download_pdf, bg="#388e3c", fg="white", width=18).grid(row=0,column=0,padx=8)
tk.Button(btn_frame2, text="Print PDF", command=print_pdf, bg="#6a1b9a", fg="white", width=18).grid(row=0,column=1,padx=8)
tk.Button(btn_frame2, text="Save to Excel", command=save_to_excel, bg="#009688", fg="white", width=18).grid(row=0,column=2,padx=8)

result_label = tk.Label(card, text="", bg="#1e1e1e", fg="#ecf0f1", font=("Segoe UI",11), justify="left")
result_label.pack(pady=15)

root.mainloop()
