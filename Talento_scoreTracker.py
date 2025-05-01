# surname_scoreTracker.py
from openpyxl import Workbook, load_workbook
import tkinter as tk
from tkinter import messagebox

# Initialize Excel file
def init_excel():
    try:
        return load_workbook("student_scores.xlsx")
    except FileNotFoundError:
        wb = Workbook()
        ws = wb.active
        ws.append(["Student Name", "Score", "Status"])
        wb.save("student_scores.xlsx")
        return wb

# Create main window
window = tk.Tk()
window.title("Score Tracker")
window.geometry("300x200")

# Larger font and spacing
big_font = ("Arial", 12)
tk.Label(window, text="Student Name:", font=big_font).pack(pady=5)
name_entry = tk.Entry(window, font=big_font)
name_entry.pack()

tk.Label(window, text="Score (0-100):", font=big_font).pack(pady=5)
score_entry = tk.Entry(window, font=big_font)
score_entry.pack()

# Pass/Fail check
def calculate_status(score):
    return "Pass" if score >= 50 else "Fail"

# Save data with basic styling
def save_score():
    try:
        score = int(score_entry.get())
        if not 0 <= score <= 100:
            raise ValueError
    except ValueError:
        messagebox.showerror("Error", "Enter 0-100")
        return

    wb = init_excel()
    ws = wb.active
    ws.append([name_entry.get(), score, calculate_status(score)])
    wb.save("student_scores.xlsx")
    
    messagebox.showinfo("Saved", "Score saved!")
    name_entry.delete(0, tk.END)
    score_entry.delete(0, tk.END)

# View records with basic table
def view_records():
    wb = init_excel()
    ws = wb.active
    
    view = tk.Toplevel(window)
    view.title("Records")
    
    # Headers
    tk.Label(view, text="Name", relief="solid", width=15, font=big_font).grid(row=0, column=0)
    tk.Label(view, text="Score", relief="solid", width=10, font=big_font).grid(row=0, column=1)
    tk.Label(view, text="Status", relief="solid", width=10, font=big_font).grid(row=0, column=2)

    # Data rows
    for row_num, row in enumerate(ws.iter_rows(min_row=2, values_only=True), 1):
        for col, value in enumerate(row):
            tk.Label(view, text=value, relief="ridge", width=15 if col==0 else 10, 
                    bg="white" if row_num%2 else "#f0f0f0").grid(row=row_num, column=col)

# Buttons with color
tk.Button(window, text="Save Score", command=save_score, 
         bg="#4CAF50", fg="white", font=big_font).pack(pady=10)
tk.Button(window, text="View Records", command=view_records,
         bg="#2196F3", fg="white", font=big_font).pack()

window.mainloop()
