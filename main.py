import pandas as pd
import tkinter as tk
from tkinter import messagebox, simpledialog
from datetime import datetime

file_path = 'food_pass_system.xlsx'

def load_data():
    try:
        registration_df = pd.read_excel(file_path, sheet_name='Registrations')
    except:
        registration_df = pd.DataFrame(columns=['Person Name', 'House Number', 'Day 1 Passes', 'Day 2 Passes', 'Day 3 Passes', 'Day 4 Passes', 'Day 5 Passes', 'Total Passes Credited'])

    try:
        distribution_df = pd.read_excel(file_path, sheet_name='Distribution')
    except:
        distribution_df = pd.DataFrame(columns=['Person Name', 'House Number', 'Day 1 Passes Left', 'Day 2 Passes Left', 'Day 3 Passes Left', 'Day 4 Passes Left', 'Day 5 Passes Left', 'Date', 'Time'])

    try:
        summary_df = pd.read_excel(file_path, sheet_name='Summary')
        if summary_df.empty:
            summary_df = pd.DataFrame({
                'Day': ['Day 1', 'Day 2', 'Day 3', 'Day 4', 'Day 5'],
                'Total Passes Credited': [0, 0, 0, 0, 0],
                'Total Passes Left': [0, 0, 0, 0, 0],
            })
    except:
        summary_df = pd.DataFrame({
            'Day': ['Day 1', 'Day 2', 'Day 3', 'Day 4', 'Day 5'],
            'Total Passes Credited': [0, 0, 0, 0, 0],
            'Total Passes Left': [0, 0, 0, 0, 0],
        })

    return registration_df, distribution_df, summary_df

def save_data(registration_df, distribution_df, summary_df):
    with pd.ExcelWriter(file_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        registration_df.to_excel(writer, sheet_name='Registrations', index=False)
        distribution_df.to_excel(writer, sheet_name='Distribution', index=False)
        summary_df.to_excel(writer, sheet_name='Summary', index=False)

def credit_pass_gui():
    def calculate_and_credit():
        person_name = person_name_entry.get()
        house_number = house_number_entry.get()
        passes_per_day = [int(day_entries[i].get()) for i in range(5)]
        total_passes = sum(passes_per_day)
        prices = [price_entries[i] for i in range(5)]
        total_cost = sum([passes_per_day[i] * prices[i] for i in range(5)])

        registration_df, distribution_df, summary_df = load_data()

        new_registration = pd.DataFrame({
            'Person Name': [person_name], 
            'House Number': [house_number],
            'Day 1 Passes': [passes_per_day[0]], 
            'Day 2 Passes': [passes_per_day[1]], 
            'Day 3 Passes': [passes_per_day[2]], 
            'Day 4 Passes': [passes_per_day[3]], 
            'Day 5 Passes': [passes_per_day[4]],
            'Total Passes Credited': [total_passes]
        })
        registration_df = pd.concat([registration_df, new_registration], ignore_index=True)

        if person_name in distribution_df['Person Name'].values:
            for i in range(5):
                distribution_df.loc[distribution_df['Person Name'] == person_name, f'Day {i+1} Passes Left'] += passes_per_day[i]
        else:
            new_distribution = pd.DataFrame({
                'Person Name': [person_name], 
                'House Number': [house_number],
                'Day 1 Passes Left': [passes_per_day[0]], 
                'Day 2 Passes Left': [passes_per_day[1]], 
                'Day 3 Passes Left': [passes_per_day[2]], 
                'Day 4 Passes Left': [passes_per_day[3]], 
                'Day 5 Passes Left': [passes_per_day[4]],
                'Date': [""],
                'Time': [""]
            })
            distribution_df = pd.concat([distribution_df, new_distribution], ignore_index=True)
        
        for i in range(5):
            summary_df.loc[i, 'Total Passes Credited'] += passes_per_day[i]
            summary_df.loc[i, 'Total Passes Left'] += passes_per_day[i]

        save_data(registration_df, distribution_df, summary_df)
        credit_window.destroy()
        messagebox.showinfo("Success", f"Credited {total_passes} passes to {person_name} (House {house_number}). Total cost: ₹{total_cost}")

    credit_window = tk.Toplevel(root)
    credit_window.title("Credit Passes")
    credit_window.geometry("500x400")
    credit_window.configure(bg="#e6f2ff")
    
    for i in range(7):
        credit_window.grid_rowconfigure(i, weight=1)
    for i in range(4):
        credit_window.grid_columnconfigure(i, weight=1)

    tk.Label(credit_window, text="Person Name:", bg="#e6f2ff", font=("Helvetica", 12)).grid(row=0, column=0, padx=10, pady=10, sticky="w")
    person_name_entry = tk.Entry(credit_window, font=("Helvetica", 12))
    person_name_entry.grid(row=0, column=1, padx=10, pady=10, sticky="w")

    tk.Label(credit_window, text="House Number:", bg="#e6f2ff", font=("Helvetica", 12)).grid(row=1, column=0, padx=10, pady=10, sticky="w")
    house_number_entry = tk.Entry(credit_window, font=("Helvetica", 12))
    house_number_entry.grid(row=1, column=1, padx=10, pady=10, sticky="w")

    day_entries = []
    price_entries = [50, 60, 70, 80, 150]
    for i in range(5):
        tk.Label(credit_window, text=f"Day {i+1} Passes:", bg="#e6f2ff", font=("Helvetica", 12)).grid(row=i+2, column=0, padx=10, pady=10, sticky="w")
        day_entry = tk.Entry(credit_window, font=("Helvetica", 12))
        day_entry.grid(row=i+2, column=1, padx=10, pady=10, sticky="w")
        day_entries.append(day_entry)

        tk.Label(credit_window, text=f"Price per Pass (Day {i+1}):", bg="#e6f2ff", font=("Helvetica", 12)).grid(row=i+2, column=2, padx=10, pady=10, sticky="w")
        price_entry = tk.Entry(credit_window, font=("Helvetica", 12))
        price_entry.insert(0, price_entries[i])
        price_entry.config(state=tk.DISABLED)
        price_entry.grid(row=i+2, column=3, padx=10, pady=10, sticky="w")

    tk.Button(credit_window, text="Credit Passes", command=calculate_and_credit, font=("Helvetica", 12, "bold"), bg="#0073e6", fg="white").grid(row=7, column=1, columnspan=2, pady=20)

def debit_pass_gui():
    house_number = simpledialog.askstring("House Number", "Enter the house number:")
    
    if house_number:
        current_day = simpledialog.askinteger("Day", "Enter the day (1-5):")
        if current_day not in range(1, 6):
            messagebox.showwarning("Input Error", "Please provide a valid day (1-5).")
            return
        
        registration_df, distribution_df, summary_df = load_data()
        
        house_number = str(house_number).strip()
        
        person = distribution_df[distribution_df['House Number'].astype(str).str.strip() == house_number]

        if not person.empty:
            remaining_passes = person[f'Day {current_day} Passes Left'].values[0]
            if remaining_passes > 0:
                distribution_df.loc[distribution_df['House Number'].astype(str).str.strip() == house_number, f'Day {current_day} Passes Left'] -= 1
                summary_df.loc[current_day - 1, 'Total Passes Left'] -= 1
                
                now = datetime.now()
                current_time = now.strftime("%H:%M:%S")
                distribution_df.loc[distribution_df['House Number'].astype(str).str.strip() == house_number, 'Date'] = pd.to_datetime(now.date())
                distribution_df.loc[distribution_df['House Number'].astype(str).str.strip() == house_number, 'Time'] = current_time

                save_data(registration_df, distribution_df, summary_df)
                messagebox.showinfo("Success", f"Pass debited. {remaining_passes - 1} passes remaining for Day {current_day}.")
            else:
                messagebox.showwarning("No Passes Left", f"No passes remaining for Day {current_day}.")
        else:
            messagebox.showwarning("Person Not Found", "No entry found for this house number.")

# GUI Setup
root = tk.Tk()
root.title("Ganesh Mahotsav Pass System")

root.geometry("450x250")
root.configure(bg="#e6f2ff")

frame = tk.Frame(root, bg="#e6f2ff")
frame.pack(expand=True)

title_label = tk.Label(frame, text="Ganesh Mahotsav Pass System", font=("Helvetica", 18, "bold"), bg="#e6f2ff", fg="#333")
title_label.pack(pady=20)

credit_button = tk.Button(frame, text="Credit Pass", font=("Helvetica", 14, "bold"), width=20, command=credit_pass_gui, bg="#4CAF50", fg="white")
credit_button.pack(pady=10)

debit_button = tk.Button(frame, text="Debit Pass", font=("Helvetica", 14, "bold"), width=20, command=debit_pass_gui, bg="#f44336", fg="white")
debit_button.pack(pady=10)

root.mainloop()
