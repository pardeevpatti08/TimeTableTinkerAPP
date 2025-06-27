import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import simpledialog
from tkcalendar import Calendar
from datetime import datetime, timedelta
import pandas as pd

# Initialize global variables
holidays = []  # Store the list of holidays
holiday_dict = {}  # Map holiday dates to their occasions
course_timetable = pd.DataFrame()  # Store the course timetable dataframe

# Function to load holidays from Excel
def load_holidays_from_excel():
    try:
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if file_path:
            df = pd.read_excel(file_path)
            if 'Holiday' in df.columns and 'Occasion' in df.columns:
                global holidays, holiday_dict
                holidays = []
                holiday_dict = {}
                
                for _, row in df.dropna(subset=['Holiday']).iterrows():
                    holiday_date = None
                    occasion = row['Occasion']
                    date = row['Holiday']

                    if isinstance(date, str):
                        cleaned_date = date.strip().lower()
                        holiday_date = datetime.strptime(cleaned_date, "%d.%m.%Y")
                    elif isinstance(date, datetime):
                        holiday_date = date
                    
                    if holiday_date:
                        holidays.append(holiday_date)
                        holiday_dict[holiday_date] = occasion

                update_holiday_listbox()
                messagebox.showinfo("Success", "Holidays loaded successfully!")
            else:
                messagebox.showerror("Error", "The selected file does not contain 'Holiday' or 'Occasion' columns.")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {str(e)}")

# Update the holiday listbox to show both date and occasion
def update_holiday_listbox():
    holiday_listbox.delete(0, tk.END)  # Clear the current listbox
    for holiday in holidays:
        occasion = holiday_dict.get(holiday, 'Manual Holiday')  # Get occasion or default to 'Manual Holiday'
        display_text = f"{holiday.strftime('%d-%m-%Y')} - {occasion}"
        holiday_listbox.insert(tk.END, display_text)  # Insert formatted text into listbox

# Function to load course timetable from an Excel file
def load_course_timetable():
    try:
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if file_path:
            global course_timetable
            course_timetable = pd.read_excel(file_path)
            messagebox.showinfo("Success", "Course timetable loaded successfully!")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {str(e)}")

# Generate and save class dates with holidays included
def generate_class_dates():
    try:
        start_dt = datetime.strptime(start_entry.get().strip(), '%Y-%m-%d')
        end_dt = datetime.strptime(end_entry.get().strip(), '%Y-%m-%d')
        selected_weekdays = weekdays_listbox.curselection()

        if not selected_weekdays:
            messagebox.showwarning("Input Error", "Please select at least one weekday.")
            return

        # Generate class dates based on selected weekdays and excluding holidays
        class_dates = []
        current_date = start_dt
        while current_date <= end_dt:
            if current_date.weekday() in selected_weekdays:
                class_dates.append(current_date)
            current_date += timedelta(days=1)

        if class_dates:
            # Create a DataFrame with class dates
            class_dates_df = pd.DataFrame({'Class Dates': class_dates})

            # Add topics to the class dates
            if not course_timetable.empty:
                topics = course_timetable['Topic'].tolist()
                num_topics = len(topics)

                repeated_topics = (topics * ((len(class_dates) // num_topics) + 1))[:len(class_dates)]
                class_dates_df['Topic'] = repeated_topics
            else:
                class_dates_df['Topic'] = None

            # Add a new column for holidays and occasions
            class_dates_df['Holiday Occasions'] = class_dates_df['Class Dates'].apply(
                lambda x: holiday_dict.get(x, None)  # Get the occasion if the date is a holiday
            )

            # Fill in any gaps with the next available topic (if a day is a holiday, assign only the occasion)
            current_topic_index = 0
            for i, row in class_dates_df.iterrows():
                if row['Holiday Occasions']:
                    # If it's a holiday, only append the occasion and leave the topic blank
                    occasion = row['Holiday Occasions']
                    class_dates_df.at[i, 'Topic'] = f"No Class - {occasion}"  # Only include the holiday occasion
                else:
                    # If it's not a holiday and we still have topics left, assign the next topic
                    if current_topic_index < len(topics):  # Ensure topics stop when we run out of them
                        class_dates_df.at[i, 'Topic'] = topics[current_topic_index]
                        current_topic_index += 1
                    else:
                        # If there are no more topics available, leave the topic empty
                        class_dates_df.at[i, 'Topic'] = ""

            # Ask user where to save the file
            file_path = filedialog.asksaveasfilename(defaultextension='.xlsx', filetypes=[('Excel files', '*.xlsx')])
            if file_path:
                # Format dates to dd-mm-yyyy
                class_dates_df['Class Dates'] = class_dates_df['Class Dates'].apply(lambda x: x.strftime('%d-%m-%Y'))

                # Remove the 'Holiday Occasions' column before saving to Excel
                class_dates_df.drop(columns=['Holiday Occasions'], inplace=True)

                # Save to Excel using openpyxl (default engine)
                class_dates_df.to_excel(file_path, index=False, sheet_name='Timetable')

                messagebox.showinfo("Success", f"Class dates with holidays saved to {file_path}")
        else:
            messagebox.showinfo("No Data", "No valid class dates were generated.")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred while generating class dates: {str(e)}")


# Add selected holiday or non-holiday to the holiday list with the occasion
def add_selected_holiday():
    selected_date = cal.get_date()  # Get the selected date from the calendar
    try:
        holiday_date = datetime.strptime(selected_date, "%m/%d/%y")  # Adjust the format to match calendar
        
        # Ask the user to select if it's a Holiday or Non-Holiday
        selected_occasion_type = occasion_type_combobox.get()

        # Ask the user for the occasion details
        occasion = simpledialog.askstring("Holiday Occasion", f"Enter the occasion for {holiday_date.strftime('%d-%m-%Y')}:")

        if occasion:
            if selected_occasion_type == "Holiday":
                # Add the date and occasion to the holiday list if it is a Holiday
                if holiday_date not in holidays:
                    holidays.append(holiday_date)
                    holiday_dict[holiday_date] = occasion
                    update_holiday_listbox()
                    messagebox.showinfo("Success", f"{holiday_date.strftime('%d-%m-%Y')} added as Holiday with occasion '{occasion}'.")
                else:
                    messagebox.showinfo("Info", f"{holiday_date.strftime('%d-%m-%Y')} is already in the holiday list.")
            elif selected_occasion_type == "Non-Holiday":
                # Add the date and occasion to the non-holiday list if it is a Non-Holiday
                if holiday_date not in holidays:
                    holidays.append(holiday_date)
                    holiday_dict[holiday_date] = f"Non-Holiday - {occasion}"
                    update_holiday_listbox()
                    messagebox.showinfo("Success", f"{holiday_date.strftime('%d-%m-%Y')} added as Non-Holiday with occasion '{occasion}'.")
                else:
                    messagebox.showinfo("Info", f"{holiday_date.strftime('%d-%m-%Y')} is already in the holiday list.")
        else:
            messagebox.showwarning("Input Error", "Please provide an occasion for the selected date.")
        
    except ValueError as e:
        messagebox.showerror("Error", f"Invalid date format: {e}")

# Delete selected holiday from the list
def delete_selected_holiday():
    try:
        selected_index = holiday_listbox.curselection()  # Get selected item from the listbox
        if not selected_index:
            messagebox.showwarning("Selection Error", "Please select a holiday to delete.")
            return

        # Get the date of the selected holiday from the listbox
        selected_text = holiday_listbox.get(selected_index)
        holiday_date_str = selected_text.split(" - ")[0]  # Extract the date from the string
        holiday_date = datetime.strptime(holiday_date_str, "%d-%m-%Y")

        # Confirm deletion
        confirm = messagebox.askyesno("Confirm Deletion", f"Are you sure you want to delete {holiday_date.strftime('%d-%m-%Y')}?")
        if confirm:
            holidays.remove(holiday_date)  # Remove from the holiday list
            del holiday_dict[holiday_date]  # Remove from the dictionary
            update_holiday_listbox()  # Update the listbox
            messagebox.showinfo("Success", f"{holiday_date.strftime('%d-%m-%Y')} has been deleted.")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {str(e)}")

# UI Setup
root = tk.Tk()
root.title("Timetable Calculator")
root.geometry("1100x600")
root.configure(bg="#e6f7ff")

main_frame = tk.Frame(root, bg="#e6f7ff")
main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

# Left Frame
left_frame = tk.Frame(main_frame, bg="#e6f7ff")
left_frame.grid(row=0, column=0, padx=20, pady=10, sticky="n")

# Right Frame
right_frame = tk.Frame(main_frame, bg="#e6f7ff")
right_frame.grid(row=0, column=1, padx=20, pady=10, sticky="n")

# File Loading Buttons
tk.Label(left_frame, text="Load Data:", font=("Helvetica", 16, "underline"), bg="#e6f7ff").pack(pady=5)
tk.Button(left_frame, text="Load Holidays", command=load_holidays_from_excel, bg="#4CAF50", fg="white", font=("Helvetica", 12)).pack(pady=5)
tk.Button(left_frame, text="Load Course Timetable", command=load_course_timetable, bg="#4CAF50", fg="white", font=("Helvetica", 12)).pack(pady=5)

# Calendar in Right Frame
cal = Calendar(right_frame, selectmode='day', year=2025, month=1, day=1)
cal.pack(pady=10)

from tkinter import ttk  # Add this import for ttk

# Add a label and combobox for selecting "Holiday" or "Non-Holiday" in the Right Frame
tk.Label(right_frame, text="Select Occasion Type:", font=("Helvetica", 12), bg="#e6f7ff").pack(pady=5)

occasion_type_combobox = ttk.Combobox(right_frame, values=["Holiday", "Non-Holiday"], state="readonly", font=("Helvetica", 12))
occasion_type_combobox.set("Holiday")  # Set the default value to "Holiday"
occasion_type_combobox.pack(pady=10)

# Holiday Listbox
holiday_listbox = tk.Listbox(right_frame, height=10, font=("Helvetica", 12), width=40)
holiday_listbox.pack(pady=5)

# Weekday Selection
weekdays_listbox = tk.Listbox(left_frame, selectmode=tk.MULTIPLE, font=("Helvetica", 12), height=7)
for i, day in enumerate(["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]):
    weekdays_listbox.insert(tk.END, day)
weekdays_listbox.pack(pady=5)

# Start Date Entry
tk.Label(left_frame, text="Start Date (YYYY-MM-DD):", font=("Helvetica", 12), bg="#e6f7ff").pack(pady=5)
start_entry = tk.Entry(left_frame, font=("Helvetica", 12))
start_entry.pack(pady=5)

# End Date Entry
tk.Label(left_frame, text="End Date (YYYY-MM-DD):", font=("Helvetica", 12), bg="#e6f7ff").pack(pady=5)
end_entry = tk.Entry(left_frame, font=("Helvetica", 12))
end_entry.pack(pady=5)

# Generate Timetable Button
generate_btn = tk.Button(left_frame, text="Generate Timetable", command=generate_class_dates, bg="#008CBA", fg="white", font=("Helvetica", 16, "bold"))
generate_btn.pack(pady=20)

# Add selected holiday button
add_holiday_btn = tk.Button(right_frame, text="Add Holiday to List", command=add_selected_holiday, bg="#FF6347", fg="white", font=("Helvetica", 12))
add_holiday_btn.pack(pady=10)

# Delete selected holiday button
delete_holiday_btn = tk.Button(right_frame, text="Delete Selected Holiday", command=delete_selected_holiday, bg="#FF6347", fg="white", font=("Helvetica", 12))
delete_holiday_btn.pack(pady=10)

root.mainloop()
