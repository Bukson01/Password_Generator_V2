# Import necessary modules
import random
import string
import tkinter as tk
from tkinter import messagebox
from openpyxl import Workbook
from datetime import datetime

def generate_password(length=12):
    """Generate a random password."""
    # Define characters to choose from
    characters = string.ascii_letters + string.digits + string.punctuation

    # Generate password using random characters
    password = ''.join(random.choice(characters) for _ in range(length))
    return password

# Function to save password to an Excel file
def save_to_excel(password):
    """Save password to an Excel file."""
    # Create a new workbook and select active worksheet
    wb = Workbook()
    ws = wb.active
    
    # Add header row and password data
    ws.append(["Generated Password", "Date & Time"])
    ws.append([password, datetime.now().strftime("%Y-%m-%d %H:%M:%S")])
    
    # Save workbook to a file named "passwords.xlsx"
    wb.save("passwords.xlsx")

# Function to generate and save password
def generate_and_save_password():
    """Generate and save password."""
    # Get password length from user input
    password_length_str = length_entry.get()
    
    # Check if length entry is empty
    if not password_length_str:
        messagebox.showerror("Error", "Please enter the length of the password.")
        return
    
    # Convert password length to integer
    try:
        password_length = int(password_length_str)
    except ValueError:
        messagebox.showerror("Error", "Invalid input. Please enter a valid integer for password length.")
        return
    
    
    # Generate password
    generated_password = generate_password(password_length)
    
    # Save password to Excel file
    save_to_excel(generated_password)
    
    # Update result label to display generated password
    result_label.config(text="Generated Password: " + generated_password)
    
    # Show message box to prompt user for saving or generating another password
    choice = messagebox.askyesnocancel("Save Password", "Do you want to save this password?")

    if choice is True:
        # Save password to Excel file
        save_to_excel(generated_password)
        # Show message box to notify user that password has been saved
        messagebox.showinfo("Password Saved", "Password saved to passwords.xlsx")
    elif choice is False:
        # Clear entry widget to allow user to generate another password
        length_entry.delete(0, tk.END)
    else:
        quit()#pass  # User cancelled operation

# Create GUI window
root = tk.Tk()
root.title("Password Generator")

# Create label and entry widgets for password length input
length_label = tk.Label(root, text="Password Length:")
length_label.pack()

length_entry = tk.Entry(root)
length_entry.pack()

# Create button to generate and save password
generate_button = tk.Button(root, text="Generate Password", command=generate_and_save_password)
generate_button.pack()

# Create label widget to display generated password
result_label = tk.Label(root, text="")
result_label.pack()

# Start GUI event loop
root.mainloop()
