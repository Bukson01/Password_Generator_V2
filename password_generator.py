import random  # Import the random module to generate random values
import string  # Import the string module for working with strings
import tkinter as tk  # Import tkinter for creating GUI applications
from tkinter import messagebox  # Import messagebox from tkinter for displaying messages
from openpyxl import Workbook  # Import Workbook from openpyxl for working with Excel files
from datetime import datetime  # Import datetime from datetime for working with dates and times

def generate_password(length=12):
    """Generate a random password."""
    # Define characters to choose from
    characters = string.ascii_letters + string.digits + string.punctuation

    # Generate password using random characters
    password = ''.join(random.choice(characters) for _ in range(length))
    return password

def save_to_excel(password):
    """Save password to an Excel file."""
    # Create a new Excel workbook and select the active worksheet
    wb = Workbook()
    ws = wb.active
    
    # Add a header row and password data to the worksheet
    ws.append(["Generated Password", "Date & Time"])
    ws.append([password, datetime.now().strftime("%Y-%m-%d %H:%M:%S")])
    
    # Save the workbook to a file named "passwords.xlsx"
    wb.save("passwords.xlsx")

def generate_and_save_password():
    """Generate and save password."""
    # Get the password length from the user input
    password_length = int(length_entry.get())
    
    # Generate a password of the specified length
    generated_password = generate_password(password_length)
    
    # Save the generated password to an Excel file
    save_to_excel(generated_password)
    
    # Display a message box to notify the user that the password has been generated and saved
    messagebox.showinfo("Password Generated", "Password saved to passwords.xlsx")

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

# Start the GUI event loop
root.mainloop()