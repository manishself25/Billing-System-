#!/usr/bin/env python
# coding: utf-8

# In[ ]:


import pandas as pd
import tkinter
from tkinter import ttk
from tkinter import messagebox
from docxtpl import DocxTemplate
import datetime
import random
from docx2pdf import convert

window = tkinter.Tk()
window.title("Billing System")
window.geometry('1024x768')  # Initial window size
window.configure(bg='#f0f8ff')  # Light background color

# Configure grid weights to make the window resizable
window.columnconfigure(0, weight=1)
window.rowconfigure([1, 2], weight=1)

# Function to update time dynamically
def update_time():
    current_time = datetime.datetime.now().strftime("%d-%m-%Y %H:%M:%S")
    time_label.config(text=current_time)
    window.after(1000, update_time)  # Update time every 1 second

# Top frame for Welcome Message and Date/Time
top_frame = tkinter.Frame(window, bg='#4682B4', padx=10, pady=10)
top_frame.grid(row=0, column=0, sticky="ew")

welcome_label = tkinter.Label(top_frame, text="Welcome to Billing System", bg='#4682B4', fg='white', font=("Arial", 16, "bold"))
welcome_label.pack(side="left", padx=10, pady=5)

time_label = tkinter.Label(top_frame, text="", bg='#4682B4', fg='white', font=("Arial", 14))
time_label.pack(side="right", padx=10, pady=5)

# Start the time update
update_time()

# Customer Information Frame
customer_frame = tkinter.LabelFrame(window, text="Customer Information", padx=10, pady=10, bg='#f0f8ff', font=("Arial", 12, "bold"))
customer_frame.grid(row=1, column=0, padx=20, pady=10, sticky="ew")

# Configure grid columns for customer information
customer_frame.columnconfigure([0, 1, 2], weight=1)

first_name_label = tkinter.Label(customer_frame, text="Name", bg='#f0f8ff', font=("Arial", 10, "bold"))
first_name_label.grid(row=0, column=0, padx=10, pady=5, sticky="w")
first_name_entry = tkinter.Entry(customer_frame)
first_name_entry.grid(row=1, column=0, padx=10, pady=5, sticky="ew")

phone_number_label = tkinter.Label(customer_frame, text="Phone No.", bg='#f0f8ff', font=("Arial", 10, "bold"))
phone_number_label.grid(row=0, column=1, padx=10, pady=5, sticky="w")
phone_number_entry = tkinter.Entry(customer_frame)
phone_number_entry.grid(row=1, column=1, padx=10, pady=5, sticky="ew")

billlabel = tkinter.Label(customer_frame, text="Billing Number", bg='#f0f8ff', font=("Arial", 10, "bold"))
billlabel.grid(row=0, column=2, padx=10, pady=5, sticky="w")
billentry = tkinter.Entry(customer_frame)
billentry.grid(row=1, column=2, padx=10, pady=5, sticky="ew")

# Menu Frame
menu_frame = tkinter.LabelFrame(window, text="Menu", padx=10, pady=10, bg='#f0f8ff', font=("Arial", 12, "bold"))
menu_frame.grid(row=2, column=0, padx=20, pady=10, sticky="nsew")

# Configure grid in menu frame
menu_frame.columnconfigure([0, 1, 2], weight=1)

################ item dict ##########################################

df = pd.read_excel('billing_items.xlsx')
# Set the "Item Name" column as the index
df.set_index('Item Name', inplace=True)
# Convert the DataFrame to a dictionary with index orientation
item_dict = df.to_dict(orient='index')
# Extract the prices from the inner dictionaries
for item, data in item_dict.items():
    item_dict[item] = data['Price']  # Assuming the price column is named 'Price'

billnumber = random.randint(1000, 9999)

def clear_item():
    qty_spinbox_main.delete(0, tkinter.END)
    qty_spinbox_main.insert(0, "1")
    desc_label_box_main.set('')

invoice_list = []

def update_total():
    subtotal = sum(item[3] for item in invoice_list)
    sgst_rate = 2.5 / 100  # 2.5%
    cgst_rate = 2.5 / 100  # 2.5%
    
    # Calculate SGST and CGST amounts based on subtotal
    sgst_amount = subtotal * sgst_rate
    cgst_amount = subtotal * cgst_rate
    total_amount = subtotal + sgst_amount + cgst_amount

    # Update the labels to display subtotal, SGST, CGST, and total
    subtotal_label.config(text=f"Subtotal: {subtotal:.2f} INR")
    sgst_label.config(text=f"SGST (2.5%): {sgst_amount:.2f} INR")
    cgst_label.config(text=f"CGST (2.5%): {cgst_amount:.2f} INR")
    total_label.config(text=f"Total: {total_amount:.2f} INR")

def add_item():
    global item_dict
    billentry.delete(0, "end")
    billentry.insert(0, billnumber)
    if first_name_entry.get() == "" or phone_number_entry.get() == "":
        messagebox.showerror("Error", "Customer Details are Required")
        return
    qty = int(qty_spinbox_main.get())
    desc = desc_label_box_main.get()
    if desc in item_dict:
        price = item_dict[desc]
    else:
        messagebox.showerror("Error", "Please add some item from the dropdown menu")
        return
    line_total = qty * price
    invoice_item = [desc, qty, price, line_total]
    tree.insert('', 0, values=invoice_item)
    clear_item()
    invoice_list.append(invoice_item)
    update_total()  # Update total after adding item

def new_invoice():
    global billnumber
    billnumber = random.randint(1000, 9999)
    first_name_entry.delete(0, tkinter.END)
    phone_number_entry.delete(0, tkinter.END)
    billentry.delete(0, tkinter.END)
    clear_item()
    tree.delete(*tree.get_children())
    invoice_list.clear()
    update_total()  # Reset total to 0
    

def generate_invoice():
    now = datetime.datetime.now()
    year = now.year
    month = now.month
    day = now.day
    hour = now.hour
    minute = now.minute
    date= f"{day}_{month}_{year}"
    time= f"{hour}_{minute}"
    
    doc = DocxTemplate("Billing_Information_template.docx")
    name = first_name_entry.get()
    phone = phone_number_entry.get()
    subtotal = sum(item[3] for item in invoice_list)
    sgst_tax = subtotal*0.025
    cgst_tax = subtotal*0.025
    total = subtotal + sgst_tax + cgst_tax
    doc.render({"billno" : billnumber,
                "dt" : date,
                "name": name,
                "phoneno": phone,
                "invoice_list": invoice_list,
                "sgst" : sgst_tax,
                "cgst" : cgst_tax,
                "subtotal": subtotal,
                "finaltotal": total})
    

    doc_name = name + " " + date + "_" + time + ".docx"
    doc.save(doc_name)
    pdf_name = name + " " + date + "_" + time + ".docx"
    convert(pdf_name)
#     convert(pdf_name)
    
    messagebox.showinfo("Success, Invoice generated and saved")
    
#     #################### save to excel ############
    
    df = pd.read_excel('billing_details.xlsx')
    serialNo = len(df) + 1
    new_row = {"Serial_number":serialNo,"Bill_number": billnumber, "Name": name, "Phone_number": phone ,"Items": invoice_list, "Amount":total, "Date":date,"Time":time}
                # Append the new row to the DataFrame
    df = df.append(new_row, ignore_index=True)  # ignore_index avoids duplicate row indexing
    df.to_excel('billing_details.xlsx', index=False)
    

def delete_item():
    selected_item = tree.selection()
    if selected_item:
        item_values = tree.item(selected_item, 'values')
        tree.delete(selected_item)
        for item in invoice_list:
            if item[0] == item_values[0] and item[1] == int(item_values[1]):
                invoice_list.remove(item)
                break
        update_total()  # Update total after deleting item

# Dropdown and spinbox for menu items in Menu Frame
desc_label_main = tkinter.Label(menu_frame, text="Items", bg='#f0f8ff', font=("Arial", 10, "bold"))
desc_label_main.grid(row=0, column=0, padx=10, pady=5, sticky="w")
desc_label_box_main = ttk.Combobox(menu_frame, values=list(item_dict.keys()))
desc_label_box_main.grid(row=1, column=0, padx=10, pady=5, sticky="ew")

qty_label_main = tkinter.Label(menu_frame, text="Quantity", bg='#f0f8ff', font=("Arial", 10, "bold"))
qty_label_main.grid(row=0, column=1, padx=10, pady=5, sticky="w")
qty_spinbox_main = tkinter.Spinbox(menu_frame, from_=1, to=10, increment=1)
qty_spinbox_main.grid(row=1, column=1, padx=10, pady=5, sticky="ew")

# Add Item Button
add_item_button = tkinter.Button(menu_frame, text="Add Item", command=add_item, bg='#4682B4', fg='white', font=("Arial", 10, "bold"))
add_item_button.grid(row=1, column=2, padx=10, pady=5)

# Billing Summary Frame
billing_frame = tkinter.LabelFrame(window, text="Billing Summary", padx=10, pady=10, bg='#f0f8ff', font=("Arial", 12, "bold"))
billing_frame.grid(row=3, column=0, padx=20, pady=10, sticky="nsew")

# Treeview for billing summary
columns = ('desc', 'qty', 'price', 'total')
tree = ttk.Treeview(billing_frame, columns=columns, show="headings")
tree.heading('desc', text='Description')
tree.heading('qty', text='Qty')
tree.heading('price', text='Price')
tree.heading('total', text='Total')
tree.grid(row=0, column=0, padx=10, pady=10, sticky="nsew")

# Scrollbar for Treeview
scrollbar = ttk.Scrollbar(billing_frame, orient="vertical", command=tree.yview)
scrollbar.grid(row=0, column=1, sticky='ns')
tree.configure(yscrollcommand=scrollbar.set)

# Frame for displaying Subtotal, SGST, CGST, and Total in a single line
summary_frame = tkinter.Frame(billing_frame, bg='#f0f8ff')
summary_frame.grid(row=2, column=0, columnspan=2, pady=10, sticky="ew")

# Labels to display subtotal, SGST, CGST, and total in a single line
subtotal_label = tkinter.Label(summary_frame, text="Subtotal: 0 INR", bg='#f0f8ff', font=("Arial", 12, "bold"))
subtotal_label.grid(row=0, column=0, padx=10, pady=5, sticky="w")

sgst_label = tkinter.Label(summary_frame, text="SGST (2.5%): 0 INR", bg='#f0f8ff', font=("Arial", 12, "bold"))
sgst_label.grid(row=0, column=1, padx=10, pady=5, sticky="w")

cgst_label = tkinter.Label(summary_frame, text="CGST (2.5%): 0 INR", bg='#f0f8ff', font=("Arial", 12, "bold"))
cgst_label.grid(row=0, column=2, padx=10, pady=5, sticky="w")

total_label = tkinter.Label(summary_frame, text="Total: 0 INR", bg='#f0f8ff', font=("Arial", 12, "bold"))
total_label.grid(row=0, column=3, padx=10, pady=5, sticky="w")

# Configure column weights in summary_frame to spread evenly across the window width
summary_frame.columnconfigure(0, weight=1)
summary_frame.columnconfigure(1, weight=1)
summary_frame.columnconfigure(2, weight=1)
summary_frame.columnconfigure(3, weight=1)

# Frame for action buttons (Save, New Invoice, Delete)
button_frame = tkinter.Frame(window, bg='#f0f8ff')
button_frame.grid(row=4, column=0, padx=10, pady=10, sticky="ew")

# Save Invoice Button
save_invoice_button = tkinter.Button(button_frame, text="Generate Invoice", command=generate_invoice, bg='#32CD32', fg='white', font=("Arial", 10, "bold"))
save_invoice_button.pack(side="left", padx=10, pady=5, anchor="center", expand=True)

# New Invoice Button
new_invoice_button = tkinter.Button(button_frame, text="New Invoice", command=new_invoice, bg='#4682B4', fg='white', font=("Arial", 10, "bold"))
new_invoice_button.pack(side="left", padx=10, pady=5, anchor="center", expand=True)

# Delete Item Button
delete_item_button = tkinter.Button(button_frame, text="Delete Item", command=delete_item, bg='#FF6347', fg='white', font=("Arial", 10, "bold"))
delete_item_button.pack(side="left", padx=10, pady=5, anchor="center", expand=True)

# Make sure the billing frame and treeview expand correctly
billing_frame.rowconfigure(0, weight=1)
billing_frame.columnconfigure(0, weight=1)

window.mainloop()

