import tkinter as tk
from tkinter import ttk, messagebox
from PIL import Image, ImageTk
import pandas as pd
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from ttkthemes import ThemedTk
from datetime import datetime
import os

def connect_to_google_sheet():
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    credentials = ServiceAccountCredentials.from_json_keyfile_name("credentials.json", scope)
    client = gspread.authorize(credentials)
    sheet = client.open("Orders").sheet1
    return sheet

def save_to_google_sheet(order_details):
    try:
        sheet = connect_to_google_sheet()
        sheet.append_row(order_details)
        messagebox.showinfo("Success", "Order details saved to Google Sheet!")
    except Exception as e:
        messagebox.showerror("Error", f"Failed to save to Google Sheets: {e}")

def create_excel_template():
    if not os.path.exists("orders.xlsx"):
        columns = ["Timestamp", "Customer Name", "Contact", "Order Details", "Amount"]
        df = pd.DataFrame(columns=columns)
        df.to_excel("orders.xlsx", index=False)
    return "orders.xlsx"

class ModernOrderManager:
    def __init__(self, root):
        self.root = root
        self.root.title("TRIMIDS Order Management System")
        self.root.geometry("1200x900")
        
        # Create main scrollable canvas
        self.canvas = tk.Canvas(root)
        self.scrollbar = ttk.Scrollbar(root, orient="vertical", command=self.canvas.yview)
        self.scrollable_frame = ttk.Frame(self.canvas)
        
        # Enable mouse wheel scrolling
        self.canvas.bind_all("<MouseWheel>", self._on_mousewheel)

        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        )
        
        self.root.grid_rowconfigure(0, weight=1)
        self.root.grid_columnconfigure(0, weight=1)

        # Configure canvas expansion
        self.canvas.pack_configure(fill="both", expand=True)

        # Make scrollable_frame use full width
        self.scrollable_frame.pack_configure(fill="both", expand=True)

        self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=self.scrollbar.set)
        
        # Style Configuration
        self.style = ttk.Style()
        self.style.configure("Custom.TFrame", background="#ffffff")
        self.style.configure("Custom.TLabel", 
                           background="#ffffff", 
                           font=("Roboto", 12),
                           foreground="#1a237e")
        
        # Main Container
        self.main_frame = ttk.Frame(self.scrollable_frame, style="Custom.TFrame", padding="20")
        self.main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Logo
        try:
            logo_img = Image.open("logo.png")
            logo_img = logo_img.resize((400, 100), Image.LANCZOS)
            self.logo = ImageTk.PhotoImage(logo_img)
            logo_label = ttk.Label(self.main_frame, image=self.logo)
            logo_label.pack(pady=20)
        except:
            print("Logo image not found")
        
        # Header
        header_label = ttk.Label(self.main_frame, 
                               text="TRIMIDS Order Management",
                               font=("Roboto", 32, "bold"),
                               style="Custom.TLabel")
        header_label.pack(pady=20)
        
        self.create_form_container()
        self.create_input_fields()
        self.create_buttons()
        self.create_order_list()
        
        # Pack scrollbar and canvas
        self.scrollbar.pack(side="right", fill="y")
        self.canvas.pack(side="left", fill="both", expand=True)
        
        # Status Bar
        self.status_var = tk.StringVar()
        self.status_bar = ttk.Label(self.main_frame, 
                                  textvariable=self.status_var,
                                  font=("Roboto", 11, "italic"),
                                  style="Custom.TLabel")
        self.status_bar.pack(fill=tk.X, pady=15)

    def _on_mousewheel(self, event):
        self.canvas.yview_scroll(int(-1*(event.delta/120)), "units")

    def create_form_container(self):
        self.form_frame = ttk.LabelFrame(self.main_frame, 
                                       text="Order Details",
                                       padding="20",
                                       style="Custom.TFrame")
        self.form_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)

    def create_input_fields(self):
        # Customer Details
        customer_frame = ttk.LabelFrame(self.form_frame, 
                                      text="Customer Information",
                                      padding="15",
                                      style="Custom.TFrame")
        customer_frame.pack(fill=tk.X, pady=10)
        
        # Input fields with modern styling
        input_style = {"width": 40, 
                      "font": ("Roboto", 11),
                      "bg": "#FFFFFF",
                      "fg": "#1a237e",
                      "relief": "solid"}
        
        self.customer_name = tk.Entry(customer_frame, **input_style)
        self.contact = tk.Entry(customer_frame, **input_style)
        self.order_details = tk.Text(customer_frame, 
                                   height=4,
                                   font=("Roboto", 11),
                                   bg="#FFFFFF",
                                   fg="#1a237e")
        self.amount = tk.Entry(customer_frame, **input_style)
        
        # Grid layout
        ttk.Label(customer_frame, text="Customer Name:", style="Custom.TLabel").grid(row=0, column=0, padx=5, pady=5)
        self.customer_name.grid(row=0, column=1, padx=5, pady=5)
        
        ttk.Label(customer_frame, text="Contact:", style="Custom.TLabel").grid(row=1, column=0, padx=5, pady=5)
        self.contact.grid(row=1, column=1, padx=5, pady=5)
        
        ttk.Label(customer_frame, text="Order Details:", style="Custom.TLabel").grid(row=2, column=0, padx=5, pady=5)
        self.order_details.grid(row=2, column=1, padx=5, pady=5)
        
        ttk.Label(customer_frame, text="Amount ($):", style="Custom.TLabel").grid(row=3, column=0, padx=5, pady=5)
        self.amount.grid(row=3, column=1, padx=5, pady=5)

    def create_buttons(self):
        button_frame = ttk.Frame(self.main_frame, style="Custom.TFrame")
        button_frame.pack(fill=tk.X, pady=20)
        
        # Modern styled buttons
        button_style = {
            "font": ("Roboto", 12, "bold"),
            "padx": 30,
            "pady": 12,
            "bd": 0,
            "cursor": "hand2"
        }
        
        save_btn = tk.Button(button_frame,
                           text="Save Order",
                           command=self.save_order,
                           bg="#1a237e",
                           fg="white",
                           **button_style)
        save_btn.pack(side=tk.LEFT, padx=5)
        
        clear_btn = tk.Button(button_frame,
                            text="Clear Form",
                            command=self.clear_form,
                            bg="#e31e24",
                            fg="white",
                            **button_style)
        clear_btn.pack(side=tk.LEFT, padx=5)
        
        exit_btn = tk.Button(button_frame,
                           text="Exit",
                           command=self.root.quit,
                           bg="#757575",
                           fg="white",
                           **button_style)
        exit_btn.pack(side=tk.RIGHT, padx=5)

    def create_order_list(self):
        # Create scrollable frame for orders
        order_list_frame = ttk.LabelFrame(self.main_frame, 
                                        text="Recent Orders",
                                        padding="10",
                                        style="Custom.TFrame")
        order_list_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)
        
        # Create Treeview with scrollbar
        columns = ("Timestamp", "Customer Name", "Contact", "Order Details", "Amount")
        self.order_tree = ttk.Treeview(order_list_frame, columns=columns, show='headings', height=10)
        
        # Configure columns
        for col in columns:
            self.order_tree.heading(col, text=col)
            self.order_tree.column(col, width=100)
        
        # Add scrollbar to Treeview
        tree_scroll = ttk.Scrollbar(order_list_frame, orient="vertical", command=self.order_tree.yview)
        self.order_tree.configure(yscrollcommand=tree_scroll.set)
        
        # Pack elements
        self.order_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        tree_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Load existing orders
        self.load_orders()

    def load_orders(self):
        try:
            if os.path.exists("orders.xlsx"):
                df = pd.read_excel("orders.xlsx")
                for index, row in df.iterrows():
                    self.order_tree.insert('', tk.END, values=list(row))
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load orders: {e}")

    def save_order(self):
        name = self.customer_name.get()
        contact = self.contact.get()
        details = self.order_details.get("1.0", tk.END).strip()
        amount = self.amount.get()
        
        if not all([name, contact, details, amount]):
            messagebox.showerror("Error", "All fields are required!")
            return
                
        try:
            amount = float(amount)
        except ValueError:
            messagebox.showerror("Error", "Amount must be a valid number!")
            return
                
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        order_data = [timestamp, name, contact, details, amount]
        
        try:
            # Load existing data or create new DataFrame
            if os.path.exists("orders.xlsx"):
                df = pd.read_excel("orders.xlsx")
            else:
                df = pd.DataFrame(columns=["Timestamp", "Customer Name", "Contact", "Order Details", "Amount"])
            
            # Append new order
            new_order = pd.DataFrame([order_data], columns=df.columns)
            df = pd.concat([df, new_order], ignore_index=True)
            
            # Save to Excel
            df.to_excel("orders.xlsx", index=False)
            save_to_google_sheet(order_data)
            
            # Update the treeview
            self.order_tree.insert('', tk.END, values=order_data)
            
            self.status_var.set("Order saved successfully!")
            self.clear_form()
            messagebox.showinfo("Success", "Order has been saved!")
            
        except Exception as e:
            self.status_var.set(f"Error: {str(e)}")
            messagebox.showerror("Error", f"Failed to save order: {str(e)}")

    def clear_form(self):
        self.customer_name.delete(0, tk.END)
        self.contact.delete(0, tk.END)
        self.order_details.delete("1.0", tk.END)
        self.amount.delete(0, tk.END)
        self.status_var.set("")

def main():
    root = ThemedTk(theme="arc")
    create_excel_template()
    app = ModernOrderManager(root)
    root.mainloop()

if __name__ == "__main__":
    main()
