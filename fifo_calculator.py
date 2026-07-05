# fifo_calculator.py
import sys
import pandas as pd
import numpy as np
from datetime import datetime
import tkinter as tk
from tkinter import filedialog, ttk, messagebox
import os
import threading
import time
from queue import Queue

# Import the DateEntry widget from tkcalendar for date picker functionality
try:
    from tkcalendar import DateEntry
    HAS_DATE_PICKER = True
except ImportError:
    HAS_DATE_PICKER = False

# ===== GUI STYLING CONSTANTS =====
# Color scheme for the application interface
PRIMARY_COLOR = "#2c3e50"  # Dark blue for primary elements
SECONDARY_COLOR = "#3498db"  # Light blue for secondary elements
ACCENT_COLOR = "#e74c3c"  # Red for accents and warnings
BACKGROUND_COLOR = "#f8f9fa"  # Light gray for backgrounds
TEXT_COLOR = "#2c3e50"  # Dark blue for text
SUCCESS_COLOR = "#27ae60"  # Green for success messages
BUTTON_HOVER = "#2980b9"  # Darker blue for button hover state
BUTTON_ACTIVE = "#1f618d"  # Even darker blue for button active state
BORDER_COLOR = "#dfe6e9"  # Light gray for borders

# Font definitions for different UI elements
TITLE_FONT = ("Segoe UI", 16, "bold")
HEADER_FONT = ("Segoe UI", 12, "bold")
LABEL_FONT = ("Segoe UI", 10)
BUTTON_FONT = ("Segoe UI", 10, "bold")

class LoginDialog:
    """
    Login dialog class that handles user authentication.
    Displays a login window and validates the entered password against predefined values.
    """
    def __init__(self, parent):
        self.parent = parent
        self.dialog = tk.Toplevel(parent)
        self.dialog.title("تسجيل الدخول")  # "Login" in Arabic
        self.dialog.geometry("450x400")
        self.dialog.resizable(False, False)
        self.dialog.configure(bg=BACKGROUND_COLOR)
        
        # Center the dialog on the parent window
        self.dialog.transient(parent)
        self.dialog.grab_set()
        
        # Valid passwords for authentication
        self.valid_passwords = ["123"]
        self.authenticated = False
        
        self.create_widgets()
        
        # Focus on password entry for better UX
        self.password_entry.focus()
        
        # Wait for the dialog to close before continuing
        self.dialog.wait_window()
    
    def create_widgets(self):
        """
        Create and arrange all the widgets for the login dialog.
        Includes title, username field, password field, login button, and status label.
        """
        # Main container with padding
        main_frame = ttk.Frame(self.dialog, padding=20)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Title section
        title_frame = ttk.Frame(main_frame)
        title_frame.pack(pady=20, fill=tk.X)
        
        ttk.Label(title_frame, text="نظام تحليل المخزون FIFO", 
                 font=TITLE_FONT, foreground=PRIMARY_COLOR).pack()
        
        ttk.Label(title_frame, text="A Program Made for FIFO Profit", 
                 font=("Segoe UI", 11), foreground=SUCCESS_COLOR).pack()
        
        # Decorative separator
        separator = ttk.Separator(main_frame, orient='horizontal')
        separator.pack(fill=tk.X, pady=15)
        
        # Username frame (disabled with fixed value)
        username_frame = ttk.Frame(main_frame)
        username_frame.pack(pady=10, padx=20, fill=tk.X)
        
        ttk.Label(username_frame, text="اسم المستخدم:", font=LABEL_FONT).pack(side=tk.LEFT, padx=5)
        
        # Username entry (disabled with fixed value)
        self.username_var = tk.StringVar(value="admin")
        self.username_entry = ttk.Entry(username_frame, textvariable=self.username_var, state="readonly")
        self.username_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        
        # Password frame
        password_frame = ttk.Frame(main_frame)
        password_frame.pack(pady=10, padx=20, fill=tk.X)
        
        ttk.Label(password_frame, text="كلمة المرور:", font=LABEL_FONT).pack(side=tk.LEFT, padx=5)
        
        self.password_entry = ttk.Entry(password_frame, show="*", font=LABEL_FONT)
        self.password_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        
        # Bind Enter key to login function for better UX
        self.password_entry.bind("<Return>", lambda event: self.login())
        
        # Buttons frame
        buttons_frame = ttk.Frame(main_frame)
        buttons_frame.pack(pady=20)
        
        login_btn = ttk.Button(buttons_frame, text="دخول", command=self.login, style="Accent.TButton")
        login_btn.pack(side=tk.LEFT, padx=10)
        
        cancel_btn = ttk.Button(buttons_frame, text="إلغاء", command=self.cancel)
        cancel_btn.pack(side=tk.LEFT, padx=10)
        
        # Status label for error messages
        self.status_label = ttk.Label(main_frame, text="", foreground=ACCENT_COLOR, font=LABEL_FONT)
        self.status_label.pack(pady=10)
    
    def login(self):
        """
        Validate the entered password against the list of valid passwords.
        If valid, set authenticated to True and close the dialog.
        If invalid, show error message and clear the password field.
        """
        password = self.password_entry.get()
        
        if password in self.valid_passwords:
            self.authenticated = True
            self.dialog.destroy()
        else:
            self.status_label.config(text="كلمة المرور غير صحيحة!")  # "Incorrect password!" in Arabic
            self.password_entry.delete(0, tk.END)
            self.password_entry.focus()
    
    def cancel(self):
        """
        Handle cancel button click.
        Set authenticated to False and close the dialog.
        """
        self.authenticated = False
        self.dialog.destroy()

class FIFOInventoryAnalyzer:
    """
    Main class for FIFO inventory analysis.
    Handles loading data, processing it, calculating FIFO values, and generating reports.
    """
    def __init__(self):
        # Initialize instance variables
        self.df = None  # DataFrame to store the loaded data
        self.selected_date = None  # Date selected for analysis
        self.processed_df = None  # Store the processed data for customer and invoice reports
        
    def load_excel_file(self, file_path):
        """
        Load an Excel file into a pandas DataFrame.
        Removes the 'الباقي' column if it exists and filters out transfer transactions.
        
        Args:
            file_path (str): Path to the Excel file to load
            
        Returns:
            bool: True if file was loaded successfully, False otherwise
        """
        try:
            self.df = pd.read_excel(file_path)
            
            # Remove the existing 'الباقي' column if it exists
            if 'الباقي' in self.df.columns:
                self.df = self.df.drop(columns=['الباقي'])
                print("Removed 'الباقي' column from original data")
            
            # Filter out transfer transactions
            if 'نوع' in self.df.columns:
                self.df = self.df[~self.df['نوع'].str.contains('B-Depo Transferi/تحويل بين المخازن', na=False)]
                print("Excluded transfer transactions between warehouses")
            
            return True
        except Exception as e:
            print(f"Error loading file: {e}")
            return False
    
    def preprocess_data(self, target_date):
        """
        Preprocess the data by filtering to the target date and identifying transaction types.
        
        Args:
            target_date (datetime): Date to filter the data to
            
        Returns:
            pd.DataFrame: Processed DataFrame with identified transaction types
        """
        self.selected_date = target_date
        
        # Filter data up to the target date
        mask = pd.to_datetime(self.df['تاربخ']) <= pd.to_datetime(target_date)
        filtered_df = self.df[mask].copy()
        
        # Create new columns to identify transaction types based on quantities
        filtered_df['is_purchase'] = False
        filtered_df['is_sale'] = False
        filtered_df['is_damage'] = False
        
        # Identify purchases: positive كمية الوارد OR negative كمية الصادر
        purchase_mask = (filtered_df['كمية الوارد'] > 0) | (filtered_df['كمية الصادر'] < 0)
        filtered_df.loc[purchase_mask, 'is_purchase'] = True
        
        # Identify sales: positive كمية الصادر OR negative كمية الوارد
        sale_mask = (filtered_df['كمية الصادر'] > 0) | (filtered_df['كمية الوارد'] < 0)
        filtered_df.loc[sale_mask, 'is_sale'] = True
        
        # Identify damage transactions (original type)
        damage_mask = filtered_df['نوع'] == 'H-Hasarlı/تلف'
        filtered_df.loc[damage_mask, 'is_damage'] = True
        
        # Create effective quantity columns for easier processing
        filtered_df['effective_in'] = 0.0
        filtered_df['effective_out'] = 0.0
        
        # For purchases: positive in OR absolute value of negative out
        filtered_df.loc[filtered_df['is_purchase'] & (filtered_df['كمية الوارد'] > 0), 'effective_in'] = filtered_df['كمية الوارد']
        filtered_df.loc[filtered_df['is_purchase'] & (filtered_df['كمية الصادر'] < 0), 'effective_in'] = -filtered_df['كمية الصادر']
        
        # For sales: positive out OR absolute value of negative in
        filtered_df.loc[filtered_df['is_sale'] & (filtered_df['كمية الصادر'] > 0), 'effective_out'] = filtered_df['كمية الصادر']
        filtered_df.loc[filtered_df['is_sale'] & (filtered_df['كمية الوارد'] < 0), 'effective_out'] = -filtered_df['كمية الوارد']
        
        # Separate purchases and sales
        purchases = filtered_df[filtered_df['is_purchase']].copy()
        sales = filtered_df[filtered_df['is_sale']].copy()
        
        # Reorder: purchases first, then sales
        reordered_df = pd.concat([purchases, sales], ignore_index=True)
        
        return reordered_df
    
    def calculate_fifo(self, df):
        """
        Calculate FIFO values, profit/loss, and inventory costs.
        
        Args:
            df (pd.DataFrame): DataFrame with transaction data
            
        Returns:
            pd.DataFrame: DataFrame with calculated FIFO values
        """
        # Initialize new columns for FIFO calculations
        df['متوسط_التكلفة'] = 0.0  # Will be renamed to FIFO1 later
        df['سعر_الربح'] = 0.0       # Will be renamed to FIFO2 later
        df['إجمالي_تكلفة_المخزون'] = 0.0
        df['الربح_خسارة_FIFO'] = 0.0
        df['الباقي_المحسوب'] = 0
        
        # List to track purchase quantities and costs (FIFO queue)
        inventory_queue = []
        current_balance = 0
        total_inventory_value = 0
        
        # Variables to track total sales and cost of goods sold
        total_units_sold = 0
        total_cogs = 0
        
        for idx, row in df.iterrows():
            is_purchase = row['is_purchase']
            is_sale = row['is_sale']
            is_damage = row['is_damage']
            
            effective_in = row['effective_in'] if pd.notna(row['effective_in']) else 0
            effective_out = row['effective_out'] if pd.notna(row['effective_out']) else 0
            unit_price = row['سعر البيع'] if pd.notna(row['سعر البيع']) else 0
            
            # Update calculated balance
            if idx == 0:
                current_balance = effective_in - effective_out
            else:
                current_balance = df.at[idx-1, 'الباقي_المحسوب'] + effective_in - effective_out
            
            df.at[idx, 'الباقي_المحسوب'] = current_balance
            
            # Track if balance will drop below zero after this transaction
            balance_after_transaction = current_balance
            
            if is_purchase:
                # Purchase transaction
                if effective_in > 0:
                    # Add to FIFO queue
                    inventory_queue.append({
                        'quantity': effective_in,
                        'unit_cost': unit_price,
                        'remaining': effective_in
                    })
                    
                    # Update total inventory value
                    total_inventory_value += effective_in * unit_price
                
                # Calculate current average cost
                if current_balance > 0:
                    current_avg_cost = total_inventory_value / current_balance
                else:
                    current_avg_cost = 0
                
                df.at[idx, 'متوسط_التكلفة'] = current_avg_cost
                df.at[idx, 'إجمالي_تكلفة_المخزون'] = total_inventory_value
                df.at[idx, 'الربح_خسارة_FIFO'] = 0
                
                # Calculate average cost per unit sold so far
                if total_units_sold > 0:
                    df.at[idx, 'سعر_الربح'] = total_cogs / total_units_sold
                else:
                    df.at[idx, 'سعر_الربح'] = 0
                
            elif is_sale:
                # Sale or damage transaction
                if effective_out > 0:
                    # For damaged items, set unit_price to 0
                    if is_damage:
                        unit_price = 0.0
                    
                    # Check if we have enough inventory
                    previous_balance = df.at[idx-1, 'الباقي_المحسوب'] if idx > 0 else 0
                    available_stock = previous_balance
                    
                    # Calculate how much we can actually sell from available stock
                    sellable_quantity = min(effective_out, available_stock) if available_stock > 0 else 0
                    excess_quantity = effective_out - sellable_quantity
                    
                    # Only process if we have some available stock
                    if sellable_quantity > 0:
                        remaining_to_sell = sellable_quantity
                        total_cost_of_goods_sold = 0
                        
                        # Track the costs of the batches being used for this sale
                        batch_costs = []
                        batch_quantities = []
                        
                        # Calculate cost of goods sold using FIFO
                        while remaining_to_sell > 0 and inventory_queue:
                            current_batch = inventory_queue[0]
                            
                            if current_batch['remaining'] >= remaining_to_sell:
                                # Use part of this batch
                                quantity_from_batch = remaining_to_sell
                                cost_from_batch = quantity_from_batch * current_batch['unit_cost']
                                total_cost_of_goods_sold += cost_from_batch
                                
                                # Track the cost and quantity from this batch
                                batch_costs.append(current_batch['unit_cost'])
                                batch_quantities.append(quantity_from_batch)
                                
                                current_batch['remaining'] -= quantity_from_batch
                                total_inventory_value -= cost_from_batch
                                remaining_to_sell = 0
                                
                            else:
                                # Use the entire batch
                                quantity_from_batch = current_batch['remaining']
                                cost_from_batch = quantity_from_batch * current_batch['unit_cost']
                                total_cost_of_goods_sold += cost_from_batch
                                
                                # Track the cost and quantity from this batch
                                batch_costs.append(current_batch['unit_cost'])
                                batch_quantities.append(quantity_from_batch)
                                
                                remaining_to_sell -= quantity_from_batch
                                total_inventory_value -= cost_from_batch
                                inventory_queue.pop(0)
                        
                        # Calculate the weighted average cost for this sale (FIFO2)
                        if batch_quantities and sum(batch_quantities) > 0:
                            weighted_avg_cost = sum(cost * qty for cost, qty in zip(batch_costs, batch_quantities)) / sum(batch_quantities)
                            fifo2_value = weighted_avg_cost
                        else:
                            fifo2_value = 0
                        
                        # Calculate profit/loss based on actual cost of goods sold
                        total_revenue = sellable_quantity * unit_price
                        profit_loss = total_revenue - total_cost_of_goods_sold
                        
                        # Update cumulative totals
                        total_units_sold += sellable_quantity
                        total_cogs += total_cost_of_goods_sold
                        
                        # Calculate current average cost after sale
                        if current_balance > 0:
                            current_avg_cost = total_inventory_value / current_balance
                        else:
                            current_avg_cost = 0
                        
                        # Update columns
                        df.at[idx, 'متوسط_التكلفة'] = current_avg_cost
                        df.at[idx, 'إجمالي_تكلفة_المخزون'] = total_inventory_value
                        df.at[idx, 'الربح_خسارة_FIFO'] = profit_loss
                        df.at[idx, 'سعر_الربح'] = fifo2_value
                    else:
                        # No available stock, set all values to 0
                        df.at[idx, 'متوسط_التكلفة'] = 0
                        df.at[idx, 'سعر_الربح'] = 0
                        df.at[idx, 'إجمالي_تكلفة_المخزون'] = 0
                        df.at[idx, 'الربح_خسارة_FIFO'] = 0
        
        # Rename columns as required
        df = df.rename(columns={
            'متوسط_التكلفة': 'FIFO1',
            'سعر_الربح': 'FIFO2'
        })
        
        return df
    
    def create_summary_sheet(self, df):
        """
        Create a FIFO summary sheet grouped by product.
        
        Args:
            df (pd.DataFrame): DataFrame with calculated FIFO values
            
        Returns:
            pd.DataFrame: Summary DataFrame with aggregated values by product
        """
        summary_data = []
        
        # Group data by product code
        for product_code in df['كود المادة'].unique():
            product_data = df[df['كود المادة'] == product_code]
            product_name = product_data['اسم المادة'].iloc[0]
            
            # Calculate totals
            total_purchases = product_data[product_data['is_purchase']]['effective_in'].sum()
            total_sales = product_data[product_data['is_sale']]['effective_out'].sum()
            
            # Calculate regular and damaged sales separately
            regular_sales = product_data[(product_data['is_sale']) & (~product_data['is_damage'])]['effective_out'].sum()
            damaged_sales = product_data[(product_data['is_sale']) & (product_data['is_damage'])]['effective_out'].sum()
            
            # Last inventory value
            final_inventory_value = product_data['إجمالي_تكلفة_المخزون'].iloc[-1] if len(product_data) > 0 else 0
            
            # Total profit/loss
            total_profit_loss = product_data['الربح_خسارة_FIFO'].sum()
            
            # Final balance
            final_balance = product_data['الباقي_المحسوب'].iloc[-1] if len(product_data) > 0 else 0
            
            # Final FIFO1 (average cost)
            final_fifo1 = product_data['FIFO1'].iloc[-1] if len(product_data) > 0 else 0
            
            # Final FIFO2 (average cost per unit sold)
            final_fifo2 = product_data['FIFO2'].iloc[-1] if len(product_data) > 0 else 0
            
            summary_data.append({
                'كود المادة': product_code,
                'اسم المادة': product_name,
                'إجمالي_المشتريات': total_purchases,
                'إجمالي_المبيعات': total_sales,
                'مبيعات_عادية': regular_sales,
                'مبيعات_تالفة': damaged_sales,
                'الرصيد_النهائي': final_balance,
                'FIFO1': final_fifo1,
                'FIFO2': final_fifo2,
                'قيمة_المخزون_النهائية': final_inventory_value,
                'إجمالي_الربح_خسارة': total_profit_loss
            })
        
        return pd.DataFrame(summary_data)
    
    def add_total_row(self, df):
        """
        Add a total row to the DataFrame.
        
        Args:
            df (pd.DataFrame): DataFrame to add the total row to
            
        Returns:
            pd.DataFrame: DataFrame with added total row
        """
        total_row = {
            'تاربخ': 'الإجمالي',  # "Total" in Arabic
            'FN_ID': '',
            'نوع': '',
            'كود المادة': '',
            'اسم المادة': '',
            'اسم المادة٢': '',
            'الاسم الثالث': '',
            'كود المشتري': '',
            'اسم المشتري': '',
            'ملاحظة': '',
            'سعر البيع': '',
            'كمية الوارد': '',
            'كمية الصادر': '',
            'الباقي_المحسوب': df['الباقي_المحسوب'].iloc[-1] if len(df) > 0 else 0,
            'المخزن': '',
            'FIFO1': '',
            'FIFO2': '',
            'إجمالي_تكلفة_المخزون': df['إجمالي_تكلفة_المخزون'].iloc[-1] if len(df) > 0 else 0,
            'الربح_خسارة_FIFO': df['الربح_خسارة_FIFO'].sum(),
            'is_purchase': '',
            'is_sale': '',
            'is_damage': '',
            'effective_in': df['effective_in'].sum(),
            'effective_out': df['effective_out'].sum(),
        }
        
        # Convert to DataFrame and add it
        total_df = pd.DataFrame([total_row])
        return pd.concat([df, total_df], ignore_index=True)
    
    def add_report_total_row(self, df, label='الإجمالي'):
        """
        Add a total row to a report DataFrame.
        
        Args:
            df (pd.DataFrame): DataFrame to add the total row to
            label (str): Label for the total row
            
        Returns:
            pd.DataFrame: DataFrame with added total row
        """
        total_row = {}
        for col in df.columns:
            if col == df.columns[0]:
                total_row[col] = label
            elif pd.api.types.is_numeric_dtype(df[col]):
                total_row[col] = df[col].sum()
            else:
                total_row[col] = ''
        
        # Convert to DataFrame
        total_df = pd.DataFrame([total_row])
        # Append
        return pd.concat([df, total_df], ignore_index=True)
    
    def analyze_inventory(self, target_date, progress_queue=None):
        """
        Main analysis method - optimized to handle multiple products.
        
        Args:
            target_date (datetime): Date to analyze inventory up to
            progress_queue (Queue): Optional queue for progress updates
            
        Returns:
            tuple: (final_df_with_total, summary_df) DataFrames with analysis results
        """
        if self.df is None:
            return None, None
        
        # Basic data preprocessing
        processed_df = self.preprocess_data(target_date)
        
        # Get all unique product codes
        product_codes = processed_df['كود المادة'].unique()
        total_products = len(product_codes)
        
        # Create a list to store results for each product
        all_products_results = []
        
        # Process each product separately
        for i, product_code in enumerate(product_codes):
            # Update progress
            if progress_queue:
                progress = (i / total_products) * 100
                progress_queue.put(f"يتم التحليل... {progress:.1f}%")
            
            # Filter data for the current product
            product_df = processed_df[processed_df['كود المادة'] == product_code].copy()
            
            # Sort data by date
            product_df = product_df.sort_values('تاربخ')
            
            # Separate purchases and sales
            purchases = product_df[product_df['is_purchase']]
            sales = product_df[product_df['is_sale']]
            
            # Reorder data (purchases first, then sales)
            reordered_product_df = pd.concat([purchases, sales], ignore_index=True)
            
            # Calculate FIFO for the current product
            fifo_product_df = self.calculate_fifo(reordered_product_df)
            
            # Add results to the list
            all_products_results.append(fifo_product_df)
        
        # Merge all results
        final_df = pd.concat(all_products_results, ignore_index=True)
        
        # Save processed data for use in customer and invoice reports
        self.processed_df = final_df.copy()
        
        # Add total row
        final_df_with_total = self.add_total_row(final_df)
        
        # Create summary
        summary_df = self.create_summary_sheet(final_df)
        
        # Add total row to the summary report
        summary_df = self.add_report_total_row(summary_df)
        
        return final_df_with_total, summary_df
    
    def generate_customer_profit_report(self, progress_queue=None):
        """
        Generate a profit report by customer.
        
        Args:
            progress_queue (Queue): Optional queue for progress updates
            
        Returns:
            pd.DataFrame: DataFrame with customer profit data
        """
        if self.processed_df is None:
            return None
        
        # Filter data to get only sales transactions (excluding damaged items)
        sales_df = self.processed_df[self.processed_df['نوع'] == 'S-Satış Fatura/ فاتورة مبيعات'].copy()
        
        # Group data by customer
        customer_profit_data = []
        
        # Get all unique customers
        unique_customers = sales_df['كود المشتري'].unique()
        total_customers = len(unique_customers)
        
        for i, customer_code in enumerate(unique_customers):
            # Update progress
            if progress_queue:
                progress = (i / total_customers) * 100
                progress_queue.put(f"يتم انشاء تقرير الربح حسب المشتري... {progress:.1f}%")
            
            customer_data = sales_df[sales_df['كود المشتري'] == customer_code]
            customer_name = customer_data['اسم المشتري'].iloc[0] if len(customer_data) > 0 else ""
            
            # Calculate total profit for the customer
            total_profit = customer_data['الربح_خسارة_FIFO'].sum()
            
            # Calculate number of invoices for the customer
            invoice_count = customer_data['FN_ID'].nunique()
            
            customer_profit_data.append({
                'كود العميل': customer_code,
                'اسم العميل': customer_name,
                'عدد الفواتير': invoice_count,
                'إجمالي_الربح': total_profit
            })
        
        # Convert to DataFrame and sort by profit
        customer_profit_df = pd.DataFrame(customer_profit_data)
        customer_profit_df = customer_profit_df.sort_values('إجمالي_الربح', ascending=False)
         
        # Add total row at the end of the report
        customer_profit_df = self.add_report_total_row(customer_profit_df)
        
        return customer_profit_df
    
    def generate_invoice_profit_report(self, progress_queue=None):
        """
        Generate a profit report by invoice.
        
        Args:
            progress_queue (Queue): Optional queue for progress updates
            
        Returns:
            pd.DataFrame: DataFrame with invoice profit data
        """
        if self.processed_df is None:
            return None
        
        # Filter data to get only sales transactions (excluding damaged items)
        sales_df = self.processed_df[(self.processed_df['is_sale']) & (~self.processed_df['is_damage'])].copy()
        
        # Group data by invoice
        invoice_profit_data = []
        
        # Get all unique invoices
        unique_invoices = sales_df['FN_ID'].unique()
        total_invoices = len(unique_invoices)
        
        for i, invoice_id in enumerate(unique_invoices):
            # Update progress
            if progress_queue:
                progress = (i / total_invoices) * 100
                progress_queue.put(f"يتم انشاء تقرير الربح حسب الفاتورة... {progress:.1f}%")
            
            invoice_data = sales_df[sales_df['FN_ID'] == invoice_id]
            
            # Get customer information
            customer_code = invoice_data['كود المشتري'].iloc[0] if len(invoice_data) > 0 else ""
            customer_name = invoice_data['اسم المشتري'].iloc[0] if len(invoice_data) > 0 else ""
            
            # Calculate total profit for the invoice
            total_profit = invoice_data['الربح_خسارة_FIFO'].sum()
            
            # Get invoice date
            invoice_date = invoice_data['تاربخ'].iloc[0] if len(invoice_data) > 0 else ""
            
            # Get invoice number
            invoice_number = invoice_data['رقم الفاتورة'].iloc[0] if 'رقم الفاتورة' in invoice_data.columns and len(invoice_data) > 0 else ""
            
            invoice_profit_data.append({
                'رقم الفاتورة': invoice_number,
                'FN_ID': invoice_id,
                'تاريخ الفاتورة': invoice_date,
                'كود العميل': customer_code,
                'اسم العميل': customer_name,
                'إجمالي_الربح': total_profit
            })
        
        # Convert to DataFrame and sort by date
        invoice_profit_df = pd.DataFrame(invoice_profit_data)
        invoice_profit_df = invoice_profit_df.sort_values('تاريخ الفاتورة', ascending=False)
        
        # Add a row for damaged items
        damaged_items = self.processed_df[self.processed_df['is_damage']]
        total_damaged_profit = damaged_items['الربح_خسارة_FIFO'].sum()
        
        damaged_row = {
            'رقم الفاتورة': 'تالف',  # "Damaged" in Arabic
            'FN_ID': 'تالف',
            'تاريخ الفاتورة': '',   # empty string
            'كود العميل': '',
            'اسم العميل': '',
            'إجمالي_الربح': total_damaged_profit
        }
        
        damaged_row_df = pd.DataFrame([damaged_row])
        invoice_profit_df = pd.concat([invoice_profit_df, damaged_row_df], ignore_index=True)
        
        # Add total row for the entire report (including damaged)
        invoice_profit_df = self.add_report_total_row(invoice_profit_df)
        
        return invoice_profit_df
    def generate_monthly_breakdown_sheet(self, progress_queue=None):
        """
        Generate a monthly breakdown sheet showing purchases, sales, and differences by month.
        Fixed to correctly calculate FIFO allocation and total difference.
        
        Args:
            progress_queue (Queue): Optional queue for progress updates
            
        Returns:
            pd.DataFrame: DataFrame with monthly breakdown data
        """
        if self.processed_df is None:
            return None
        
        # Get the selected date
        if self.selected_date is None:
            return None
        
        # Extract year and month from selected date
        selected_year = self.selected_date.year
        selected_month = self.selected_date.month
        
        # Create list of months: مدور (carry-over) + months 1 to selected month
        months = ['مدور']
        for m in range(1, selected_month + 1):
            months.append(f'شهر{m}')
        
        # Get unique products
        product_codes = self.processed_df['كود المادة'].unique()
        total_products = len(product_codes)
        
        # Prepare data list
        data = []
        
        for i, product_code in enumerate(product_codes):
            if progress_queue:
                progress = (i / total_products) * 100
                progress_queue.put(f"Creating monthly breakdown... {progress:.1f}%")
            
            # Get product data
            product_df = self.processed_df[self.processed_df['كود المادة'] == product_code]
            product_name = product_df['اسم المادة'].iloc[0]
            
            # Initialize dictionaries for purchases and sales by month
            purchases_by_month = {month: 0 for month in months}
            sales_by_month = {month: 0 for month in months}
            
            # Process each transaction to record purchases and sales by month
            for _, row in product_df.iterrows():
                date = row['تاربخ']
                
                # Determine month category
                if date < datetime(selected_year, 1, 1):
                    month_key = 'مدور'
                else:
                    # For the selected year, get the month number
                    if date.year == selected_year:
                        month_num = date.month
                        if month_num <= selected_month:
                            month_key = f'شهر{month_num}'
                        else:
                            # Skip months beyond the selected month
                            continue
                    else:
                        # Skip transactions from other years
                        continue
                
                # Record purchases and sales
                if row['is_purchase']:
                    purchases_by_month[month_key] += row['effective_in']
                if row['is_sale']:
                    sales_by_month[month_key] += row['effective_out']
            
            # Create batches for purchases (each batch is a month with the quantity purchased)
            batches = []
            for month in months:
                if purchases_by_month[month] > 0:
                    batches.append({
                        'month': month,
                        'quantity': purchases_by_month[month],
                        'remaining': purchases_by_month[month]
                    })
            
            # Create a list of sales by month (in chronological order)
            sales_list = []
            for month in months:
                if sales_by_month[month] > 0:
                    sales_list.append({
                        'month': month,
                        'quantity': sales_by_month[month]
                    })
            
            # Process sales in FIFO order
            for sale in sales_list:
                remaining_sale = sale['quantity']
                # Go through batches in order (oldest first)
                for batch in batches:
                    if remaining_sale <= 0:
                        break
                    if batch['remaining'] > 0:
                        # Take as much as possible from this batch
                        take = min(batch['remaining'], remaining_sale)
                        batch['remaining'] -= take
                        remaining_sale -= take
            
            # Create difference dictionary (remaining stock by purchase month)
            difference_by_month = {month: 0 for month in months}
            for batch in batches:
                difference_by_month[batch['month']] = batch['remaining']
            
            # Build the three rows for this product
            # Row 1: Purchases
            row_purchase = {
                'كود المادة': product_code,
                'اسم المادة': product_name,
                'نوع': 'وارد-purchase'
            }
            for month in months:
                row_purchase[month] = purchases_by_month[month]
            
            # Row 2: Sales
            row_sale = {
                'كود المادة': product_code,
                'اسم المادة': product_name,
                'نوع': 'صادر-sale'
            }
            for month in months:
                row_sale[month] = sales_by_month[month]
            
            # Row 3: Difference
            row_diff = {
                'كود المادة': product_code,
                'اسم المادة': product_name,
                'نوع': 'الفرق-the difference'
            }
            for month in months:
                row_diff[month] = difference_by_month[month]
            
            data.append(row_purchase)
            data.append(row_sale)
            data.append(row_diff)
        
        # Create DataFrame
        df_monthly = pd.DataFrame(data)
        
        # Reorder columns: fixed columns first, then months in order
        fixed_columns = ['كود المادة', 'اسم المادة', 'نوع']
        columns_order = fixed_columns + months
        df_monthly = df_monthly[columns_order]
        
        # Add a final "Total" column: sum of all month columns for each row
        # (each product has 3 rows - purchase, sale, difference - so this gives
        # total purchases, total sales, and total remaining stock respectively)
        df_monthly['Total'] = df_monthly[months].sum(axis=1)
        
        return df_monthly
    def export_current_stock_report_v2(self, progress_queue=None):
        """
        Export current stock report with batch details (version 2).
        Uses الرصيد_النهائي from summary report for باقي العام value.
        """
        if self.processed_df is None:
            return None
        
        # Generate summary report to get الرصيد_النهائي values
        summary_df = self.create_summary_sheet(self.processed_df)
        
        # Create a mapping of product code to final balance
        product_balance_map = {}
        for _, row in summary_df.iterrows():
            product_balance_map[row['كود المادة']] = row['الرصيد_النهائي']
        
        # Initialize result dataframe
        result_data = []
        
        # Get unique product codes
        product_codes = self.processed_df['كود المادة'].unique()
        total_products = len(product_codes)
        
        for i, product_code in enumerate(product_codes):
            # Update progress
            if progress_queue:
                progress = (i / total_products) * 100
                progress_queue.put(f"يتم انشاء تقرير المخزون الباقي... {progress:.1f}%")
            
            # Get the final balance from the summary report
            final_balance = product_balance_map.get(product_code, 0)
            
            # Skip products with no remaining stock
            if final_balance <= 0:
                continue
            
            # Get all transactions for this product
            # IMPORTANT: Do NOT re-sort by date. processed_df already has the correct
            # FIFO order: all purchases first (sorted by date), then all sales (sorted by date).
            product_df = self.processed_df[self.processed_df['كود المادة'] == product_code].copy()
            
            # Get product name
            product_name = product_df['اسم المادة'].iloc[0]
            
            # We'll create a list of purchase batches
            batches = []  # each batch: { 'date', 'quantity', 'remaining', 'fn_id', 'invoice_number', 'original_row' }
            
            # Process each transaction in the SAME order as calculate_fifo
            for _, row in product_df.iterrows():
                if row['is_purchase']:
                    effective_in = row['effective_in'] if pd.notna(row['effective_in']) else 0
                    if effective_in > 0:
                        # Extract invoice number if available, otherwise use FN_ID
                        invoice_number = row['رقم الفاتورة'] if 'رقم الفاتورة' in row and pd.notna(row['رقم الفاتورة']) else row['FN_ID']
                        
                        batches.append({
                            'date': row['تاربخ'],
                            'quantity': effective_in,
                            'remaining': effective_in,
                            'fn_id': row['FN_ID'],
                            'invoice_number': invoice_number,
                            'original_row': row
                        })
                elif row['is_sale']:
                    effective_out = row['effective_out'] if pd.notna(row['effective_out']) else 0
                    if effective_out > 0:
                        self._process_sale(batches, effective_out)
            
            # Now, find the oldest batch with remaining > 0
            oldest_batch_with_stock = None
            for batch in batches:
                if batch['remaining'] > 0:
                    oldest_batch_with_stock = batch
                    break  # because batches are in chronological order (oldest first)
            
            if oldest_batch_with_stock:
                result_data.append({
                    'كود المادة': product_code,
                    'اسم المادة': product_name,
                    'رقم الفاتورة': oldest_batch_with_stock['invoice_number'],
                    'FN_ID': oldest_batch_with_stock['fn_id'],
                    'تاريخ اقدم فاتورة موجودة': oldest_batch_with_stock['date'],
                    'الكمية المتبقية من اقدم فاتورة موجودة': oldest_batch_with_stock['remaining'],
                    'الكمية الاصلية في اقدم فاتورة موجودة': oldest_batch_with_stock['quantity'],
                    'باقي العام': final_balance  # Use the value from summary report
                })
        
        # Convert to DataFrame
        result_df = pd.DataFrame(result_data)
        
        return result_df

    def _process_sale(self, batches, sale_qty):
        """
        Helper function to process a sale against the batches.
        
        Args:
            batches (list): List of purchase batches
            sale_qty (float): Quantity to sell
        """
        remaining_sale = sale_qty
        for batch in batches:
            if remaining_sale <= 0:
                break
            if batch['remaining'] > 0:
                # Use as much as possible from this batch
                use_from_batch = min(batch['remaining'], remaining_sale)
                batch['remaining'] -= use_from_batch
                remaining_sale -= use_from_batch

class LoadingDialog:
    """
    Dialog class to show loading progress during long operations.
    Displays a progress bar and status messages.
    """
    def __init__(self, parent, title="Processing", message="Please wait..."):
        self.top = tk.Toplevel(parent)
        self.top.title(title)
        self.top.configure(bg=BACKGROUND_COLOR)
        
        # Make dialog modal
        self.top.transient(parent)
        self.top.grab_set()
        
        # Center the dialog
        self.top.geometry("+%d+%d" % (parent.winfo_rootx() + 50, parent.winfo_rooty() + 50))
        
        # Main container with padding
        main_frame = ttk.Frame(self.top, padding=20)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Add message
        self.message_label = ttk.Label(main_frame, text=message, font=LABEL_FONT)
        self.message_label.pack(padx=20, pady=(20, 10))
        
        # Add progress bar
        self.progress = ttk.Progressbar(main_frame, mode='indeterminate', style="Accent.Horizontal.TProgressbar")
        self.progress.pack(padx=20, pady=(0, 10), fill=tk.X)
        self.progress.start()
        
        # Add progress percentage
        self.progress_label = ttk.Label(main_frame, text="0%", font=LABEL_FONT)
        self.progress_label.pack(padx=20, pady=(0, 20))
        
        # Make window not resizable
        self.top.resizable(False, False)
        
        # Initialize progress tracking
        self.progress_value = 0
        
    def update_progress(self, message):
        """
        Update progress message and label.
        
        Args:
            message (str): Progress message to display
        """
        self.message_label.config(text=message)
        
        # Extract percentage from message if available
        if "%" in message:
            try:
                # Extract the percentage value
                percent_str = message.split("%")[0].split()[-1]
                self.progress_value = float(percent_str)
                self.progress_label.config(text=f"{self.progress_value:.1f}%")
            except:
                pass
        
        # Update the UI
        self.top.update()
    
    def close(self):
        """Stop the progress bar and close the dialog."""
        self.progress.stop()
        self.top.destroy()

class FIFOApp:
    """
    Main application class for the FIFO inventory analyzer.
    Creates the GUI and handles user interactions.
    """
    def __init__(self, root):
        self.root = root
        self.root.title("نظام تحليل المخزون FIFO")  # "FIFO Inventory Analysis System"
        self.root.geometry("1400x800")
        self.root.configure(bg=BACKGROUND_COLOR)
        
        # Configure styles
        self.setup_styles()
        
        self.analyzer = FIFOInventoryAnalyzer()
        
        # Show login dialog first
        self.show_login()
    
    def setup_styles(self):
        """
        Configure the styles for the application.
        Sets up colors, fonts, and other visual elements.
        """
        style = ttk.Style()
        
        # Configure the theme
        style.theme_use('clam')
        
        # Configure styles for frames
        style.configure('TFrame', background=BACKGROUND_COLOR)
        
        # Configure styles for labels
        style.configure('TLabel', background=BACKGROUND_COLOR, foreground=TEXT_COLOR, font=LABEL_FONT)
        
        # Configure styles for buttons
        style.configure('TButton', font=BUTTON_FONT, padding=6, relief=tk.FLAT)
        style.map('TButton', 
                 background=[('active', BUTTON_HOVER), ('!active', SECONDARY_COLOR)],
                 foreground=[('active', 'white'), ('!active', 'white')])
        
        # Configure accent button style
        style.configure('Accent.TButton', font=BUTTON_FONT, padding=6, relief=tk.FLAT)
        style.map('Accent.TButton', 
                 background=[('active', BUTTON_ACTIVE), ('!active', PRIMARY_COLOR)],
                 foreground=[('active', 'white'), ('!active', 'white')])
        
        # Configure styles for treeview
        style.configure('Treeview', background='white', foreground=TEXT_COLOR, 
                       fieldbackground=BACKGROUND_COLOR, borderwidth=1, relief=tk.SOLID)
        style.map('Treeview', background=[('selected', SECONDARY_COLOR)])
        
        # Configure styles for progress bar
        style.configure("Accent.Horizontal.TProgressbar", background=SECONDARY_COLOR, 
                       troughcolor=BACKGROUND_COLOR, borderwidth=0)
        
        # Configure styles for heading
        style.configure('Treeview.Heading', background=PRIMARY_COLOR, foreground='white', 
                       font=HEADER_FONT, relief=tk.FLAT)
        
        # Configure styles for separator
        style.configure('TSeparator', background=BORDER_COLOR)
    
    def show_login(self):
        """
        Show login dialog and check authentication.
        If authentication fails, exit the application.
        """
        login_dialog = LoginDialog(self.root)
        
        if login_dialog.authenticated:
            self.create_widgets()
        else:
            # Exit if authentication failed
            self.root.destroy()
            sys.exit(0)
    
    def create_widgets(self):
        """
        Create and arrange all the widgets for the main application.
        Includes header, control panel, data display, and summary sections.
        """
        # Main container with padding
        main_container = ttk.Frame(self.root, padding=20)
        main_container.pack(fill=tk.BOTH, expand=True)
        
        # Header section with logo placeholder
        header_frame = ttk.Frame(main_container)
        header_frame.pack(fill=tk.X, pady=(0, 20))
        
        # Title and subtitle
        title_frame = ttk.Frame(header_frame)
        title_frame.pack(fill=tk.X, pady=10)
        
        ttk.Label(title_frame, text="نظام تحليل المخزون باستخدام FIFO", 
                 font=TITLE_FONT, foreground=PRIMARY_COLOR).pack()
        
        ttk.Label(title_frame, text="A Program Made for FIFO Profit", 
                 font=("Segoe UI", 11), foreground=SUCCESS_COLOR).pack()
        
        # Decorative separator
        separator = ttk.Separator(main_container, orient='horizontal')
        separator.pack(fill=tk.X, pady=10)
        
        # Control panel with card-like appearance
        control_panel = ttk.Frame(main_container)
        control_panel.pack(fill=tk.X, pady=10, padx=10)
        
        # Create a card-like background for controls
        control_card = ttk.Frame(control_panel, style="Card.TFrame")
        control_card.pack(fill=tk.X, pady=5, padx=5)
        
        # Configure card style
        style = ttk.Style()
        style.configure("Card.TFrame", background="white", relief=tk.RAISED, borderwidth=1)
        
        # Add padding inside the card
        card_inner = ttk.Frame(control_card, padding=15)
        card_inner.pack(fill=tk.BOTH, expand=True)
        
        # First row of controls
        row1 = ttk.Frame(card_inner)
        row1.pack(fill=tk.X, pady=5)
        
        ttk.Button(row1, text="تحميل ملف Excel", 
                  command=self.load_file, style="Accent.TButton").pack(side=tk.LEFT, padx=5)
        
        ttk.Label(row1, text="التاريخ:", font=LABEL_FONT).pack(side=tk.LEFT, padx=(20, 5))
        
        # Create date picker widget if available, otherwise use regular entry
        if HAS_DATE_PICKER:
            self.date_entry = DateEntry(
                row1, 
                width=12, 
                background='darkblue',
                foreground='white',
                borderwidth=2,
                date_pattern='yyyy-mm-dd'  # Set format to YYYY-MM-DD
            )
            # Set default date to today
            self.date_entry.set_date(datetime.now())
        else:
            self.date_entry = ttk.Entry(row1, font=LABEL_FONT)
            self.date_entry.pack(side=tk.LEFT, padx=5)
            self.date_entry.insert(0, datetime.now().strftime("%Y-%m-%d"))
        
        self.date_entry.pack(side=tk.LEFT, padx=5)
        
        ttk.Button(row1, text="تحليل البيانات", 
                  command=self.analyze_data_thread, style="Accent.TButton").pack(side=tk.LEFT, padx=5)
        
        # Second row of controls
        row2 = ttk.Frame(card_inner)
        row2.pack(fill=tk.X, pady=5)
        
        ttk.Button(row2, text="تقرير الربح و تكلفة المخزن حسب المادة", 
                  command=self.save_results).pack(side=tk.LEFT, padx=5)
        
        # Buttons for new reports
        ttk.Button(row2, text="تقرير الربح حسب العميل", 
                  command=self.generate_customer_report_thread).pack(side=tk.LEFT, padx=5)
        
        ttk.Button(row2, text="تقرير الربح حسب الفاتورة", 
                  command=self.generate_invoice_report_thread).pack(side=tk.LEFT, padx=5)
        
        # Button to export current stock
        ttk.Button(row2, text="تقرير المخزون المتبقي", 
                  command=self.export_current_stock_thread_v2).pack(side=tk.LEFT, padx=5)
        
        # Data display section
        data_frame = ttk.Frame(main_container)
        data_frame.pack(fill=tk.BOTH, expand=True, pady=10)
        
        # Create a card-like background for the treeview
        tree_card = ttk.Frame(data_frame, style="Card.TFrame")
        tree_card.pack(fill=tk.BOTH, expand=True, pady=5, padx=5)
        
        # Add padding inside the card
        tree_inner = ttk.Frame(tree_card, padding=10)
        tree_inner.pack(fill=tk.BOTH, expand=True)
        
        # Create Treeview for displaying data
        columns = ('تاريخ', 'نوع', 'كود', 'اسم', 'داخل', 'خارج', 'باقي', 'FIFO1', 'FIFO2', 'تكلفة', 'ربح')
        self.tree = ttk.Treeview(tree_inner, columns=columns, show='headings', height=20)
        
        # Define headings
        self.tree.heading('تاريخ', text='تاريخ')
        self.tree.heading('نوع', text='نوع الحركة')
        self.tree.heading('كود', text='كود المادة')
        self.tree.heading('اسم', text='اسم المادة')
        self.tree.heading('داخل', text='كمية داخلة')
        self.tree.heading('خارج', text='كمية خارجة')
        self.tree.heading('باقي', text='الباقي')
        self.tree.heading('FIFO1', text='FIFO1')
        self.tree.heading('FIFO2', text='FIFO2')
        self.tree.heading('تكلفة', text='تكلفة المخزون')
        self.tree.heading('ربح', text='ربح/خسارة')
        
        # Define columns
        self.tree.column('تاريخ', width=100)
        self.tree.column('نوع', width=120)
        self.tree.column('كود', width=80)
        self.tree.column('اسم', width=150)
        self.tree.column('داخل', width=80)
        self.tree.column('خارج', width=80)
        self.tree.column('باقي', width=60)
        self.tree.column('FIFO1', width=70)
        self.tree.column('FIFO2', width=70)
        self.tree.column('تكلفة', width=100)
        self.tree.column('ربح', width=100)
        
        # Scrollbar
        scrollbar = ttk.Scrollbar(tree_inner, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Summary section
        summary_frame = ttk.Frame(main_container)
        summary_frame.pack(fill=tk.X, pady=10)
        
        # Create a card-like background for the summary
        summary_card = ttk.Frame(summary_frame, style="Card.TFrame")
        summary_card.pack(fill=tk.X, pady=5, padx=5)
        
        # Add padding inside the card
        summary_inner = ttk.Frame(summary_card, padding=15)
        summary_inner.pack(fill=tk.BOTH, expand=True)
        
        # Summary text area with a label
        summary_label = ttk.Label(summary_inner, text="ملخص التحليل", font=HEADER_FONT, foreground=PRIMARY_COLOR)
        summary_label.pack(anchor=tk.W, pady=(0, 10))
        
        # Create a frame for the text area with a border
        text_frame = ttk.Frame(summary_inner, relief=tk.SOLID, borderwidth=1)
        text_frame.pack(fill=tk.BOTH, expand=True)
        
        self.summary_text = tk.Text(text_frame, height=8, width=100, font=LABEL_FONT, 
                                   bg="white", fg=TEXT_COLOR, relief=tk.FLAT, 
                                   borderwidth=0, wrap=tk.WORD)
        self.summary_text.pack(fill=tk.BOTH, expand=True, padx=1, pady=1)
        
        # Add a scrollbar to the text area
        text_scrollbar = ttk.Scrollbar(text_frame, orient=tk.VERTICAL, command=self.summary_text.yview)
        self.summary_text.configure(yscrollcommand=text_scrollbar.set)
        text_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Add a note about date picker if not available
        if not HAS_DATE_PICKER:
            self.summary_text.insert(tk.END, "Note: For a better experience, please install the tkcalendar library: pip install tkcalendar\n")
        
        # Add welcome message
        self.summary_text.insert(tk.END, f"Welcome to the FIFO Inventory Analysis System\n")
        self.summary_text.insert(tk.END, f"Login successful\n\n")
    
    def load_file(self):
        """
        Load an Excel file - executed in the main thread.
        Shows a file dialog and loads the selected file.
        """
        file_path = filedialog.askopenfilename(
            title="Select Excel file",
            filetypes=[("Excel files", "*.xls *.xlsx")]
        )
        
        if file_path:
            # Show loading dialog
            loading_dialog = LoadingDialog(self.root, "يتم تحميل الملف", "...يتم تحميل الملف, الرجاء الانتظار...")
            
            # Start a thread to load the file
            def load_file_thread():
                try:
                    success = self.analyzer.load_excel_file(file_path)
                    
                    # Close loading dialog in main thread
                    self.root.after(0, loading_dialog.close)
                    
                    if success:
                        # Add verification that data is loaded
                        if self.analyzer.df is not None and not self.analyzer.df.empty:
                            self.root.after(0, lambda: self.summary_text.insert(tk.END, f"✓ File loaded successfully: {file_path}\n"))
                            self.root.after(0, lambda: self.summary_text.insert(tk.END, f"✓ Number of rows: {len(self.analyzer.df)}\n"))
                        else:
                            self.root.after(0, lambda: self.summary_text.insert(tk.END, "✗ File is empty or contains no data\n"))
                    else:
                        self.root.after(0, lambda: self.summary_text.insert(tk.END, "✗ Failed to load file\n"))
                except Exception as e:
                    self.root.after(0, loading_dialog.close)
                    self.root.after(0, lambda err=str(e): self.summary_text.insert(tk.END, f"✗ Error loading file: {err}\n"))
            
            thread = threading.Thread(target=load_file_thread)
            thread.daemon = True
            thread.start()
    
    def analyze_data_thread(self):
        """
        Analyze data in a separate thread to prevent UI freezing.
        Shows a loading dialog with progress updates.
        """
        loading_dialog = LoadingDialog(self.root, "يتم تحليل البيانات", "...يتم تحليل البيانات, الرجاء الانتظار...")
        
        # Create a queue for progress updates
        progress_queue = Queue()
        
        def progress_monitor():
            """
            Monitor progress updates from the queue and update the loading dialog.
            """
            try:
                while True:
                    # Get progress update from queue
                    message = progress_queue.get_nowait()
                    # Update the loading dialog
                    loading_dialog.update_progress(message)
            except:
                # No more messages in queue
                pass
            finally:
                # Schedule the next check
                self.root.after(100, progress_monitor)
        
        # Start progress monitoring
        self.root.after(100, progress_monitor)
        
        # Start analysis in a separate thread
        def analyze_thread():
            try:
                # Get selected date
                selected_date = self.date_entry.get_date() if HAS_DATE_PICKER else datetime.strptime(self.date_entry.get(), "%Y-%m-%d")
                
                # Analyze inventory
                result_df, summary_df = self.analyzer.analyze_inventory(selected_date, progress_queue)
                
                # Close loading dialog in main thread
                self.root.after(0, loading_dialog.close)
                
                if result_df is not None and summary_df is not None:
                    # Update UI with results
                    self.root.after(0, lambda: self.display_results(result_df))
                    self.root.after(0, lambda: self.summary_text.insert(tk.END, f"✓ Data analyzed successfully\n"))
                    self.root.after(0, lambda: self.summary_text.insert(tk.END, f"✓ Analyzed {len(summary_df)-1} products\n"))  # -1 to exclude total row
                else:
                    self.root.after(0, lambda: self.summary_text.insert(tk.END, "✗ Failed to analyze data\n"))
            except Exception as e:
                self.root.after(0, loading_dialog.close)
                self.root.after(0, lambda err=str(e): self.summary_text.insert(tk.END, f"✗ Error analyzing data: {err}\n"))
        
        thread = threading.Thread(target=analyze_thread)
        thread.daemon = True
        thread.start()
    
    def display_results(self, df):
        """
        Display analysis results in the treeview.
        
        Args:
            df (pd.DataFrame): DataFrame with analysis results
        """
        # Clear existing data
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        # Add new data
        for _, row in df.iterrows():
            # Format values for display
            date = row['تاربخ'] if pd.notna(row['تاربخ']) else ""
            transaction_type = row['نوع'] if pd.notna(row['نوع']) else ""
            product_code = row['كود المادة'] if pd.notna(row['كود المادة']) else ""
            product_name = row['اسم المادة'] if pd.notna(row['اسم المادة']) else ""
            in_qty = row['كمية الوارد'] if pd.notna(row['كمية الوارد']) else 0
            out_qty = row['كمية الصادر'] if pd.notna(row['كمية الصادر']) else 0
            balance = row['الباقي_المحسوب'] if pd.notna(row['الباقي_المحسوب']) else 0
            fifo1 = row['FIFO1'] if pd.notna(row['FIFO1']) else 0
            fifo2 = row['FIFO2'] if pd.notna(row['FIFO2']) else 0
            inventory_cost = row['إجمالي_تكلفة_المخزون'] if pd.notna(row['إجمالي_تكلفة_المخزون']) else 0
            profit_loss = row['الربح_خسارة_FIFO'] if pd.notna(row['الربح_خسارة_FIFO']) else 0
            
            # Format numeric values
            in_qty = f"{in_qty:.2f}" if in_qty != "" else ""
            out_qty = f"{out_qty:.2f}" if out_qty != "" else ""
            balance = f"{balance:.2f}" if balance != "" else ""
            fifo1 = f"{fifo1:.2f}" if fifo1 != "" else ""
            fifo2 = f"{fifo2:.2f}" if fifo2 != "" else ""
            inventory_cost = f"{inventory_cost:.2f}" if inventory_cost != "" else ""
            profit_loss = f"{profit_loss:.2f}" if profit_loss != "" else ""
            
            # Insert row
            self.tree.insert("", "end", values=(
                date, transaction_type, product_code, product_name, 
                in_qty, out_qty, balance, fifo1, fifo2, inventory_cost, profit_loss
            ))
    
    def save_results(self):
        """
        Save analysis results to an Excel file.
        Shows a file dialog and saves the results to the selected location.
        """
        if self.analyzer.processed_df is None:
            messagebox.showerror("خطأ", "الرجاء تحليل البيانات اولا")
            return
        
        file_path = filedialog.asksaveasfilename(
            title="Save Results",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")]
        )
        
        if file_path:
            # Show loading dialog
            loading_dialog = LoadingDialog(self.root, "يتم حفظ النتائج", "يتم حفظ النتائج, الرجاء الانتظار")
            
            # Start a thread to save the file
            def save_file_thread():
                try:
                    # Get selected date
                    selected_date = self.date_entry.get_date() if HAS_DATE_PICKER else datetime.strptime(self.date_entry.get(), "%Y-%m-%d")
                    
                    # Analyze inventory to get fresh results
                    result_df, summary_df = self.analyzer.analyze_inventory(selected_date)
                    
                    if result_df is not None and summary_df is not None:
                        # Create Excel writer
                        with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
                            # Save detailed results
                            result_df.to_excel(writer, sheet_name='FIFO Details', index=False)
                            
                            # Save summary
                            summary_df.to_excel(writer, sheet_name='FIFO Summary', index=False)
                        
                        # Close loading dialog in main thread
                        self.root.after(0, loading_dialog.close)
                        self.root.after(0, lambda: self.summary_text.insert(tk.END, f"✓ Results saved successfully: {file_path}\n"))
                    else:
                        self.root.after(0, loading_dialog.close)
                        self.root.after(0, lambda: self.summary_text.insert(tk.END, "✗ Failed to save results\n"))
                except Exception as e:
                    self.root.after(0, loading_dialog.close)
                    self.root.after(0, lambda err=str(e): self.summary_text.insert(tk.END, f"✗ Error saving results: {err}\n"))
            
            thread = threading.Thread(target=save_file_thread)
            thread.daemon = True
            thread.start()
    
    def generate_customer_report_thread(self):
        """
        Generate customer profit report in a separate thread to prevent UI freezing.
        Shows a loading dialog with progress updates.
        """
        if self.analyzer.processed_df is None:
            messagebox.showerror("خطأ", "الرجاء تحليل البيانات اولا")
            return
        
        loading_dialog = LoadingDialog(self.root, "تقرير الربح حسب العميل", "يتم انشاء تقرير الربح حسب العميل, الرجاء الانتظار")
        
        # Create a queue for progress updates
        progress_queue = Queue()
        
        def progress_monitor():
            """
            Monitor progress updates from the queue and update the loading dialog.
            """
            try:
                while True:
                    # Get progress update from queue
                    message = progress_queue.get_nowait()
                    # Update the loading dialog
                    loading_dialog.update_progress(message)
            except:
                # No more messages in queue
                pass
            finally:
                # Schedule the next check
                self.root.after(100, progress_monitor)
        
        # Start progress monitoring
        self.root.after(100, progress_monitor)
        
        # Start report generation in a separate thread
        def generate_report_thread():
            try:
                # Generate customer report
                customer_df = self.analyzer.generate_customer_profit_report(progress_queue)
                
                # Close loading dialog in main thread
                self.root.after(0, loading_dialog.close)
                
                if customer_df is not None:
                    # Show save dialog
                    file_path = filedialog.asksaveasfilename(
                        title="Save Customer Report",
                        defaultextension=".xlsx",
                        filetypes=[("Excel files", "*.xlsx")]
                    )
                    
                    if file_path:
                        # Save report
                        customer_df.to_excel(file_path, index=False)
                        self.root.after(0, lambda: self.summary_text.insert(tk.END, f"✓ Customer report saved successfully: {file_path}\n"))
                else:
                    self.root.after(0, lambda: self.summary_text.insert(tk.END, "✗ Failed to generate customer report\n"))
            except Exception as e:
                self.root.after(0, loading_dialog.close)
                self.root.after(0, lambda err=str(e): self.summary_text.insert(tk.END, f"✗ Error generating customer report: {err}\n"))
        
        thread = threading.Thread(target=generate_report_thread)
        thread.daemon = True
        thread.start()
    
    def generate_invoice_report_thread(self):
        """
        Generate invoice profit report in a separate thread to prevent UI freezing.
        Shows a loading dialog with progress updates.
        """
        if self.analyzer.processed_df is None:
            messagebox.showerror("خطأ", "الرجاء تحليل البيانات اولا")
            return
        
        loading_dialog = LoadingDialog(self.root, "تقرير الربح حسب الفاتورة", "يتم انشاء تقرير الربح حسب الفاتورة, الرجاء الانتظار")
        
        # Create a queue for progress updates
        progress_queue = Queue()
        
        def progress_monitor():
            """
            Monitor progress updates from the queue and update the loading dialog.
            """
            try:
                while True:
                    # Get progress update from queue
                    message = progress_queue.get_nowait()
                    # Update the loading dialog
                    loading_dialog.update_progress(message)
            except:
                # No more messages in queue
                pass
            finally:
                # Schedule the next check
                self.root.after(100, progress_monitor)
        
        # Start progress monitoring
        self.root.after(100, progress_monitor)
        
        # Start report generation in a separate thread
        def generate_report_thread():
            try:
                # Generate invoice report
                invoice_df = self.analyzer.generate_invoice_profit_report(progress_queue)
                
                # Close loading dialog in main thread
                self.root.after(0, loading_dialog.close)
                
                if invoice_df is not None:
                    # Show save dialog
                    file_path = filedialog.asksaveasfilename(
                        title="Save Invoice Report",
                        defaultextension=".xlsx",
                        filetypes=[("Excel files", "*.xlsx")]
                    )
                    
                    if file_path:
                        # Save report
                        invoice_df.to_excel(file_path, index=False)
                        self.root.after(0, lambda: self.summary_text.insert(tk.END, f"✓ Invoice report saved successfully: {file_path}\n"))
                else:
                    self.root.after(0, lambda: self.summary_text.insert(tk.END, "✗ Failed to generate invoice report\n"))
            except Exception as e:
                self.root.after(0, loading_dialog.close)
                self.root.after(0, lambda err=str(e): self.summary_text.insert(tk.END, f"✗ Error generating invoice report: {err}\n"))
        
        thread = threading.Thread(target=generate_report_thread)
        thread.daemon = True
        thread.start()
    
    def export_current_stock_thread_v2(self):
        """
        Thread function to export current stock report with monthly breakdown.
        Creates two sheets: current stock report and monthly breakdown.
        """
        # Create a loading dialog
        loading_dialog = LoadingDialog(self.root, "إنشاء تقرير المخزون المتبقي", "الرجاء الانتظار...")
        
        # Create a queue for progress updates
        progress_queue = Queue()
        
        # Create a thread to run the export
        def export_thread():
            try:
                # Get the current stock report
                current_stock_df = self.analyzer.export_current_stock_report_v2(progress_queue)
                
                # Get the monthly breakdown sheet
                monthly_df = self.analyzer.generate_monthly_breakdown_sheet(progress_queue)
                
                # If both are not None, save to Excel
                if current_stock_df is not None and monthly_df is not None:
                    # Ask for save location
                    file_path = filedialog.asksaveasfilename(
                        defaultextension=".xlsx",
                        filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
                        title="حفظ تقرير المخزون المتبقي"
                    )
                    
                    if file_path:
                        # Save to Excel with two sheets
                        with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
                            current_stock_df.to_excel(writer, sheet_name='تقرير المخزون المتبقي', index=False)
                            monthly_df.to_excel(writer, sheet_name='التحليل الشهري', index=False)
                        
                        # Show success message
                        progress_queue.put("success")
                else:
                    progress_queue.put("error")
            except Exception as e:
                progress_queue.put(f"error: {str(e)}")
        
        # Start the thread
        thread = threading.Thread(target=export_thread)
        thread.start()
        
        # Update the loading dialog with progress
        while thread.is_alive():
            if not progress_queue.empty():
                message = progress_queue.get()
                if message == "success":
                    loading_dialog.close()
                    messagebox.showinfo("نجاح", "تم حفظ التقرير بنجاح")
                    break
                elif message.startswith("error"):
                    loading_dialog.close()
                    error_msg = message.replace("error: ", "") if message.startswith("error: ") else "فشل في إنشاء التقارير"
                    messagebox.showerror("خطأ", error_msg)
                    break
                else:
                    loading_dialog.update_progress(message)
            self.root.update()
            time.sleep(0.1)
        
        loading_dialog.close()

# Main entry point
if __name__ == "__main__":
    root = tk.Tk()
    app = FIFOApp(root)
    root.mainloop()
