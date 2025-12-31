import tkinter as tk
from tkinter import ttk, messagebox, filedialog, scrolledtext
from datetime import datetime, timedelta
from scr.parsers.mpesa_parser import MPesaParser
from scr.database.db_manager import DatabaseManager
from scr.excel.excel_handler import ExcelHandler


class MPesaApp:
    """Main GUI application with enhanced features"""
    
    def __init__(self, root):
        self.root = root
        self.root.title("M-Pesa Transaction Manager")
        self.root.geometry("950x800")
        
        self.parser = MPesaParser()
        self.db_manager = DatabaseManager()
        self.excel_handler = None
        
        self.stat_card_labels = {}
        
        self.setup_ui()
        self.root.after(100, self.load_statistics)
    
    def setup_ui(self):
        """Setup the enhanced user interface"""
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        
        main_canvas = tk.Canvas(self.root)
        scrollbar = ttk.Scrollbar(self.root, orient="vertical", command=main_canvas.yview)
        
        main_frame = ttk.Frame(main_canvas, padding="10")
        
        main_frame.bind(
            "<Configure>",
            lambda e: main_canvas.configure(scrollregion=main_canvas.bbox("all"))
        )
        
        main_canvas.create_window((0, 0), window=main_frame, anchor="nw")
        main_canvas.configure(yscrollcommand=scrollbar.set)
        
        main_canvas.grid(row=0, column=0, sticky="nsew")
        scrollbar.grid(row=0, column=1, sticky="ns")
        
        # Title
        title_frame = ttk.Frame(main_frame)
        title_frame.grid(row=0, column=0, columnspan=2, pady=10)
        
        title = ttk.Label(title_frame, text="M-Pesa Transaction Manager", 
                         font=('Arial', 18, 'bold'), foreground='#2c3e50')
        title.pack()
        
        subtitle = ttk.Label(title_frame, text="Manage your M-Pesa transactions efficiently", 
                           font=('Arial', 9, 'italic'), foreground='#7f8c8d')
        subtitle.pack()
        
        # Excel file selection
        excel_frame = ttk.LabelFrame(main_frame, text="Excel File Configuration", padding="10")
        excel_frame.grid(row=1, column=0, columnspan=2, sticky="we", pady=10)
        
        self.excel_path_var = tk.StringVar(value="transactions.xlsx")
        
        ttk.Label(excel_frame, text="File Path:").grid(row=0, column=0, sticky=tk.W, padx=5)
        ttk.Entry(excel_frame, textvariable=self.excel_path_var, width=50).grid(
            row=0, column=1, padx=5, pady=5, sticky="we")
        ttk.Button(excel_frame, text="Browse", command=self.browse_excel).grid(
            row=0, column=2, padx=5)
        ttk.Button(excel_frame, text="Export All", 
                  command=self.export_all, style='Accent.TButton').grid(
            row=0, column=3, padx=5)
        
        excel_frame.columnconfigure(1, weight=1)
        
        # Message input
        input_frame = ttk.LabelFrame(main_frame, text="Paste M-Pesa Messages Here", padding="10")
        input_frame.grid(row=2, column=0, columnspan=2, sticky="nsew", pady=10)
        
        instructions = ttk.Label(input_frame, 
            text="Tip: Paste one or multiple M-Pesa messages (separate with blank lines)",
            font=('Arial', 9, 'italic'), foreground='#3498db')
        instructions.grid(row=0, column=0, sticky=tk.W, pady=(0, 5))
        
        self.message_text = scrolledtext.ScrolledText(input_frame, height=12, width=85, 
                                                       font=('Courier', 9),
                                                       wrap=tk.WORD)
        self.message_text.grid(row=1, column=0, sticky="nsew", pady=5)
        
        sample_frame = ttk.Frame(input_frame)
        sample_frame.grid(row=2, column=0, sticky=tk.W, pady=5)
        
        ttk.Button(sample_frame, text="Load Sample", 
                  command=self.load_sample_message).pack(side=tk.LEFT, padx=5)
        ttk.Button(sample_frame, text="Paste from Clipboard", 
                  command=self.paste_from_clipboard).pack(side=tk.LEFT, padx=5)
        
        input_frame.columnconfigure(0, weight=1)
        input_frame.rowconfigure(1, weight=1)
        
        # Action Buttons
        button_frame = ttk.Frame(main_frame, padding="5")
        button_frame.grid(row=3, column=0, columnspan=2, pady=10)
        
        ttk.Button(button_frame, text="Parse & Save", command=self.parse_and_save,
                  style='Accent.TButton', width=20).grid(row=0, column=0, padx=5)
        ttk.Button(button_frame, text="Clear Input", command=self.clear_input,
                  width=15).grid(row=0, column=1, padx=5)
        ttk.Button(button_frame, text="View All Transactions", 
                  command=self.view_all_transactions, width=20).grid(row=0, column=2, padx=5)
        ttk.Button(button_frame, text="Refresh Stats", 
                  command=self.load_statistics, width=15).grid(row=0, column=3, padx=5)
        
        # Quick Stats Cards
        cards_frame = ttk.Frame(main_frame)
        cards_frame.grid(row=4, column=0, columnspan=2, pady=10, sticky="we")
        
        self.create_stat_cards(cards_frame)
        
        # Detailed Statistics
        stats_frame = ttk.LabelFrame(main_frame, text="Statistics Dashboard", padding="10")
        stats_frame.grid(row=5, column=0, columnspan=2, sticky="we", pady=10)
        
        stats_content = ttk.Frame(stats_frame)
        stats_content.grid(row=0, column=0, sticky="we")
        
        self.stats_text = tk.Text(stats_content, height=12, width=85, 
                                 font=('Courier', 10), 
                                 relief=tk.GROOVE, 
                                 borderwidth=2,
                                 background='#f8f9fa',
                                 state='normal')
        self.stats_text.grid(row=0, column=0, sticky="we", padx=5, pady=5)
        
        self.stats_text.tag_configure('header', font=('Arial', 12, 'bold'), foreground='#2c3e50')
        self.stats_text.tag_configure('value', font=('Courier', 11, 'bold'), foreground='#27ae60')
        self.stats_text.tag_configure('label', font=('Arial', 10), foreground='#34495e')
        self.stats_text.tag_configure('type', font=('Arial', 9), foreground='#7f8c8d')
        self.stats_text.tag_configure('warning', font=('Arial', 10, 'italic'), foreground='#e74c3c')
        
        stats_content.columnconfigure(0, weight=1)
        
        # Status bar
        status_frame = ttk.Frame(main_frame, relief=tk.SUNKEN, borderwidth=1)
        status_frame.grid(row=6, column=0, columnspan=2, sticky="we", pady=(10, 0))
        
        self.status_var = tk.StringVar(value="Ready")
        status_label = ttk.Label(status_frame, textvariable=self.status_var, 
                                anchor=tk.W, font=('Arial', 9))
        status_label.grid(row=0, column=0, sticky="we", padx=5, pady=2)
        
        self.progress = ttk.Progressbar(status_frame, mode='indeterminate', length=200)
        self.progress.grid(row=0, column=1, padx=5, pady=2)
        
        status_frame.columnconfigure(0, weight=1)
        
        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(2, weight=1)
    
    def create_stat_cards(self, parent):
        """Create quick stat cards"""
        self.stat_card_labels['total_trans'] = self.create_card(
            parent, "Total Transactions", "0", "#3498db", 0, 0)
        
        self.stat_card_labels['total_amount'] = self.create_card(
            parent, "Total Amount", "KSh 0.00", "#27ae60", 0, 1)
        
        self.stat_card_labels['total_fees'] = self.create_card(
            parent, "Total Fees", "KSh 0.00", "#e74c3c", 0, 2)
        
        self.stat_card_labels['net_amount'] = self.create_card(
            parent, "Net Amount", "KSh 0.00", "#9b59b6", 0, 3)
        
        for i in range(4):
            parent.columnconfigure(i, weight=1)
    
    def create_card(self, parent, title, value, color, row, col):
        """Create a single stat card and return the value label"""
        card = ttk.Frame(parent, relief=tk.RAISED, borderwidth=2)
        card.grid(row=row, column=col, padx=5, sticky="ew")
        
        title_label = ttk.Label(card, text=title, font=('Arial', 9), foreground='#7f8c8d')
        title_label.pack(pady=(5, 0))
        
        value_label = ttk.Label(card, text=value, font=('Arial', 14, 'bold'), 
                               foreground=color)
        value_label.pack(pady=(0, 5))
        
        return value_label
    
    def update_stat_cards(self, stats):
        """Update quick stat cards"""
        self.stat_card_labels['total_trans'].config(
            text=f"{stats['total_transactions']:,}")
        
        self.stat_card_labels['total_amount'].config(
            text=f"KSh {stats['total_amount']:,.2f}")
        
        self.stat_card_labels['total_fees'].config(
            text=f"KSh {stats['total_fees']:,.2f}")
        
        net = stats['total_amount'] - stats['total_fees']
        self.stat_card_labels['net_amount'].config(
            text=f"KSh {net:,.2f}")
    
    def load_sample_message(self):
        """Load sample M-Pesa messages for testing"""
        sample = """RBK1234567 Confirmed. Ksh500.00 sent to JOHN DOE 254712345678 on 15/1/24 at 2:30 PM. New M-PESA balance is Ksh2,500.00. Transaction cost, Ksh15.00

RBK9876543 Confirmed. You have received Ksh1,000.00 from JANE SMITH 254723456789 on 16/1/24 at 3:45 PM. New M-PESA balance is Ksh3,500.00

RBL1122334 Confirmed. Ksh200.00 paid to KPLC PREPAID on 17/1/24 at 10:15 AM. Account 71234567890. New M-PESA balance is Ksh3,285.00. Transaction cost, Ksh15.00"""
        
        self.message_text.delete("1.0", tk.END)
        self.message_text.insert("1.0", sample)
        self.status_var.set("Sample messages loaded - Click 'Parse & Save' to process")
    
    def paste_from_clipboard(self):
        """Paste from clipboard"""
        try:
            clipboard_content = self.root.clipboard_get()
            self.message_text.delete("1.0", tk.END)
            self.message_text.insert("1.0", clipboard_content)
            self.status_var.set("Pasted from clipboard")
        except:
            messagebox.showwarning("Warning", "Clipboard is empty or contains invalid data")
    
    def browse_excel(self):
        """Browse for Excel file"""
        filename = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
            title="Select Excel file location"
        )
        if filename:
            self.excel_path_var.set(filename)
            self.status_var.set(f"Excel file set to: {filename}")
    
    def get_excel_handler(self):
        """Get or create Excel handler"""
        if not self.excel_handler or str(self.excel_handler.excel_path) != self.excel_path_var.get():
            self.excel_handler = ExcelHandler(self.excel_path_var.get())
        return self.excel_handler
    
    def start_progress(self):
        """Start progress bar animation"""
        self.progress.start(10)
    
    def stop_progress(self):
        """Stop progress bar animation"""
        self.progress.stop()
    
    def parse_and_save(self):
        """Parse messages and save to database and Excel with validation"""
        messages_text = self.message_text.get("1.0", tk.END).strip()
        
        if not messages_text:
            messagebox.showwarning("Warning", "Please paste M-Pesa messages first")
            return
        
        try:
            self.start_progress()
            
            messages = [msg.strip() for msg in messages_text.split('\n\n') if msg.strip()]
            if not messages:
                messages = [messages_text]
            
            self.status_var.set(f"Parsing {len(messages)} message(s)...")
            self.root.update()
            
            transactions = self.parser.parse_multiple_messages(messages)
            
            if not transactions:
                self.stop_progress()
                messagebox.showerror("Error", 
                    "Could not parse any valid transactions.\n\n" +
                    "Please ensure the messages are in valid M-Pesa format.\n" +
                    "Click 'Load Sample' to see example messages.")
                self.status_var.set("Error: No valid transactions found")
                return
            
            if not self.confirm_transactions(transactions):
                self.stop_progress()
                self.status_var.set("Operation cancelled by user")
                return
            
            self.status_var.set("Saving to database...")
            self.root.update()
            success, duplicates = self.db_manager.insert_multiple_transactions(transactions)
            
            self.status_var.set("Updating Excel file...")
            self.root.update()
            excel_handler = self.get_excel_handler()
            excel_success = excel_handler.append_transactions(transactions)
            
            self.stop_progress()
            
            message = f"Successfully processed {len(transactions)} transaction(s):\n\n"
            message += f"  New transactions saved: {success}\n"
            message += f"  Duplicates skipped: {duplicates}\n"
            message += f"  Excel file updated: {'Yes' if excel_success else 'No'}\n\n"
            
            if excel_success:
                message += f"File location: {self.excel_path_var.get()}"
            
            messagebox.showinfo("Success", message)
            self.status_var.set(f"Successfully saved {success} new transaction(s)")
            
            self.clear_input()
            self.load_statistics()
            
        except Exception as e:
            self.stop_progress()
            messagebox.showerror("Error", f"An error occurred:\n\n{str(e)}")
            self.status_var.set(f"Error: {str(e)}")
    
    def confirm_transactions(self, transactions):
        """Show confirmation dialog with transaction preview"""
        if len(transactions) <= 3:
            return True
        
        preview = f"Found {len(transactions)} transactions:\n\n"
        for i, trans in enumerate(transactions[:3], 1):
            preview += f"{i}. {trans.transaction_code} - KSh {trans.amount:,.2f} ({trans.transaction_type})\n"
        
        if len(transactions) > 3:
            preview += f"... and {len(transactions) - 3} more\n"
        
        preview += f"\nTotal Amount: KSh {sum(t.amount for t in transactions):,.2f}\n"
        preview += f"Total Fees: KSh {sum(t.fee for t in transactions):,.2f}\n\n"
        preview += "Do you want to save these transactions?"
        
        return messagebox.askyesno("Confirm Transactions", preview)
    
    def clear_input(self):
        """Clear the message input"""
        self.message_text.delete("1.0", tk.END)
        self.status_var.set("Input cleared")
    
    def load_statistics(self):
        """Load and display enhanced statistics"""
        try:
            self.status_var.set("Loading statistics...")
            self.root.update()
            
            stats = self.db_manager.get_statistics()
            
            self.update_stat_cards(stats)
            
            self.stats_text.config(state='normal')
            self.stats_text.delete("1.0", tk.END)
            
            self.stats_text.insert(tk.END, "=" * 85 + "\n")
            self.stats_text.insert(tk.END, "  TRANSACTION SUMMARY REPORT\n", 'header')
            self.stats_text.insert(tk.END, "=" * 85 + "\n\n")
            
            if stats['total_transactions'] == 0:
                self.stats_text.insert(tk.END, "No transactions found in database.\n\n", 'warning')
                self.stats_text.insert(tk.END, "   Get started by:\n", 'label')
                self.stats_text.insert(tk.END, "   1. Click 'Load Sample' to see example messages\n", 'type')
                self.stats_text.insert(tk.END, "   2. Paste your M-Pesa messages above\n", 'type')
                self.stats_text.insert(tk.END, "   3. Click 'Parse & Save' to process\n\n", 'type')
            else:
                self.stats_text.insert(tk.END, "Financial Overview:\n", 'label')
                self.stats_text.insert(tk.END, "-" * 85 + "\n")
                
                self.stats_text.insert(tk.END, f"  Total Transactions:    ")
                self.stats_text.insert(tk.END, f"{stats['total_transactions']:,}\n", 'value')
                
                self.stats_text.insert(tk.END, f"  Gross Amount:          ")
                self.stats_text.insert(tk.END, f"KSh {stats['total_amount']:>15,.2f}\n", 'value')
                
                self.stats_text.insert(tk.END, f"  Total Fees Paid:       ")
                self.stats_text.insert(tk.END, f"KSh {stats['total_fees']:>15,.2f}\n", 'value')
                
                net_amount = stats['total_amount'] - stats['total_fees']
                self.stats_text.insert(tk.END, f"  Net Amount:            ")
                self.stats_text.insert(tk.END, f"KSh {net_amount:>15,.2f}\n", 'value')
                
                if stats['total_transactions'] > 0:
                    avg_trans = stats['total_amount'] / stats['total_transactions']
                    self.stats_text.insert(tk.END, f"  Average Transaction:   ")
                    self.stats_text.insert(tk.END, f"KSh {avg_trans:>15,.2f}\n", 'value')
                
                if stats['by_type']:
                    self.stats_text.insert(tk.END, "\n\nBreakdown by Transaction Type:\n", 'label')
                    self.stats_text.insert(tk.END, "-" * 85 + "\n")
                    
                    self.stats_text.insert(tk.END, 
                        f"  {'Type':<15} {'Count':>8} {'Total Amount':>18} {'Avg Amount':>18} {'%':>8}\n", 
                        'type')
                    self.stats_text.insert(tk.END, "  " + "-" * 80 + "\n")
                    
                    for item in stats['by_type']:
                        trans_type = item['transaction_type'].capitalize()
                        count = item['count']
                        total = item['total']
                        avg = total / count if count > 0 else 0
                        percentage = (count / stats['total_transactions'] * 100) if stats['total_transactions'] > 0 else 0
                        
                        self.stats_text.insert(tk.END, f"  {trans_type:<15}", 'type')
                        self.stats_text.insert(tk.END, f"{count:>8}", 'value')
                        self.stats_text.insert(tk.END, f"  KSh {total:>12,.2f}", 'value')
                        self.stats_text.insert(tk.END, f"  KSh {avg:>12,.2f}", 'type')
                        self.stats_text.insert(tk.END, f"{percentage:>8.1f}%\n", 'type')
            
            self.stats_text.insert(tk.END, "\n" + "=" * 85 + "\n")
            self.stats_text.insert(tk.END, 
                f"Last updated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n", 'type')
            
            self.stats_text.config(state='disabled')
            
            self.status_var.set("Statistics loaded successfully")
            
        except Exception as e:
            self.stats_text.config(state='normal')
            self.stats_text.delete("1.0", tk.END)
            self.stats_text.insert(tk.END, f"Error loading statistics:\n\n{str(e)}\n\n", 'warning')
            self.stats_text.insert(tk.END, "Please try refreshing or check the database connection.", 'type')
            self.stats_text.config(state='disabled')
            self.status_var.set(f"Error loading statistics")
            print(f"Statistics error: {e}")
    
    def view_all_transactions(self):
        """Open enhanced transaction viewer window"""
        try:
            ViewTransactionsWindow(self.root, self.db_manager, self)
            self.status_var.set("Opened transactions viewer")
        except Exception as e:
            messagebox.showerror("Error", f"Could not open transactions viewer:\n\n{str(e)}")
    
    def export_all(self):
        """Export all database records to Excel"""
        try:
            self.start_progress()
            self.status_var.set("Exporting to Excel...")
            self.root.update()
            
            excel_handler = self.get_excel_handler()
            if excel_handler.export_from_database(self.db_manager):
                self.stop_progress()
                messagebox.showinfo("Success", 
                    f"All transactions exported successfully!\n\n" +
                    f"Total: {self.db_manager.get_statistics()['total_transactions']} transactions\n" +
                    f"File: {self.excel_path_var.get()}")
                self.status_var.set("Export completed successfully")
            else:
                self.stop_progress()
                messagebox.showerror("Error", "Failed to export to Excel")
                self.status_var.set("Export failed")
        except Exception as e:
            self.stop_progress()
            messagebox.showerror("Error", f"Export error:\n\n{str(e)}")
            self.status_var.set(f"Error: {str(e)}")


class ViewTransactionsWindow:
    """Enhanced window to view, search, filter, and manage transactions"""
    
    def __init__(self, parent, db_manager, main_app):
        self.window = tk.Toplevel(parent)
        self.window.title("Transaction Viewer")
        self.window.geometry("1200x700")
        
        self.db_manager = db_manager
        self.main_app = main_app
        self.all_transactions = []
        
        self.sort_reverse = {}
        
        self.setup_ui()
        self.load_data()
    
    def setup_ui(self):
        """Setup enhanced UI with search and filter"""
        main_frame = ttk.Frame(self.window, padding="10")
        main_frame.grid(row=0, column=0, sticky="nsew")
        
        title = ttk.Label(main_frame, text="Transaction History", 
                         font=('Arial', 16, 'bold'))
        title.grid(row=0, column=0, columnspan=4, pady=10)
        
        filter_frame = ttk.LabelFrame(main_frame, text="Search & Filter", padding="10")
        filter_frame.grid(row=1, column=0, columnspan=4, sticky="ew", pady=10)
        
        ttk.Label(filter_frame, text="Search:").grid(row=0, column=0, padx=5)
        self.search_var = tk.StringVar()
        self.search_var.trace('w', lambda *args: self.apply_filters())
        search_entry = ttk.Entry(filter_frame, textvariable=self.search_var, width=30)
        search_entry.grid(row=0, column=1, padx=5)
        
        ttk.Label(filter_frame, text="Type:").grid(row=0, column=2, padx=5)
        self.type_var = tk.StringVar(value="All")
        type_combo = ttk.Combobox(filter_frame, textvariable=self.type_var, 
                                  values=["All", "Sent", "Received", "Withdrawn", "Paybill", "Airtime"],
                                  width=15, state='readonly')
        type_combo.grid(row=0, column=3, padx=5)
        type_combo.bind('<<ComboboxSelected>>', lambda e: self.apply_filters())
        
        ttk.Label(filter_frame, text="From:").grid(row=0, column=4, padx=5)
        self.date_from_var = tk.StringVar()
        date_from = ttk.Entry(filter_frame, textvariable=self.date_from_var, width=12)
        date_from.grid(row=0, column=5, padx=5)
        
        ttk.Label(filter_frame, text="To:").grid(row=0, column=6, padx=5)
        self.date_to_var = tk.StringVar()
        date_to = ttk.Entry(filter_frame, textvariable=self.date_to_var, width=12)
        date_to.grid(row=0, column=7, padx=5)
        
        ttk.Button(filter_frame, text="Today", command=lambda: self.set_date_range(0)).grid(
            row=0, column=8, padx=2)
        ttk.Button(filter_frame, text="Week", command=lambda: self.set_date_range(7)).grid(
            row=0, column=9, padx=2)
        ttk.Button(filter_frame, text="Month", command=lambda: self.set_date_range(30)).grid(
            row=0, column=10, padx=2)
        
        ttk.Button(filter_frame, text="Apply", command=self.apply_filters,
                  style='Accent.TButton').grid(row=0, column=11, padx=5)
        ttk.Button(filter_frame, text="Clear", command=self.clear_filters).grid(
            row=0, column=12, padx=5)
        
        tree_frame = ttk.Frame(main_frame)
        tree_frame.grid(row=2, column=0, columnspan=4, sticky="nsew", pady=10)
        
        columns = ('Code', 'Date', 'Type', 'Amount', 'Fee', 'Sender', 'Recipient', 'Balance')
        
        self.tree = ttk.Treeview(tree_frame, columns=columns, show='headings', height=20)
        
        for col in columns:
            self.tree.heading(col, text=col, 
                            command=lambda c=col: self.sort_column(c))
        
        self.tree.column('Code', width=120)
        self.tree.column('Date', width=150)
        self.tree.column('Type', width=100)
        self.tree.column('Amount', width=120, anchor='e')
        self.tree.column('Fee', width=80, anchor='e')
        self.tree.column('Sender', width=150)
        self.tree.column('Recipient', width=150)
        self.tree.column('Balance', width=120, anchor='e')
        
        vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(tree_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        
        self.tree.tag_configure('oddrow', background='#f0f0f0')
        self.tree.tag_configure('evenrow', background='white')
        
        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        
        tree_frame.columnconfigure(0, weight=1)
        tree_frame.rowconfigure(0, weight=1)
        
        self.context_menu = tk.Menu(self.tree, tearoff=0)
        self.context_menu.add_command(label="Copy Transaction Code", 
                                     command=self.copy_transaction_code)
        self.context_menu.add_command(label="Delete Transaction", 
                                     command=self.delete_transaction)
        self.context_menu.add_separator()
        self.context_menu.add_command(label="Export Selected", 
                                     command=self.export_selected)
        
        self.tree.bind("<Button-3>", self.show_context_menu)
        
        bottom_frame = ttk.Frame(main_frame)
        bottom_frame.grid(row=3, column=0, columnspan=4, pady=10, sticky="ew")
        
        self.stats_label = ttk.Label(bottom_frame, text="", font=('Arial', 10))
        self.stats_label.pack(side=tk.LEFT, padx=10)
        
        button_container = ttk.Frame(bottom_frame)
        button_container.pack(side=tk.RIGHT)
        
        ttk.Button(button_container, text="Refresh", command=self.load_data).pack(
            side=tk.LEFT, padx=5)
        ttk.Button(button_container, text="Export Filtered", 
                  command=self.export_filtered).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_container, text="Delete Selected", 
                  command=self.delete_transaction).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_container, text="Close", command=self.window.destroy).pack(
            side=tk.LEFT, padx=5)
        
        self.window.columnconfigure(0, weight=1)
        self.window.rowconfigure(0, weight=1)
        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(2, weight=1)
    
    def set_date_range(self, days_back):
        """Set date range for quick filters"""
        today = datetime.now()
        from_date = today - timedelta(days=days_back)
        
        self.date_from_var.set(from_date.strftime('%Y-%m-%d'))
        self.date_to_var.set(today.strftime('%Y-%m-%d'))
        self.apply_filters()
    
    def load_data(self):
        """Load all transactions"""
        self.all_transactions = self.db_manager.get_all_transactions()
        self.display_transactions(self.all_transactions)
    
    def display_transactions(self, transactions):
        """Display transactions in treeview"""
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        for idx, trans in enumerate(transactions):
            tag = 'evenrow' if idx % 2 == 0 else 'oddrow'
            
            self.tree.insert('', tk.END, values=(
                trans['transaction_code'],
                trans['date'],
                trans['transaction_type'].capitalize(),
                f"KSh {trans['amount']:,.2f}",
                f"KSh {trans['fee']:,.2f}",
                trans['sender'] or '-',
                trans['recipient'] or '-',
                f"KSh {trans['balance']:,.2f}" if trans['balance'] else '-'
            ), tags=(tag,))
        
        total_amount = sum(t['amount'] for t in transactions)
        total_fees = sum(t['fee'] for t in transactions)
        
        self.stats_label.config(
            text=f"Showing: {len(transactions)} transaction(s) | " +
                 f"Total: KSh {total_amount:,.2f} | Fees: KSh {total_fees:,.2f}")
    
    def apply_filters(self):
        """Apply search and filters"""
        filtered = self.all_transactions.copy()
        
        search_term = self.search_var.get().lower()
        if search_term:
            filtered = [t for t in filtered if 
                       search_term in t['transaction_code'].lower() or
                       search_term in str(t.get('sender', '')).lower() or
                       search_term in str(t.get('recipient', '')).lower()]
        
        trans_type = self.type_var.get()
        if trans_type != "All":
            filtered = [t for t in filtered if 
                       t['transaction_type'].lower() == trans_type.lower()]
        
        date_from = self.date_from_var.get()
        date_to = self.date_to_var.get()
        
        if date_from:
            filtered = [t for t in filtered if t['date'] >= date_from]
        
        if date_to:
            filtered = [t for t in filtered if t['date'] <= date_to + ' 23:59:59']
        
        self.display_transactions(filtered)
    
    def clear_filters(self):
        """Clear all filters"""
        self.search_var.set("")
        self.type_var.set("All")
        self.date_from_var.set("")
        self.date_to_var.set("")
        self.display_transactions(self.all_transactions)
    
    def sort_column(self, col):
        """Sort treeview by column with toggle reverse"""
        self.sort_reverse[col] = not self.sort_reverse.get(col, False)
        reverse = self.sort_reverse[col]
        
        items = [(self.tree.set(item, col), item) for item in self.tree.get_children('')]
        
        if col in ['Amount', 'Fee', 'Balance']:
            def get_numeric_value(val):
                if not val or val == '-':
                    return 0.0
                val_str = str(val)
                clean_val = val_str.replace('KSh', '').replace(',', '').replace(' ', '').strip()
                if clean_val == '-' or clean_val == '':
                    return 0.0
                try:
                    return float(clean_val)
                except (ValueError, AttributeError):
                    return 0.0
            
            items.sort(key=lambda x: get_numeric_value(x[0]), reverse=reverse)
        
        elif col == 'Date':
            def get_date_value(val):
                if not val:
                    return ''
                try:
                    return datetime.strptime(str(val), '%Y-%m-%d %H:%M:%S')
                except:
                    return str(val)
            
            items.sort(key=lambda x: get_date_value(x[0]), reverse=reverse)
        
        else:
            items.sort(key=lambda x: str(x[0]).lower() if x[0] else '', reverse=reverse)
        
        for index, (val, item) in enumerate(items):
            self.tree.move(item, '', index)
        
        for index, item in enumerate(self.tree.get_children('')):
            tag = 'evenrow' if index % 2 == 0 else 'oddrow'
            self.tree.item(item, tags=(tag,))
        
        for c in ['Code', 'Date', 'Type', 'Amount', 'Fee', 'Sender', 'Recipient', 'Balance']:
            if c == col:
                arrow = " v" if reverse else " ^"
                self.tree.heading(c, text=f"{c}{arrow}")
            else:
                self.tree.heading(c, text=c)
    
    def show_context_menu(self, event):
        """Show context menu on right-click"""
        item = self.tree.identify_row(event.y)
        if item:
            self.tree.selection_set(item)
            self.context_menu.post(event.x_root, event.y_root)
    
    def copy_transaction_code(self):
        """Copy transaction code to clipboard"""
        selection = self.tree.selection()
        if selection:
            item = selection[0]
            code = self.tree.item(item)['values'][0]
            self.window.clipboard_clear()
            self.window.clipboard_append(code)
            messagebox.showinfo("Copied", f"Transaction code '{code}' copied to clipboard")
    
    def delete_transaction(self):
        """Delete selected transaction"""
        selection = self.tree.selection()
        if not selection:
            messagebox.showwarning("Warning", "Please select a transaction to delete")
            return
        
        item = selection[0]
        code = self.tree.item(item)['values'][0]
        
        if messagebox.askyesno("Confirm Delete", 
                              f"Are you sure you want to delete transaction:\n{code}?"):
            if self.db_manager.delete_transaction(code):
                messagebox.showinfo("Success", "Transaction deleted successfully")
                self.load_data()
                if self.main_app:
                    self.main_app.load_statistics()
            else:
                messagebox.showerror("Error", "Failed to delete transaction")
    
    def export_filtered(self):
        """Export currently filtered transactions"""
        items = self.tree.get_children()
        if not items:
            messagebox.showwarning("Warning", "No transactions to export")
            return
        
        filename = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
            title="Export Filtered Transactions"
        )
        
        if filename:
            try:
                import pandas as pd
                
                data = []
                for item in items:
                    values = self.tree.item(item)['values']
                    data.append({
                        'Transaction Code': values[0],
                        'Date': values[1],
                        'Type': values[2],
                        'Amount': values[3],
                        'Fee': values[4],
                        'Sender': values[5],
                        'Recipient': values[6],
                        'Balance': values[7]
                    })
                
                df = pd.DataFrame(data)
                df.to_excel(filename, index=False)
                
                messagebox.showinfo("Success", 
                    f"Exported {len(data)} transaction(s) to:\n{filename}")
            except Exception as e:
                messagebox.showerror("Error", f"Export failed:\n{str(e)}")
    
    def export_selected(self):
        """Export only selected transactions"""
        selection = self.tree.selection()
        if not selection:
            messagebox.showwarning("Warning", "Please select transactions to export")
            return
        
        filename = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
            title="Export Selected Transactions"
        )
        
        if filename:
            try:
                import pandas as pd
                
                data = []
                for item in selection:
                    values = self.tree.item(item)['values']
                    data.append({
                        'Transaction Code': values[0],
                        'Date': values[1],
                        'Type': values[2],
                        'Amount': values[3],
                        'Fee': values[4],
                        'Sender': values[5],
                        'Recipient': values[6],
                        'Balance': values[7]
                    })
                
                df = pd.DataFrame(data)
                df.to_excel(filename, index=False)
                
                messagebox.showinfo("Success", 
                    f"Exported {len(data)} selected transaction(s) to:\n{filename}")
            except Exception as e:
                messagebox.showerror("Error", f"Export failed:\n{str(e)}")