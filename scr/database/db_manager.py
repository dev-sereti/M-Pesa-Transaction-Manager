import sqlite3
from typing import List, Optional
from contextlib import contextmanager
from scr.models.transaction import Transaction

class DatabaseManager:
    """Manage SQLite database for transactions"""
    
    def __init__(self, db_path: str = "mpesa_transactions.db"):
        self.db_path = db_path
        self.create_tables()
    
    @contextmanager
    def get_connection(self):
        """Context manager for database connections"""
        conn = sqlite3.connect(self.db_path)
        conn.row_factory = sqlite3.Row
        try:
            yield conn
            conn.commit()
        except Exception as e:
            conn.rollback()
            raise e
        finally:
            conn.close()
    
    def create_tables(self):
        """Create database tables if they don't exist"""
        with self.get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS transactions (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    transaction_code TEXT UNIQUE NOT NULL,
                    amount REAL NOT NULL,
                    date TEXT NOT NULL,
                    fee REAL NOT NULL,
                    transaction_type TEXT NOT NULL,
                    recipient TEXT,
                    sender TEXT,
                    balance REAL,
                    message TEXT,
                    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                )
            ''')
            
            # Create indexes for better query performance
            cursor.execute('''
                CREATE INDEX IF NOT EXISTS idx_transaction_code 
                ON transactions(transaction_code)
            ''')
            cursor.execute('''
                CREATE INDEX IF NOT EXISTS idx_date 
                ON transactions(date)
            ''')
    
    def insert_transaction(self, transaction: Transaction) -> bool:
        """Insert a single transaction"""
        try:
            with self.get_connection() as conn:
                cursor = conn.cursor()
                cursor.execute('''
                    INSERT INTO transactions 
                    (transaction_code, amount, date, fee, transaction_type, 
                     recipient, sender, balance, message)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                ''', (
                    transaction.transaction_code,
                    transaction.amount,
                    transaction.date.strftime('%Y-%m-%d %H:%M:%S'),
                    transaction.fee,
                    transaction.transaction_type,
                    transaction.recipient,
                    transaction.sender,
                    transaction.balance,
                    transaction.message
                ))
                return True
        except sqlite3.IntegrityError:
            print(f"Transaction {transaction.transaction_code} already exists")
            return False
        except Exception as e:
            print(f"Error inserting transaction: {e}")
            return False
    
    def insert_multiple_transactions(self, transactions: List[Transaction]) -> tuple:
        """Insert multiple transactions and return (success_count, duplicate_count)"""
        success = 0
        duplicates = 0
        
        for transaction in transactions:
            if self.insert_transaction(transaction):
                success += 1
            else:
                duplicates += 1
        
        return success, duplicates
    
    def get_all_transactions(self) -> List[dict]:
        """Retrieve all transactions"""
        with self.get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute('SELECT * FROM transactions ORDER BY date DESC')
            return [dict(row) for row in cursor.fetchall()]
    
    def get_transactions_by_date_range(self, start_date: str, end_date: str) -> List[dict]:
        """Get transactions within a date range"""
        with self.get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute('''
                SELECT * FROM transactions 
                WHERE date BETWEEN ? AND ?
                ORDER BY date DESC
            ''', (start_date, end_date))
            return [dict(row) for row in cursor.fetchall()]
    
    def get_transaction_by_code(self, code: str) -> Optional[dict]:
        """Get a specific transaction by code"""
        with self.get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute('SELECT * FROM transactions WHERE transaction_code = ?', (code,))
            row = cursor.fetchone()
            return dict(row) if row else None
    
    def delete_transaction(self, transaction_code: str) -> bool:
        """Delete a transaction by code"""
        try:
            with self.get_connection() as conn:
                cursor = conn.cursor()
                cursor.execute('DELETE FROM transactions WHERE transaction_code = ?', 
                             (transaction_code,))
                return cursor.rowcount > 0
        except Exception as e:
            print(f"Error deleting transaction: {e}")
            return False
    
    def get_statistics(self) -> dict:
        """Get transaction statistics"""
        with self.get_connection() as conn:
            cursor = conn.cursor()
            
            # Total transactions
            cursor.execute('SELECT COUNT(*) as total FROM transactions')
            total = cursor.fetchone()['total']
            
            # Total amount
            cursor.execute('SELECT SUM(amount) as total_amount FROM transactions')
            total_amount = cursor.fetchone()['total_amount'] or 0
            
            # Total fees
            cursor.execute('SELECT SUM(fee) as total_fees FROM transactions')
            total_fees = cursor.fetchone()['total_fees'] or 0
            
            # By transaction type
            cursor.execute('''
                SELECT transaction_type, COUNT(*) as count, SUM(amount) as total
                FROM transactions
                GROUP BY transaction_type
            ''')
            by_type = [dict(row) for row in cursor.fetchall()]
            
            return {
                'total_transactions': total,
                'total_amount': total_amount,
                'total_fees': total_fees,
                'by_type': by_type
            }