from dataclasses import dataclass
from datetime import  datetime
from typing import Optional

@dataclass
class Transaction:
    """M-Pesa transaction data model"""
    transaction_code: str
    amount: float
    date: datetime
    fee: float
    transaction_type: str  # Sent, Received, Withdrawn, etc.
    recipient: Optional[str] = None
    sender: Optional[str] = None
    balance: Optional[float] = None
    message: Optional[str] = None

    def to_dict(self):
        return{
            'transaction_code': self.transaction_code,
            'amount': self.amount,
            'date': self.date.strftime('%Y-%m-%d %H:%M:%S'),
            'fee': self.fee,
            'transaction_type': self.transaction_type,
            'recipient': self.recipient,
            'sender': self.sender,
            'balance': self.balance,
            'message': self.message
        }