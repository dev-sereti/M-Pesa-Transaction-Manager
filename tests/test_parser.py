import unittest
from scr.parsers.mpesa_parser import MPesaParser
# from src.parsers.mpesa_parser import MPesaParser

class TestMPesaParser(unittest.TestCase):
    
    def setUp(self):
        self.parser = MPesaParser()
    
    def test_sent_money(self):
        message = """RBK1234567 Confirmed. Ksh500.00 sent to JOHN DOE 254712345678 
        on 15/1/24 at 2:30 PM. New M-PESA balance is Ksh2,500.00. 
        Transaction cost, Ksh15.00"""
        
        transaction = self.parser.parse_message(message)

        self.assertIsNotNone(transaction)
        assert transaction is not None  # type narrowing for Pylance
        self.assertEqual(transaction.transaction_code, 'RBK1234567')
        self.assertEqual(transaction.amount, 500.00)
        self.assertEqual(transaction.fee, 15.00)
        self.assertEqual(transaction.transaction_type, 'sent')
          
        # self.assertIsNotNone(transaction)
        # self.assertEqual(transaction.transaction_code, 'RBK1234567')
        # self.assertEqual(transaction.amount, 500.00)
        # self.assertEqual(transaction.fee, 15.00)
        # self.assertEqual(transaction.transaction_type, 'sent')
    
def test_received_money(self):
    message = """RBK9876543 Confirmed. You have received Ksh1,000.00 from 
    JANE SMITH 254723456789 on 15/1/24 at 3:45 PM. 
    New M-PESA balance is Ksh3,500.00"""
    
    transaction = self.parser.parse_message(message)
    
    self.assertIsNotNone(transaction)
    assert transaction is not None

    self.assertEqual(transaction.amount, 1000.00)
    self.assertEqual(transaction.transaction_type, 'received')

    
if __name__ == '__main__':
    unittest.main()