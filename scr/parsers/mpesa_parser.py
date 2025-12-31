import re
from datetime import datetime
from typing import Optional, List

from scr.models.transaction import Transaction


class MPesaParser:

    PATTERNS = {
        "sent": {
            "code": r"([A-Z0-9]{10})\s+confirmed",
            "amount": r"Ksh\s*([\d,]+\.?\d*)",
            "date": r"on\s+(\d{1,2}/\d{1,2}/\d{2,4})\s+at\s+(\d{1,2}:\d{2}\s*(?:AM|PM)?)",
            "fee": r"Transaction cost[,:]?\s*Ksh\s*([\d,]+\.?\d*)",
            "recipient": r"sent to\s+(.+?)\s+on",
        },
        "received": {
            "code": r"([A-Z0-9]{10})\s+confirmed",
            "amount": r"Ksh\s*([\d,]+\.?\d*)",
            "date": r"on\s+(\d{1,2}/\d{1,2}/\d{2,4})\s+at\s+(\d{1,2}:\d{2}\s*(?:AM|PM)?)",
            "sender": r"from\s+(.+?)\s+on",
            # Received transactions typically do not show a "Transaction cost" line
            "fee": None,
        },
        "withdrawn": {
            "code": r"([A-Z0-9]{10})\s+confirmed",
            "amount": r"Ksh\s*([\d,]+\.?\d*)",
            "date": r"on\s+(\d{1,2}/\d{1,2}/\d{2,4})\s+at\s+(\d{1,2}:\d{2}\s*(?:AM|PM)?)",
            "fee": r"Transaction cost[,:]?\s*Ksh\s*([\d,]+\.?\d*)",
            "agent": r"from\s+(.+?)\s",  # M‑Pesa agent / outlet
        },
        "paybill": {
            "code": r"([A-Z0-9]{10})\s+confirmed",
            "amount": r"Ksh\s*([\d,]+\.?\d*)",
            "date": r"on\s+(\d{1,2}/\d{1,2}/\d{2,4})\s+at\s+(\d{1,2}:\d{2}\s*(?:AM|PM)?)",
            "fee": r"Transaction cost[,:]?\s*Ksh\s*([\d,]+\.?\d*)",
            "recipient": r"paid to\s+(.+?)(?:\.|on)",
            "account": r"Account\s+(.+?)\s*(?:on|\.|Transaction)",
        },
        "airtime": {
            "code": r"([A-Z0-9]{10})\s+confirmed",
            "amount": r"Ksh\s*([\d,]+\.?\d*)",
            "date": r"on\s+(\d{1,2}/\d{1,2}/\d{2,4})\s+at\s+(\d{1,2}:\d{2}\s*(?:AM|PM)?)",
            # Usually no explicit transaction cost line for airtime
            "fee": None,
        },
    }

    # Default fees when there is no explicit "Transaction cost" line.
    DEFAULT_FEES = {
        "received": 0.0,
        "airtime": 0.0,
        # Other types default to 0.0 if the fee pattern doesn't match
    }
    
    #                    Detection and normalisation                     #

    @staticmethod
    def detect_transaction_type(message: str) -> Optional[str]:
        """
        Try to detect the transaction type from the message text.
        Returns: 'sent', 'received', 'withdrawn', 'paybill', 'airtime', or None.
        """
        msg = message.lower()

        if "sent to" in msg:
            return "sent"
        if "received" in msg and "from" in msg:
            return "received"
        if "withdraw" in msg:
            return "withdrawn"
        if "paid to" in msg or "paybill" in msg:
            return "paybill"
        if "bought" in msg and "airtime" in msg:
            return "airtime"

        return None

    @staticmethod
    def clean_amount(amount_str: str) -> float:
        """Convert '1,234.50' to 1234.50 (float)."""
        return float(amount_str.replace(",", "").strip())

    @staticmethod
    def parse_date(date_str: str, time_str: str) -> datetime:
        """
        Parse date & time from M‑Pesa SMS text.
        Supports:
            - 2 or 4 digit years
            - 12‑hour with AM/PM
            - 24‑hour without AM/PM
        """
        time_norm = time_str.strip().upper()
        dt_str = f"{date_str.strip()} {time_norm}"

        formats = [
            "%d/%m/%y %I:%M %p",
            "%d/%m/%Y %I:%M %p",
            "%d/%m/%y %H:%M",
            "%d/%m/%Y %H:%M",
        ]

        last_error: Optional[Exception] = None
        for fmt in formats:
            try:
                return datetime.strptime(dt_str, fmt)
            except ValueError as exc:
                last_error = exc

        raise ValueError(
            f"Could not parse date/time '{dt_str}' with known formats. "
            f"Last error: {last_error}"
        )

    # ------------------------------------------------------------------ #
    #                          Regex utilities                           #
    # ------------------------------------------------------------------ #

    @staticmethod
    def _safe_search(pattern: Optional[str], message: str) -> Optional[re.Match]:
        """Safe wrapper around re.search (returns None for empty/None patterns)."""
        if not pattern:
            return None
        return re.search(pattern, message, re.IGNORECASE)

    def extract_field(
        self,
        message: str,
        pattern: Optional[str],
        group: int = 1,
    ) -> Optional[str]:
        """
        Extract a single field with a regex pattern.

        - If pattern is None/empty, returns None.
        - If the requested group does not exist, falls back to group(0).
        - Returns the matched string stripped, or None if no match.
        """
        if not pattern:
            return None

        match = self._safe_search(pattern, message)
        if not match:
            return None

        try:
            value = match.group(group)
        except IndexError:
            value = match.group(0)

        return value.strip() if isinstance(value, str) else str(value).strip()

    # ------------------------------------------------------------------ #
    #                         Public parse methods                       #
    # ------------------------------------------------------------------ #

    def parse_message(self, message: str) -> Optional[Transaction]:
        """
        Parse a single M‑Pesa SMS.
        Returns a Transaction on success, or None if parsing fails.
        """
        message = message.strip()
        if not message:
            return None

        try:
            # 1. Detect type
            trans_type = self.detect_transaction_type(message)
            if not trans_type or trans_type not in self.PATTERNS:
                raise ValueError(f"Unknown or unsupported transaction type: {trans_type!r}")

            patterns = self.PATTERNS[trans_type]

            # 2. Core fields: code, amount, date/time
            code = self.extract_field(message, patterns.get("code"))
            if not code:
                raise ValueError("Could not extract transaction code")

            amount_str = self.extract_field(message, patterns.get("amount"))
            if not amount_str:
                raise ValueError("Could not extract amount")
            amount = self.clean_amount(amount_str)

            date_match = self._safe_search(patterns.get("date"), message)
            if not date_match or len(date_match.groups()) < 2:
                raise ValueError("Could not extract date/time")

            date_str, time_str = date_match.group(1), date_match.group(2)
            transaction_date = self.parse_date(date_str, time_str)

            # 3. Fee
            fee_str = self.extract_field(message, patterns.get("fee"))
            if fee_str:
                fee = self.clean_amount(fee_str)
            else:
                fee = self.DEFAULT_FEES.get(trans_type, 0.0)

            # 4. Parties / extras
            recipient = self.extract_field(message, patterns.get("recipient"))
            sender = self.extract_field(message, patterns.get("sender"))

            # For withdrawals, use agent name as recipient (to keep the model consistent)
            if not recipient and trans_type == "withdrawn":
                recipient = self.extract_field(message, patterns.get("agent"))

            # (Account is defined in PATTERNS for paybill but not stored in Transaction yet.)

            # 5. Balance
            balance_match = re.search(
                r"balance.*?Ksh\s*([\d,]+\.?\d*)",
                message,
                re.IGNORECASE,
            )
            balance = (
                self.clean_amount(balance_match.group(1))
                if balance_match
                else None
            )

            # 6. Build Transaction object
            return Transaction(
                transaction_code=code,
                amount=amount,
                date=transaction_date,
                fee=fee,
                transaction_type=trans_type,
                recipient=recipient,
                sender=sender,
                balance=balance,
                message=message,
            )

        except Exception as exc:
            # Print to console so GUI can show a generic message but you can still debug.
            print(f"[MPesaParser] Error parsing message: {exc}\nMessage: {message}\n")
            return None

    def parse_multiple_messages(self, messages: List[str]) -> List[Transaction]:
        """
        Parse a list of SMS messages.
        Returns only successfully parsed Transaction objects.
        """
        transactions: List[Transaction] = []

        for raw_msg in messages:
            msg = raw_msg.strip()
            if not msg:
                continue

            tx = self.parse_message(msg)
            if tx is not None:
                transactions.append(tx)

        return transactions