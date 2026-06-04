"""
Email service for sending Outlook notifications when payments are marked as paid.

This implementation uses the local Outlook client via pywin32 (win32com).
The machine running this code must have Outlook installed and configured.
"""

from typing import Dict, Optional

try:
    import win32com.client  # type: ignore
except ImportError:  # pragma: no cover - import-time guard
    win32com = None  # type: ignore


def get_sales_rep_email_mapping() -> Dict[str, str]:
    """
    Map sales rep names to their email addresses.

    This is a simple in-memory mapping. In a production environment,
    you may want to load this data from a database or configuration file.
    """
    return {
        # Example:
        # "Alice Smith": "alice.smith@yourcompany.com",
        # "Bob Jones": "bob.jones@yourcompany.com",
    }


def resolve_sales_rep_email(sales_rep_name: str) -> Optional[str]:
    """
    Resolve the email address for a given sales rep name.

    Returns None if there is no configured email address.
    """
    mapping = get_sales_rep_email_mapping()
    return mapping.get(sales_rep_name)


def send_outlook_email(recipient: str, subject: str, body: str) -> None:
    """
    Send an email using the local Outlook client.

    Raises:
        RuntimeError if pywin32 is not available or Outlook is not installed.
    """
    if win32com is None:
        raise RuntimeError(
            "pywin32 is not installed. Install it to enable Outlook email sending."
        )

    outlook = win32com.client.Dispatch("Outlook.Application")
    mail = outlook.CreateItem(0)  # 0 = MailItem
    mail.To = recipient
    mail.Subject = subject
    mail.Body = body
    mail.Send()


def send_payment_notification_email(record: dict) -> Optional[str]:
    """
    Send a payment notification email for a specific compensation record.

    Args:
        record: Dictionary containing a single compensation record as
                returned from the database module.

    Returns:
        An optional warning message if no email could be sent; otherwise None.
    """
    sales_rep_name = record.get("sales_rep_name", "")
    recipient = resolve_sales_rep_email(sales_rep_name)
    if not recipient:
        return (
            f"No email mapping found for sales rep '{sales_rep_name}'. "
            "Payment marked as Paid but no email was sent."
        )

    subject = f"Compensation Payment Confirmed - Deal {record.get('deal_id')}"
    body_lines = [
        f"Hello {sales_rep_name},",
        "",
        "Your compensation payment has been marked as Paid.",
        "",
        f"Deal ID: {record.get('deal_id')}",
        f"Region: {record.get('region')}",
        f"Business Unit: {record.get('business_unit')}",
        f"Quarter: {record.get('quarter')}",
        "",
        f"Revenue: {record.get('revenue')}",
        f"Target: {record.get('target')}",
        f"Achievement: {record.get('achievement_percent')}%",
        f"Payout Rate: {record.get('payout_rate_percent')}%",
        f"Payout Amount: {record.get('payout_amount')}",
        "",
        "Best regards,",
        "Partner Compensation Team",
    ]
    body = "\n".join(body_lines)

    send_outlook_email(recipient=recipient, subject=subject, body=body)
    return None

