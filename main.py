import imaplib
import email
from email.header import decode_header
from nltk.sentiment import SentimentIntensityAnalyzer
import re
import pandas as pd
import os

# Initialize VADER sentiment analyzer
# nltk.download('vader_lexicon')
vader_analyzer = SentimentIntensityAnalyzer()

# Email account credentials
EMAIL = 'email'
PASSWORD = 'app pass'


def extract_feedback(email_body):
    # Regex to find feedback or relevant keywords
    feedback_pattern = re.compile(r"(feedback|review|suggestion|comment|dress)", re.IGNORECASE)
    if feedback_pattern.search(email_body):
        return email_body.strip()  # Clean and return the feedback
    else:
         return "No feedback found."


def extract_customer_name(email_body):
    """Extract customer name from the email body."""
    name_patterns = [
        r"my name is\s+([A-Za-z]+)",    # Matches "my name is <Name>"
        r"i'?m\s+([A-Za-z]+)",          # Matches "I'm <Name>"
        r"this is\s+([A-Za-z]+)",       # Matches "this is <Name>"
        r"here'?s\s+([A-Za-z]+)",       # Matches "here's <Name>"
        r"\bi am\s+([A-Za-z]+)",        # Matches "I am <Name>"
        r"\bi'?m\s+([A-Za-z]+)"         # Matches variations like "I’m Arshad"
    ]

    for pattern in name_patterns:
        match = re.search(pattern, email_body, re.IGNORECASE)
        if match:
            return match.group(1).strip()
        else:
            return "Not available"  # Default if no name is found


def extract_order_id(email_body):
    """Extract order ID from the email body."""
    order_id_patterns = [
        r"order id is\s*([A-Za-z0-9\-]+)",    # Matches "order id is <OrderID>"
        r"my order id\s*[:\s]*([A-Za-z0-9\-]+)",  # Matches "my order id: <OrderID>"
        r"order #?\s*([A-Za-z0-9\-]+)",  # Matches "order # <OrderID>"
        r"order number\s*[:\s]*([A-Za-z0-9\-]+)",  # Matches "order number: <OrderID>"
    ]

    for pattern in order_id_patterns:
        match = re.search(pattern, email_body, re.IGNORECASE)
        if match:
            return match.group(1).strip()
        else:
          return "Not available"  # Default if no order ID is found


def analyze_sentiment(feedback):
    """Analyze feedback sentiment using VADER."""
    vader_scores = vader_analyzer.polarity_scores(feedback)
    compound_score = vader_scores['compound']

    negative_keywords = ['terrible', 'awful', 'bad', 'worst', 'horrible', 'poor', 'disappointed', 'hate']
    positive_keywords = ['excellent', 'great', 'amazing', 'fantastic', 'good', 'wonderful', 'love', 'best']

    if any(neg_word in feedback.lower() for neg_word in negative_keywords):
        return "Negative"
    if any(pos_word in feedback.lower() for pos_word in positive_keywords):
        return "Positive"

    if compound_score >= 0.3:  # Adjusted threshold for positive
        return "Positive"
    elif compound_score <= -0.3:  # Adjusted threshold for negative
        return "Negative"
    else:
        return "Neutral"


def get_unread_emails():
    """Fetch unread emails and extract feedback and customer details."""
    # Connect to the IMAP server and login
    imap = imaplib.IMAP4_SSL('imap.gmail.com')
    imap.login(EMAIL, PASSWORD)

    # Select the mailbox you want to check (INBOX)
    imap.select("inbox")

    # Search for all unread emails
    status, messages = imap.search(None, 'UNSEEN')

    # Convert the result into a list of email IDs
    email_ids = messages[0].split()

    # Check if no unread emails were found
    if not email_ids:
        print("No new emails found.")
        return

    print(f"Found {len(email_ids)} new email(s).")

    # Prepare a list to store extracted data
    extracted_data = []

    for email_id in email_ids:
        res, msg = imap.fetch(email_id, "(RFC822)")

        for response_part in msg:
            if isinstance(response_part, tuple):
                # Parse the email content
                msg = email.message_from_bytes(response_part[1])
                email_subject = decode_header(msg["Subject"])[0][0]
                if isinstance(email_subject, bytes):
                    email_subject = email_subject.decode()

                email_from = decode_header(msg.get("From"))[0][0]
                if isinstance(email_from, bytes):
                    email_from = email_from.decode()

                date_received = msg["Date"]
                date_object = email.utils.parsedate_to_datetime(date_received)
                formatted_date = date_object.strftime("%d/%m/%Y")

                # Extract email body
                email_body = ""
                if msg.is_multipart():
                    for part in msg.walk():
                        if part.get_content_type() == "text/plain":
                            try:
                                email_body = part.get_payload(decode=True).decode()
                            except UnicodeDecodeError:
                                email_body = part.get_payload(decode=True).decode('latin-1')
                else:
                    email_body = msg.get_payload(decode=True).decode()

                # Extract feedback, customer name, and order ID
                feedback = extract_feedback(email_body)
                customer_name = None
                order_id = None
                
                if feedback:  # Only extract name and order ID if feedback is present
                    customer_name = extract_customer_name(email_body)
                    order_id = extract_order_id(email_body)

                if feedback and customer_name != "Not available" and order_id != "Not available":
                    print(f"From: {email_from}")
                    print(f"Subject: {email_subject}")
                    print(f"Date: {formatted_date}")
                    print(f"Customer Name: {customer_name}")
                    print(f"Order ID: {order_id}")
                    print(f"Feedback Extracted: {feedback}")

                    sentiment = analyze_sentiment(feedback)
                    print(f"Sentiment: {sentiment}")

                    # Append the extracted data to the list
                    extracted_data.append({
                        'Email From': email_from,
                        'Subject': email_subject,
                        'Date': formatted_date,
                        'Customer Name': customer_name,
                        'Order ID': order_id,
                        'Feedback': feedback,
                        'Sentiment': sentiment
                    })
                else:
                    if not feedback:
                        print(f"No feedback found in email from {email_from}.")
                    if customer_name == "Not available":
                        print(f"No customer name found in email from {email_from}.")
                    if order_id == "Not available":
                        print(f"No order ID found in email from {email_from}.")

                print("=" * 50)

    # Close the connection
    imap.close()
    imap.logout()

    # If we have extracted data, write it to an Excel file
    if extracted_data:
        # Define the directory to save the Excel file
        save_dir = os.path.expanduser("~/Documents/customer-feedback-analysis")
        os.makedirs(save_dir, exist_ok=True)  # Create the directory if it doesn't exist

        # Path for the Excel files
        excel_file_path = os.path.join(save_dir, "extracted_feedback.xlsx")
        
        # Create a DataFrame and save it to Excel
        df = pd.DataFrame(extracted_data)
        df.to_excel(excel_file_path, index=False)
        print(f"Extracted details have been saved to '{excel_file_path}'.")


# Call the function to get unread emails and extract feedback
get_unread_emails()
