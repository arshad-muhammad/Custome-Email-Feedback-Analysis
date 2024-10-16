import imaplib
import email
from email.header import decode_header
from datetime import datetime
from textblob import TextBlob  # Import TextBlob for sentiment analysis
import re
from concurrent.futures import ThreadPoolExecutor

# Your email and app password
username = "jayanthidress@gmail.com"
password = "kqbv nxgy bgok fovc"

# Compile feedback pattern once (cached)
feedback_pattern = re.compile(r"(feedback|review|suggestion|comment|jayanthidress|dress)", re.IGNORECASE)

def clean_text(text):
    return "".join(filter(lambda x: x.isprintable(), text))

def extract_feedback(email_body):
    if feedback_pattern.search(email_body):
        return email_body.strip()  # Clean and return the feedback
    return None  # Return None if no feedback is found

def format_email_date(email_date):
    # Parse the date into a datetime object
    date_obj = email.utils.parsedate_to_datetime(email_date)
    # Format the date into dd/mm/yyyy
    return date_obj.strftime("%d/%m/%Y")

def decode_payload(part):
    try:
        payload = part.get_payload(decode=True)
        if isinstance(payload, bytes):
            for encoding in ['utf-8', 'ISO-8859-1', 'latin1']:
                try:
                    return payload.decode(encoding)
                except UnicodeDecodeError:
                    continue
        return str(payload)
    except Exception as e:
        return ""

def analyze_sentiment(feedback):
    # Analyze feedback sentiment using TextBlob
    analysis = TextBlob(feedback)
    polarity = analysis.sentiment.polarity  # Get polarity score: -1 (negative) to +1 (positive)
    
    if polarity > 0:
        return "Positive"
    elif polarity < 0:
        return "Negative"
    else:
        return "Neutral"

def process_email(email_data):
    # Parse the email content
    msg = email.message_from_bytes(email_data)
    subject, encoding = decode_header(msg["Subject"])[0]

    if isinstance(subject, bytes):
        subject = subject.decode(encoding if encoding else "utf-8")

    from_ = msg.get("From")
    date_ = msg.get("Date")

    # If the email message is multipart
    if msg.is_multipart():
        for part in msg.walk():
            content_type = part.get_content_type()
            content_disposition = str(part.get("Content-Disposition"))

            # Get the body of the email
            if "attachment" not in content_disposition and content_type == "text/plain":
                email_body = decode_payload(part)
                cleaned_body = clean_text(email_body)
                feedback = extract_feedback(cleaned_body)

                if feedback:  # Only process if feedback is found
                    formatted_date = format_email_date(date_)
                    sentiment = analyze_sentiment(feedback)  # Perform sentiment analysis
                    print(f"From: {from_}")
                    print(f"Subject: {subject}")
                    print(f"Date: {formatted_date}")
                    print(f"Feedback Extracted: {feedback}")
                    print(f"Sentiment: {sentiment}\n")

    else:
        # If the email isn't multipart (simple plain text)
        email_body = decode_payload(msg)
        cleaned_body = clean_text(email_body)
        feedback = extract_feedback(cleaned_body)

        if feedback:  # Only process if feedback is found
            formatted_date = format_email_date(date_)
            sentiment = analyze_sentiment(feedback)  # Perform sentiment analysis
            print(f"From: {from_}")
            print(f"Subject: {subject}")
            print(f"Date: {formatted_date}")
            print(f"Feedback Extracted: {feedback}")
            print(f"Sentiment: {sentiment}\n")

def get_unread_emails():
    # Connect to Gmail's IMAP server
    imap = imaplib.IMAP4_SSL("imap.gmail.com")

    # Login to the account
    imap.login(username, password)

    # Select the mailbox you want to read (inbox)
    imap.select("inbox")

    # Search for all unread emails
    status, messages = imap.search(None, "UNSEEN")

    # Get the list of email IDs (unread emails)
    email_ids = messages[0].split()

    if not email_ids:
        print("No new emails found.")
        return

    # Decode email IDs from bytes to strings
    email_ids = [email_id.decode() for email_id in email_ids]

    # Process emails in batches (fetch all in one go)
    max_emails_to_fetch = 50  # Adjust as necessary
    email_ids = email_ids[:max_emails_to_fetch]
    
    # Fetch all emails in a single request
    res, msgs = imap.fetch(",".join(email_ids), "(RFC822)")

    # Each response from the fetch can contain multiple parts, so loop through each one
    for i in range(0, len(msgs), 2):
        response = msgs[i]
        if isinstance(response, tuple):
            email_data = response[1]
            process_email(email_data)  # Process each email individually

    # Mark the emails as read by adding the \Seen flag to all processed emails
    imap.store(",".join(email_ids), '+FLAGS', '\Seen')

    # Close the connection and logout
    imap.close()
    imap.logout()

if __name__ == "__main__":
    get_unread_emails()
