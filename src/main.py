import imaplib
import email
from email.header import decode_header
from datetime import datetime

# Your email and app password
username = "jayanthidress@gmail.com"
password = "kqbv nxgy bgok fovc"

def clean_text(text):
    return "".join(filter(lambda x: x.isprintable(), text))

def extract_feedback(email_body):
    # Regex to find feedback or relevant keywords
    import re
    feedback_pattern = re.compile(r"(feedback|review|suggestion|comment)", re.IGNORECASE)
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
        return payload.decode('utf-8')
    except UnicodeDecodeError:
        try:
            return payload.decode('ISO-8859-1')  # Try a common alternative encoding
        except UnicodeDecodeError:
            return payload.decode('latin1')  # Fallback to 'latin1'

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

    # Process each unread email
    for email_id in email_ids:
        # Fetch the email by ID
        res, msg = imap.fetch(email_id, "(RFC822)")

        for response_part in msg:
            if isinstance(response_part, tuple):
                # Parse the email content
                msg = email.message_from_bytes(response_part[1])
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
                        if "attachment" not in content_disposition:
                            if content_type == "text/plain":
                                try:
                                    email_body = decode_payload(part)
                                    # Clean and extract feedback
                                    cleaned_body = clean_text(email_body)
                                    feedback = extract_feedback(cleaned_body)

                                    if feedback:  # Only print if feedback is found
                                        formatted_date = format_email_date(date_)
                                        print(f"From: {from_}")
                                        print(f"Subject: {subject}")
                                        print(f"Date: {formatted_date}")
                                        print(f"Feedback Extracted: {feedback}\n")
                                except Exception as e:
                                    print(f"Error decoding email: {e}")

                else:
                    # If the email isn't multipart (simple plain text)
                    try:
                        email_body = decode_payload(msg)
                        cleaned_body = clean_text(email_body)
                        feedback = extract_feedback(cleaned_body)

                        if feedback:  # Only print if feedback is found
                            formatted_date = format_email_date(date_)
                            print(f"From: {from_}")
                            print(f"Subject: {subject}")
                            print(f"Date: {formatted_date}")
                            print(f"Feedback Extracted: {feedback}\n")
                    except Exception as e:
                        print(f"Error decoding email: {e}")

        # Mark the email as read by adding the \Seen flag
        imap.store(email_id, '+FLAGS', '\Seen')

    # Close the connection and logout
    imap.close()
    imap.logout()

if __name__ == "__main__":
    get_unread_emails()
