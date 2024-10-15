import imaplib
import email
from email.header import decode_header
import re

# Your email and app password
username = "muhammadarshadra2@gmail.com"
password = "kklg hkou kcbo rrvd"

def clean_text(text):
    # Remove unwanted characters and decode
    return "".join(filter(lambda x: x.isprintable(), text))

def extract_feedback(email_body):
    # Regex to find feedback or relevant keywords
    feedback_pattern = re.compile(r"(feedback|review|suggestion|comment)", re.IGNORECASE)
    if feedback_pattern.search(email_body):
        return email_body.strip()  # Clean and return the feedback
    return "No feedback found."

def get_emails():
    # Connect to the Gmail IMAP server
    imap = imaplib.IMAP4_SSL("imap.gmail.com")
    
    # Login to the account
    imap.login(username, password)

    # Select the mailbox you want to extract emails from
    imap.select("inbox")

    # Search for all emails
    status, messages = imap.search(None, "ALL")
    email_ids = messages[0].split()

    # Process the latest email
    latest_email_id = email_ids[-1]

    # Fetch the email by ID
    res, msg = imap.fetch(latest_email_id, "(RFC822)")

    for response_part in msg:
        if isinstance(response_part, tuple):
            # Parse the email content
            msg = email.message_from_bytes(response_part[1])
            subject, encoding = decode_header(msg["Subject"])[0]

            if isinstance(subject, bytes):
                # Decode the subject if it's in bytes
                subject = subject.decode(encoding if encoding else "utf-8")

            # Extract the email sender
            from_ = msg.get("From")

            # If the email message is multipart
            if msg.is_multipart():
                for part in msg.walk():
                    content_type = part.get_content_type()
                    content_disposition = str(part.get("Content-Disposition"))

                    # Get the body of the email
                    if "attachment" not in content_disposition:
                        if content_type == "text/plain":
                            email_body = part.get_payload(decode=True).decode()

                            # Clean the email body and extract feedback
                            cleaned_body = clean_text(email_body)
                            feedback = extract_feedback(cleaned_body)
                            print(f"Subject: {subject}")
                            print(f"From: {from_}")
                            print(f"Feedback Extracted: {feedback}")

            else:
                # If not multipart, process plain text email
                content_type = msg.get_content_type()
                if content_type == "text/plain":
                    email_body = msg.get_payload(decode=True).decode()
                    cleaned_body = clean_text(email_body)
                    feedback = extract_feedback(cleaned_body)
                    print(f"Subject: {subject}")
                    print(f"From: {from_}")
                    print(f"Feedback Extracted: {feedback}")

    # Close the connection and logout
    imap.close()
    imap.logout()

if __name__ == "__main__":
    get_emails()
