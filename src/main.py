import imaplib
import email
from email.header import decode_header
from nltk.sentiment import SentimentIntensityAnalyzer
import re
import pandas as pd
import os
import matplotlib.pyplot as plt
import smtplib
from io import BytesIO

# Initialize VADER sentiment analyzer
vader_analyzer = SentimentIntensityAnalyzer()

# Email account credentials
EMAIL = 'jayanthidress@gmail.com'
PASSWORD = 'kqbv nxgy bgok fovc'
SERVICE_EMAIL = 'muhd.arshad@gmail.com'


def extract_feedback(email_body):
    feedback_pattern = re.compile(r"(feedback|review|suggestion|comment|dress)", re.IGNORECASE)
    if feedback_pattern.search(email_body):
        return email_body.strip()  # Clean and return the feedback
    else:
        return "No feedback found."


def extract_customer_name(email_body):
    name_patterns = [
        r"my name is\s+([A-Za-z]+)",
        r"i'?m\s+([A-Za-z]+)",
        r"this is\s+([A-Za-z]+)",
        r"here'?s\s+([A-Za-z]+)",
        r"\bi am\s+([A-Za-z]+)",
        r"\bi'?m\s+([A-Za-z]+)"
    ]

    for pattern in name_patterns:
        match = re.search(pattern, email_body, re.IGNORECASE)
        if match:
            return match.group(1).strip()

    return "Not available"


def extract_order_id(email_body):
    order_id_patterns = [
        r"order id is\s*([A-Za-z0-9\-]+)",
        r"my order id\s*[:\s]*([A-Za-z0-9\-]+)",
        r"order #?\s*([A-Za-z0-9\-]+)",
        r"order number\s*[:\s]*([A-Za-z0-9\-]+)",
    ]

    for pattern in order_id_patterns:
        match = re.search(pattern, email_body, re.IGNORECASE)
        if match:
            return match.group(1).strip()

    return "Not available"


def analyze_sentiment(feedback):
    vader_scores = vader_analyzer.polarity_scores(feedback)
    compound_score = vader_scores['compound']

    negative_keywords = ['terrible', 'awful', 'bad', 'worst', 'horrible', 'poor', 'disappointed', 'hate']
    positive_keywords = ['excellent', 'great', 'amazing', 'fantastic', 'good', 'wonderful', 'love', 'best']

    if any(neg_word in feedback.lower() for neg_word in negative_keywords):
        return "Negative"
    if any(pos_word in feedback.lower() for pos_word in positive_keywords):
        return "Positive"

    if compound_score >= 0.3:
        return "Positive"
    elif compound_score <= -0.3:
        return "Negative"
    else:
        return "Neutral"


def plot_sentiment_chart(sentiment_data):
    # Count sentiment occurrences
    sentiment_counts = sentiment_data.value_counts()
    
    plt.figure(figsize=(8, 6))
    sentiment_counts.plot(kind='bar', color=['green', 'red', 'grey'])
    plt.title('Sentiment Analysis Summary')
    plt.xlabel('Sentiment')
    plt.ylabel('Count')
    plt.xticks(rotation=0)
    plt.tight_layout()

    # Save chart to a BytesIO object
    img_data = BytesIO()
    plt.savefig(img_data, format='png')
    img_data.seek(0)  # Seek to the start of the stream
    plt.close()  # Close the plot
    return img_data


def send_summary_email(extracted_data, sentiment_scores):
    """Send a summary email to the customer service team."""
    # Prepare the email content
    subject = "Customer Feedback Summary"
    body = "Here is the summary of customer feedback:\n\n"
    for data in extracted_data:
        body += f"From: {data['Email From']}\n"
        body += f"Subject: {data['Subject']}\n"
        body += f"Date: {data['Date']}\n"
        body += f"Customer Name: {data['Customer Name']}\n"
        body += f"Order ID: {data['Order ID']}\n"
        body += f"Feedback: {data['Feedback']}\n"
        body += f"Sentiment: {data['Sentiment']}\n"
        body += "=" * 50 + "\n"
    
    # Chart for sentiment summary
    img_data = plot_sentiment_chart(sentiment_scores)

    # Sending the email
    msg = email.message.EmailMessage()
    msg['From'] = EMAIL
    msg['To'] = SERVICE_EMAIL
    msg['Subject'] = subject
    msg.set_content(body)

    # Attach the image to the email
    msg.add_attachment(img_data.read(), maintype='image', subtype='png', filename='sentiment_summary.png')

    with smtplib.SMTP('smtp.gmail.com', 587) as server:
        server.starttls()
        server.login(EMAIL, PASSWORD)
        server.send_message(msg)

    print(f"Summary email sent to {SERVICE_EMAIL}.")


def get_unread_emails():
    imap = imaplib.IMAP4_SSL('imap.gmail.com')
    imap.login(EMAIL, PASSWORD)

    imap.select("inbox")
    status, messages = imap.search(None, 'UNSEEN')
    email_ids = messages[0].split()

    if not email_ids:
        print("No new emails found.")
        return

    print(f"Found {len(email_ids)} new email(s).")

    extracted_data = []
    sentiment_scores = []  # To hold sentiment results

    for email_id in email_ids:
        res, msg = imap.fetch(email_id, "(RFC822)")

        for response_part in msg:
            if isinstance(response_part, tuple):
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

                feedback = extract_feedback(email_body)
                customer_name = None
                order_id = None
                
                if feedback:
                    customer_name = extract_customer_name(email_body)
                    order_id = extract_order_id(email_body)

                if feedback and customer_name != "Not available" and order_id != "Not available":
                    sentiment = analyze_sentiment(feedback)
                    sentiment_scores.append(sentiment)  # Collect sentiment scores

                    extracted_data.append({
                        'Email From': email_from,
                        'Subject': email_subject,
                        'Date': formatted_date,
                        'Customer Name': customer_name,
                        'Order ID': order_id,
                        'Feedback': feedback,
                        'Sentiment': sentiment
                    })

                    print(f"From: {email_from}")
                    print(f"Subject: {email_subject}")
                    print(f"Date: {formatted_date}")
                    print(f"Customer Name: {customer_name}")
                    print(f"Order ID: {order_id}")
                    print(f"Feedback Extracted: {feedback}")
                    print(f"Sentiment: {sentiment}")
                    print("=" * 50)

                else:
                    if not feedback:
                        print(f"No feedback found in email from {email_from}.")
                    if customer_name == "Not available":
                        print(f"No customer name found in email from {email_from}.")
                    if order_id == "Not available":
                        print(f"No order ID found in email from {email_from}.")

    imap.close()
    imap.logout()

    if extracted_data:
        save_dir = os.path.expanduser("~/Documents/customer-feedback-analysis")
        os.makedirs(save_dir, exist_ok=True)
        excel_file_path = os.path.join(save_dir, "extracted_feedback.xlsx")
        
        df = pd.DataFrame(extracted_data)
        df.to_excel(excel_file_path, index=False)
        print(f"Extracted details has been saved to '{excel_file_path}'.")

        # Send summary email with the collected data
        send_summary_email(extracted_data, pd.Series(sentiment_scores))


# Call the functions to get unread emails and extract feedback
get_unread_emails()
