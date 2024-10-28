import imaplib
import email
from email.header import decode_header
from email.message import EmailMessage
from email.utils import parsedate_to_datetime
from nltk.sentiment import SentimentIntensityAnalyzer
import re
import pandas as pd
import os
import matplotlib.pyplot as plt
from io import BytesIO
import smtplib
from dataclasses import dataclass
from typing import List, Optional
from datetime import datetime

@dataclass
class EmailConfig:
    """Email configuration settings"""
    email: str
    password: str
    service_email: str
    imap_server: str = 'imap.gmail.com'
    smtp_server: str = 'smtp.gmail.com'
    smtp_port: int = 587

@dataclass
class FeedbackData:
    """Structure for storing extracted feedback data"""
    email_from: str
    subject: str
    date: str
    customer_name: str
    order_id: str
    feedback: str
    sentiment: str

class FeedbackAnalyzer:
    """Main class for analyzing customer feedback from emails"""
    
    def __init__(self, config: EmailConfig):
        self.config = config
        self.vader_analyzer = SentimentIntensityAnalyzer()
        self.feedback_pattern = re.compile(r"(feedback|review|suggestion|comment|dress)", re.IGNORECASE)
        self.name_patterns = [
            r"my name is\s+([A-Za-z]+)",
            r"i'?m\s+([A-Za-z]+)",
            r"this is\s+([A-Za-z]+)",
            r"here'?s\s+([A-Za-z]+)",
            r"\bi am\s+([A-Za-z]+)"
        ]
        self.order_patterns = [
            r"order id is\s*([A-Za-z0-9\-]+)",
            r"my order id\s*[:\s]*([A-Za-z0-9\-]+)",
            r"order #?\s*([A-Za-z0-9\-]+)",
            r"order number\s*[:\s]*([A-Za-z0-9\-]+)"
        ]

    def extract_feedback(self, email_body: str) -> str:
        """Extract feedback from email body"""
        if self.feedback_pattern.search(email_body):
            return email_body.strip()
        return "No feedback found."

    def extract_customer_name(self, email_body: str) -> str:
        """Extract customer name from email body"""
        for pattern in self.name_patterns:
            if match := re.search(pattern, email_body, re.IGNORECASE):
                return match.group(1).strip()
        return "Not available"

    def extract_order_id(self, email_body: str) -> str:
        """Extract order ID from email body"""
        for pattern in self.order_patterns:
            if match := re.search(pattern, email_body, re.IGNORECASE):
                return match.group(1).strip()
        return "Not available"

    def analyze_sentiment(self, feedback: str) -> str:
        """Analyze sentiment of feedback text"""
        vader_scores = self.vader_analyzer.polarity_scores(feedback)
        compound_score = vader_scores['compound']

        negative_keywords = {'terrible', 'awful', 'bad', 'worst', 'horrible', 'poor', 'disappointed', 'hate'}
        positive_keywords = {'excellent', 'great', 'amazing', 'fantastic', 'good', 'wonderful', 'love', 'best'}

        feedback_lower = feedback.lower()
        if any(word in feedback_lower for word in negative_keywords):
            return "Negative"
        if any(word in feedback_lower for word in positive_keywords):
            return "Positive"

        return "Positive" if compound_score >= 0.3 else "Negative" if compound_score <= -0.3 else "Neutral"

    def create_sentiment_chart(self, sentiment_data: pd.Series) -> BytesIO:
        """Create a chart visualizing sentiment distribution"""
        plt.figure(figsize=(8, 6))
        sentiment_counts = sentiment_data.value_counts()
        colors = {'Positive': 'green', 'Negative': 'red', 'Neutral': 'grey'}
        sentiment_counts.plot(kind='bar', color=[colors[x] for x in sentiment_counts.index])
        plt.title('Sentiment Analysis Summary')
        plt.xlabel('Sentiment')
        plt.ylabel('Count')
        plt.xticks(rotation=0)
        plt.tight_layout()

        img_buffer = BytesIO()
        plt.savefig(img_buffer, format='png')
        img_buffer.seek(0)
        plt.close()
        return img_buffer

    def send_summary_email(self, feedback_data: List[FeedbackData], sentiment_scores: pd.Series):
        """Send summary email with feedback analysis"""
        msg = EmailMessage()
        msg['From'] = self.config.email
        msg['To'] = self.config.service_email
        msg['Subject'] = "Customer Feedback Summary"

        body = "Customer Feedback Summary Report\n\n"
        for data in feedback_data:
            body += (f"From: {data.email_from}\n"
                    f"Subject: {data.subject}\n"
                    f"Date: {data.date}\n"
                    f"Customer Name: {data.customer_name}\n"
                    f"Order ID: {data.order_id}\n"
                    f"Feedback: {data.feedback}\n"
                    f"Sentiment: {data.sentiment}\n"
                    f"{'='*50}\n")

        msg.set_content(body)

        # Add sentiment chart
        chart_data = self.create_sentiment_chart(sentiment_scores)
        msg.add_attachment(chart_data.read(), maintype='image', 
                         subtype='png', filename='sentiment_summary.png')

        with smtplib.SMTP(self.config.smtp_server, self.config.smtp_port) as server:
            server.starttls()
            server.login(self.config.email, self.config.password)
            server.send_message(msg)

    def save_feedback_data(self, feedback_data: List[FeedbackData]):
        """Save feedback data to Excel file"""
        save_dir = os.path.expanduser("~/Documents/customer-feedback-analysis")
        os.makedirs(save_dir, exist_ok=True)
        file_path = os.path.join(save_dir, "extracted_feedback.xlsx")

        df = pd.DataFrame([vars(data) for data in feedback_data])
        df.to_excel(file_path, index=False)
        print(f"Feedback data saved to '{file_path}'")

    def process_email_body(self, msg) -> Optional[str]:
        """Extract and decode email body"""
        if msg.is_multipart():
            for part in msg.walk():
                if part.get_content_type() == "text/plain":
                    try:
                        return part.get_payload(decode=True).decode()
                    except UnicodeDecodeError:
                        return part.get_payload(decode=True).decode('latin-1')
        else:
            return msg.get_payload(decode=True).decode()
        return None

    def process_emails(self):
        """Main method to process unread emails and extract feedback"""
        with imaplib.IMAP4_SSL(self.config.imap_server) as imap:
            imap.login(self.config.email, self.config.password)
            imap.select("inbox")
            
            _, messages = imap.search(None, 'UNSEEN')
            email_ids = messages[0].split()

            if not email_ids:
                print("No new emails found.")
                return

            print(f"Processing {len(email_ids)} new email(s)...")
            
            feedback_data = []
            sentiment_scores = []

            for email_id in email_ids:
                _, msg_data = imap.fetch(email_id, "(RFC822)")
                
                for response_part in msg_data:
                    if not isinstance(response_part, tuple):
                        continue

                    msg = email.message_from_bytes(response_part[1])
                    
                    # Extract email metadata
                    subject = decode_header(msg["Subject"])[0][0]
                    subject = subject.decode() if isinstance(subject, bytes) else subject
                    
                    sender = decode_header(msg.get("From"))[0][0]
                    sender = sender.decode() if isinstance(sender, bytes) else sender
                    
                    date = parsedate_to_datetime(msg["Date"])
                    formatted_date = date.strftime("%d/%m/%Y")

                    # Process email body
                    if email_body := self.process_email_body(msg):
                        feedback = self.extract_feedback(email_body)
                        customer_name = self.extract_customer_name(email_body)
                        order_id = self.extract_order_id(email_body)

                        if (feedback != "No feedback found." and 
                            customer_name != "Not available" and 
                            order_id != "Not available"):
                            
                            sentiment = self.analyze_sentiment(feedback)
                            sentiment_scores.append(sentiment)

                            feedback_data.append(FeedbackData(
                                email_from=sender,
                                subject=subject,
                                date=formatted_date,
                                customer_name=customer_name,
                                order_id=order_id,
                                feedback=feedback,
                                sentiment=sentiment
                            ))

                            print(f"Processed feedback from {sender}")
                        else:
                            print(f"Incomplete feedback data in email from {sender}")

            if feedback_data:
                self.save_feedback_data(feedback_data)
                self.send_summary_email(feedback_data, pd.Series(sentiment_scores))
                print("Analysis complete. Summary email sent.")
            else:
                print("No valid feedback found in the processed emails.")

def main():
    """Main entry point of the script"""
    config = EmailConfig(
        email='jayanthidress@gmail.com',
        password='kqbv nxgy bgok fovc',
        service_email='muhd.arshad@gmail.com'
    )
    
    analyzer = FeedbackAnalyzer(config)
    analyzer.process_emails()

if __name__ == "__main__":
    main()
