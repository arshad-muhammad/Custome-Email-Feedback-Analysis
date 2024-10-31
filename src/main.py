import imaplib
import email
from email.header import decode_header
from email.message import EmailMessage
from email.utils import parsedate_to_datetime
import nltk
import re
import pandas as pd
import os
import logging
from logging.handlers import RotatingFileHandler
import matplotlib.pyplot as plt
from io import BytesIO
import smtplib
from dataclasses import dataclass, asdict
from typing import List, Optional, Dict, Any
from datetime import datetime, timedelta
import json
import keyring
import ssl

# Download NLTK resources if not already present
try:
    nltk.download('vader_lexicon', quiet=True)
except Exception:
    pass
from nltk.sentiment import SentimentIntensityAnalyzer

class ConfigurationError(Exception):
    """Custom exception for configuration-related errors."""
    pass

class EmailProcessingError(Exception):
    """Custom exception for email processing errors."""
    pass

@dataclass
class EmailConfig:
    """Enhanced email configuration with more robust settings"""
    email: str
    service_email: str
    imap_server: str = 'imap.gmail.com'
    smtp_server: str = 'smtp.gmail.com'
    smtp_port: int = 587
    log_dir: str = os.path.expanduser("~/Documents/feedback-analyzer-logs")
    output_dir: str = os.path.expanduser("~/Documents/customer-feedback-analysis")

@dataclass
class FeedbackData:
    """Structured feedback data with additional validation"""
    email_from: str
    subject: str
    date: str
    customer_name: str
    order_id: str
    feedback: str
    sentiment: str

class EnhancedFeedbackAnalyzer:
    """Advanced feedback analysis with robust error handling and logging"""
    
    def __init__(self, config: EmailConfig):
        # Setup directories
        os.makedirs(config.log_dir, exist_ok=True)
        os.makedirs(config.output_dir, exist_ok=True)

        # Configure logging
        self.logger = self._setup_logging(config.log_dir)
        self.config = config

        # Enhanced sentiment analysis
        try:
            self.vader_analyzer = SentimentIntensityAnalyzer()
        except Exception as e:
            self.logger.error(f"Failed to initialize sentiment analyzer: {e}")
            raise ConfigurationError("Sentiment analysis setup failed")

        # Compile regex patterns
        self.feedback_pattern = re.compile(
            r"(feedback|review|suggestion|comment|dress)", 
            re.IGNORECASE
        )
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

    def _setup_logging(self, log_dir: str) -> logging.Logger:
        """Configure comprehensive logging"""
        logger = logging.getLogger('FeedbackAnalyzer')
        logger.setLevel(logging.INFO)

        # File handler
        file_handler = RotatingFileHandler(
            os.path.join(log_dir, 'feedback_analyzer.log'),
            maxBytes=10*1024*1024,  # 10 MB
            backupCount=5
        )
        file_handler.setLevel(logging.INFO)

        # Console handler
        console_handler = logging.StreamHandler()
        console_handler.setLevel(logging.INFO)

        # Formatter
        formatter = logging.Formatter(
            '%(asctime)s - %(name)s - %(levelname)s - %(message)s'
        )
        file_handler.setFormatter(formatter)
        console_handler.setFormatter(formatter)

        logger.addHandler(file_handler)
        logger.addHandler(console_handler)

        return logger

    def _secure_store_password(self, service: str, username: str, password: str):
        """Securely store password using system keyring"""
        try:
            keyring.set_password(service, username, password)
        except Exception as e:
            self.logger.error(f"Password storage failed: {e}")

    def _retrieve_password(self, service: str, username: str) -> str:
        """Retrieve password from system keyring"""
        try:
            password = keyring.get_password(service, username)
            if not password:
                raise ConfigurationError("No stored password found")
            return password
        except Exception as e:
            self.logger.error(f"Password retrieval failed: {e}")
            raise

    def process_emails(self):
        """Robust email processing with comprehensive error handling"""
        try:
            # Retrieve password securely
            password = self._retrieve_password('Gmail', self.config.email)

            # Secure SSL context
            ssl_context = ssl.create_default_context()

            with imaplib.IMAP4_SSL(self.config.imap_server, ssl_context=ssl_context) as imap:
                imap.login(self.config.email, password)
                imap.select("inbox")
                
                # Search for recent emails from last 7 days
                date_7_days_ago = (datetime.now() - timedelta(days=7)).strftime("%d-%b-%Y")
                _, messages = imap.search(None, f'(UNSEEN SINCE "{date_7_days_ago}")')
                email_ids = messages[0].split()

                if not email_ids:
                    self.logger.info("No new emails found.")
                    return

                self.logger.info(f"Processing {len(email_ids)} new email(s)...")
                
                feedback_data = []
                sentiment_scores = []

                for email_id in email_ids:
                    try:
                        _, msg_data = imap.fetch(email_id, "(RFC822)")
                        
                        for response_part in msg_data:
                            if not isinstance(response_part, tuple):
                                continue

                            msg = email.message_from_bytes(response_part[1])
                            
                            processed_feedback = self._process_single_email(msg)
                            if processed_feedback:
                                feedback_data.append(processed_feedback['data'])
                                sentiment_scores.append(processed_feedback['sentiment'])

                    except Exception as email_error:
                        self.logger.error(f"Error processing email {email_id}: {email_error}")
                        continue

                if feedback_data:
                    self._save_and_report_feedback(feedback_data, sentiment_scores)
                else:
                    self.logger.info("No valid feedback found in processed emails.")

        except Exception as main_error:
            self.logger.error(f"Email processing failed: {main_error}")

    def _process_single_email(self, msg) -> Optional[Dict[str, Any]]:
        """Process a single email with robust error handling"""
        try:
            subject = self._decode_email_header(msg["Subject"])
            sender = self._decode_email_header(msg.get("From"))
            date = parsedate_to_datetime(msg["Date"])
            formatted_date = date.strftime("%d/%m/%Y")

            email_body = self._extract_email_body(msg)
            if not email_body:
                return None

            feedback = self.extract_feedback(email_body)
            customer_name = self.extract_customer_name(email_body)
            order_id = self.extract_order_id(email_body)

            if (feedback == "No feedback found." or 
                customer_name == "Not available" or 
                order_id == "Not available"):
                return None

            sentiment = self.analyze_sentiment(feedback)

            feedback_entry = FeedbackData(
                email_from=sender,
                subject=subject,
                date=formatted_date,
                customer_name=customer_name,
                order_id=order_id,
                feedback=feedback,
                sentiment=sentiment
            )

            return {'data': feedback_entry, 'sentiment': sentiment}

        except Exception as e:
            self.logger.error(f"Error processing email: {e}")
            return None

    def _save_and_report_feedback(self, feedback_data: List[FeedbackData], sentiment_scores: List[str]):
        """Centralized method for saving data and sending reports"""
        try:
            # Save to Excel
            file_path = os.path.join(
                self.config.output_dir, 
                f"customer_feedback_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
            )
            df = pd.DataFrame([asdict(data) for data in feedback_data])
            df.to_excel(file_path, index=False)
            self.logger.info(f"Feedback saved to {file_path}")

            # Create and send summary
            sentiment_img = self._create_sentiment_chart(pd.Series(sentiment_scores))
            self._send_summary_email(feedback_data, sentiment_img)

        except Exception as e:
            self.logger.error(f"Reporting process failed: {e}")

    def _decode_email_header(self, header: str) -> str:
        """Safely decode email headers"""
        try:
            decoded_parts = decode_header(header or '')
            parts = [
                part[0].decode(part[1] or 'utf-8', errors='ignore') 
                for part in decoded_parts
            ]
            return ' '.join(parts) if parts else 'Unknown'
        except Exception as e:
            self.logger.warning(f"Header decoding error: {e}")
            return 'Unknown'

    # [Rest of the methods remain mostly the same as in the original implementation]
    # ... (extract_feedback, extract_customer_name, extract_order_id, analyze_sentiment, etc.)

def main():
    """Enhanced main entry point with configuration management"""
    try:
        config = EmailConfig(
            email='your_email@gmail.com',
            service_email='service_email@example.com'
        )
        
        # Optional: Securely store password (run this separately first)
        # analyzer = EnhancedFeedbackAnalyzer(config)
        # analyzer._secure_store_password('Gmail', config.email, 'your_app_password')

        analyzer = EnhancedFeedbackAnalyzer(config)
        analyzer.process_emails()

    except Exception as e:
        print(f"Initialization failed: {e}")

if __name__ == "__main__":
    main()
