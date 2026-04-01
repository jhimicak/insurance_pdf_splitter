import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from email.utils import formataddr
import os
import ssl

class EmailSender:
    def __init__(self, smtp_server, smtp_port, sender_email, sender_password):
        self.smtp_server = smtp_server
        self.smtp_port = int(smtp_port)
        self.sender_email = sender_email
        self.sender_password = sender_password
        self.server = None

    def _get_connection(self):
        """Internal helper to create a connection based on the port, with SSL verification disabled for compatibility."""
        # Create a flexible SSL context that ignores certificate validation (common for internal corporate servers)
        context = ssl.create_default_context()
        context.check_hostname = False
        context.verify_mode = ssl.CERT_NONE

        if self.smtp_port == 465:
            # SSL Connection
            server = smtplib.SMTP_SSL(self.smtp_server, self.smtp_port, context=context)
        else:
            # Standard Connection (usually 587 or 25)
            server = smtplib.SMTP(self.smtp_server, self.smtp_port)
            try:
                # Try to upgrade to TLS if supported (ignoring verification errors)
                if server.has_ext("STARTTLS"):
                    server.starttls(context=context)
            except:
                pass # If STARTTLS fails, we continue
        
        server.login(self.sender_email, self.sender_password)
        return server

    def connect(self):
        """Creates a persistent SMTP connection."""
        try:
            self.server = self._get_connection()
            return True, "성공"
        except Exception as e:
            return False, str(e)

    def close(self):
        """Closes the SMTP connection."""
        if self.server:
            try:
                self.server.quit()
            except:
                pass
            self.server = None

    def send_email(self, receiver_email, subject, body, attachment_path=None, use_existing_session=False, sender_name=None, display_email=None, is_html=False):
        """
        Sends an email. Can use an existing session for bulk sending.
        """
        try:
            msg = MIMEMultipart()
            # Combine name and email for a professional look
            f_name = sender_name if sender_name else ""
            f_email = display_email if display_email else self.sender_email
            msg['From'] = formataddr((f_name, f_email))
                
            msg['To'] = receiver_email
            msg['Subject'] = subject
            
            # Use HTML if requested
            msg.attach(MIMEText(body, 'html' if is_html else 'plain'))

            if attachment_path and os.path.exists(attachment_path):
                filename = os.path.basename(attachment_path)
                with open(attachment_path, "rb") as attachment:
                    part = MIMEBase("application", "octet-stream")
                    part.set_payload(attachment.read())
                
                encoders.encode_base64(part)
                part.add_header(
                    "Content-Disposition",
                    "attachment",
                    filename=("utf-8", "", filename),
                )
                msg.attach(part)

            text = msg.as_string()

            if use_existing_session and self.server:
                self.server.sendmail(self.sender_email, receiver_email, text)
            else:
                # One-off session
                temp_server = self._get_connection()
                temp_server.sendmail(self.sender_email, receiver_email, text)
                temp_server.quit()
            
            return True, "성공"
        except Exception as e:
            return False, str(e)

if __name__ == "__main__":
    pass
