import smtplib
import yaml
import csv
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

# Load Configuration from YAML
with open("config.yml", "r", encoding="utf-8") as file:
    config = yaml.safe_load(file)

# Extract details
your_details = config["your_details"]
SMTP_SERVER = config["smtp"]["server"]
SMTP_PORT = config["smtp"]["port"]
SENDER_EMAIL = config["smtp"]["sender_email"]
SENDER_PASSWORD = config["smtp"]["sender_password"]

EMAIL_TEMPLATE_FILE = config["files"]["email_template"]
ATTACHMENT_FILE = config["files"]["resume"]
EMAIL_LIST_FILE = config["files"]["email_list"]

# Read Email Template
with open('templates\\' + EMAIL_TEMPLATE_FILE, "r", encoding="utf-8") as file:
    template = file.read()

# Function to send an email
def get_msg(receiver_name, receiver_email):
    msg = MIMEMultipart()
    msg["From"] = SENDER_EMAIL
    msg["To"] = receiver_email
    msg["Subject"] = f"Inquiry About Open Positions at {receiver_name}"

    # Personalize the template
    email_body = template
    for key, value in your_details.items():
        email_body = email_body.replace(f"[{key}]", value)
        
    email_body = email_body.replace("[Company Name]", receiver_name)

    # Attach text body
    msg.attach(MIMEText(email_body, "html"))

    # Attach the resume
    with open('attachment\\' + ATTACHMENT_FILE, "rb") as attachment:
        part = MIMEBase("application", "octet-stream")
        part.set_payload(attachment.read())
    encoders.encode_base64(part)
    part.add_header("Content-Disposition", f"attachment; filename={ATTACHMENT_FILE}")
    msg.attach(part)
        
    return msg

# Read and send emails in bulk
with open('data\\' + EMAIL_LIST_FILE, "r", encoding="utf-8") as file:
    reader = csv.reader(file)
    try:
        with smtplib.SMTP_SSL(SMTP_SERVER, SMTP_PORT) as server:  # Use SMTP_SSL
            server.login(SENDER_EMAIL, SENDER_PASSWORD)
            for name, email in reader:
                msg = get_msg(name, email)
                try:
                    server.sendmail(SENDER_EMAIL, email, msg.as_string())
                    print(f"✅ Email sent to {email} successfully!")
                except Exception as e:
                    print(f"❌ Failed to send email to {email}: {e}")
                
    except Exception as e:
        print(f"Failed to login: {e}")