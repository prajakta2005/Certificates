import gspread
from oauth2client.service_account import ServiceAccountCredentials
from PIL import Image, ImageDraw, ImageFont
import os
import re  # For sanitizing file names
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from dotenv import load_dotenv
from reportlab.pdfgen import canvas

from dotenv import load_dotenv

load_dotenv()
#require('dotenv').config();
#console.log(process.env);

# Load environment variables from .env file
load_dotenv()

# Access environment variables
credentials_file = os.getenv("GOOGLE_CREDENTIALS_FILE")
sheet_name = os.getenv("SHEET_NAME")
template_path = os.getenv("TEMPLATE_PATH")
output_path = os.getenv("OUTPUT_PATH")
font_path = os.getenv("FONT_PATH")
smtp_server = os.getenv("SMTP_SERVER")
port = int(os.getenv("SMTP_PORT"))
sender_email = os.getenv("SENDER_EMAIL")
sender_password = os.getenv("SENDER_PASSWORD")
subject = os.getenv("SUBJECT")
body_template = os.getenv("BODY_TEMPLATE")


def authenticate_google_sheets(credentials_file, sheet_name):
    """
    Authenticate with Google Sheets API and return the specified sheet.
    """
    try:
        scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
        creds = ServiceAccountCredentials.from_json_keyfile_name(credentials_file, scope)
        client = gspread.authorize(creds)
        sheet = client.open(sheet_name).sheet1
        return sheet
    except Exception as e:
        print(f"Error authenticating Google Sheets: {e}")
        exit()

def sanitize_file_name(name):
    """
    Sanitize the file name by removing invalid characters.
    """
    return re.sub(r'[<>:"/\\|?*]', '_', name)  # Replace invalid characters with an underscore

def create_certificate(template_path, output_path, name, font_path, font_size, position):
    """
    Generate a certificate by placing the name on the provided template.
    """
    try:
        # Load the template image
        template = Image.open(template_path)
        draw = ImageDraw.Draw(template)
        
        # Load the font with bold weight
        font = ImageFont.truetype(font_path, font_size)
        
        # Draw the name on the certificate
        draw.text(position, name, font=font, fill="black")  # Adjust color if necessary
        
        # Ensure the output directory exists
        os.makedirs(output_path, exist_ok=True)
        
        # Sanitize file name to avoid invalid characters
        sanitized_name = sanitize_file_name(name)
        output_file = os.path.join(output_path, f"{sanitized_name}_certificate.png")
        
        # Save the certificate as a new file
        template.save(output_file)
        print(f"Certificate saved for {name} at {output_file}")
        return output_file
    except Exception as e:
        print(f"Error creating certificate for {name}: {e}")
        return None


def send_email(smtp_server, port, sender_email, sender_password, recipient_email, subject, body, attachment_path):
    """
    Send an email with an attachment.
    """
    try:
        # Set up the email
        msg = MIMEMultipart()
        msg['From'] = sender_email
        msg['To'] = recipient_email
        msg['Subject'] = subject

        # Attach the email body
        msg.attach(MIMEText(body, 'plain'))

        # Attach the certificate
        if attachment_path:
            with open(attachment_path, 'rb') as attachment:
                part = MIMEBase('application', 'octet-stream')
                part.set_payload(attachment.read())
                encoders.encode_base64(part)
                part.add_header('Content-Disposition', f'attachment; filename={os.path.basename(attachment_path)}')
                msg.attach(part)

        # Send the email
        with smtplib.SMTP(smtp_server, port) as server:
            server.starttls()
            server.login(sender_email, sender_password)
            server.send_message(msg)
            print(f"Email sent to {recipient_email}")
    except Exception as e:
        print(f"Error sending email to {recipient_email}: {e}")


def convert_to_pdf(image_path, output_path):
    """
    Convert a PNG image to a PDF.
    """
    try:
        pdf_path = os.path.splitext(image_path)[0] + ".pdf"  # Change file extension to .pdf
        image = Image.open(image_path)
        pdf_canvas = canvas.Canvas(pdf_path)

        # Adjust PDF size to match the image dimensions
        pdf_canvas.setPageSize(image.size)

        # Draw the image onto the PDF canvas
        pdf_canvas.drawImage(image_path, 0, 0, width=image.size[0], height=image.size[1])

        # Finalize the PDF
        pdf_canvas.save()
        print(f"Converted {image_path} to {pdf_path}")
        return pdf_path
    except Exception as e:
        print(f"Error converting {image_path} to PDF: {e}")
        return None

def main():
    try:
        credentials_file = r"C:/Users/admin/OneDrive/My work (Prajakta)/Python/tempcertificates/certificates-448816-0745b9fd877c.json"
        sheet_name = "demo(Responses)"
        sheet = authenticate_google_sheets(credentials_file, sheet_name)
        
        names = sheet.col_values(2)[1:]  # Assumes names are in the second column, excluding header
        emails = sheet.col_values(3)[1:]  # Assumes emails are in the third column, excluding header

        if not names or not emails:
            print("No names or emails found in the Google Sheet.")
            return

        template_path = "C:/Users/admin/OneDrive/My work (Prajakta)/Python/tempcertificates/Certificate_20250122_201746_0000.png"
        output_path = "C:/Users/admin/OneDrive/My work (Prajakta)/Python/tempcertificates/certificates"
        font_path = r"C:/Users/admin/OneDrive/My work (Prajakta)/Python/tempcertificates/Noto_Sans/NotoSans-VariableFont_wdth,wght.ttf"
        font_size = 125
        position = (1300, 1200)

        smtp_server = "smtp.gmail.com"
        port = 587
        sender_email = "prajakta23joshi@gmail.com"
        sender_password = "qrxq duwh awdk jxxi"
        subject = "Certificate of Participation – GMRT Technical Event 2025"
        body_template = "Dear {name},\n\nThank you for your valuable participation in the GMRT Technical Event 2025, held on 17 January 2025. It was a pleasure having you as part of this event, and your enthusiasm greatly contributed to its success. \nPlease find attached your Certificate of Participation as a token of appreciation for your involvement.\n\nWe look forward to your participation in future events!\n\nBest regards,\nTeam Antariksh ✨"

        for name, email in zip(names, emails):
            if name.strip() and email.strip():
                # Generate the certificate as a PNG
                certificate_path = create_certificate(template_path, output_path, name.strip(), font_path, font_size, position)
                if certificate_path:
                    # Convert the PNG to a PDF
                    pdf_certificate_path = convert_to_pdf(certificate_path, output_path)
                    if pdf_certificate_path:
                        # Send the PDF certificate via email
                        body = body_template.format(name=name.strip())
                        send_email(smtp_server, port, sender_email, sender_password, email.strip(), subject, body, pdf_certificate_path)
    except Exception as e:
        print(f"Error in main function: {e}")

if __name__ == "__main__":
    main()
