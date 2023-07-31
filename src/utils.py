import smtplib, ssl
from smtplib import SMTPAuthenticationError, SMTPException
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from jinja2 import Environment, FileSystemLoader, TemplateNotFound
import pandas as pd
import os


def send_email(sender_email, password, receiver_email, subject, message_html):
    """
    Send an email with HTML content using Gmail's SMTP server.

    Args:
        sender_email (str): The email address of the sender.
        password (str): The password or app-specific password of the sender's email account.
        receiver_email (str): The email address of the recipient.
        subject (str): The subject of the email.
        message_html (str): The HTML content of the email message.

    Raises:
        smtplib.SMTPAuthenticationError: If the login credentials are incorrect.
        smtplib.SMTPException: If an error occurs during the sending process.

    Returns:
        None: The function has no return value; it sends the email directly.

    Notes:
        - This function uses Gmail's SMTP server to send the email. Make sure to enable "Less Secure Apps"
          or create an "App Password" in your Google account settings if you encounter authentication issues.
        - The `ssl.create_default_context()` method is used to establish a secure connection.
        - The email is sent as a multipart message with HTML content using the MIMEText class.
          The content type is set to "html" for displaying the HTML content correctly in the email.
        - The sender's email address is set in the "From" field, and the recipient's email address is set in the "To" field.
        - The subject of the email is set using the "Subject" field.
        - The `server.sendmail()` method is used to send the email with the sender's and recipient's email addresses.

    Example:
        send_email("sender@gmail.com", "password", "receiver@example.com", "Important Update",
                   "<html><body><h1>Hello, John!</h1><p>This is an important update.</p></body></html>")
    """
    message = MIMEMultipart("alternative")
    message["Subject"] = subject
    message["From"] = sender_email
    message["To"] = receiver_email

    # Attach the HTML content to the email as a MIMEText object
    message.attach(MIMEText(message_html, "html"))

    # Create a secure SSL/TLS context for the connection
    context = ssl.create_default_context()

    # Connect to Gmail's SMTP server over SSL and send the email
    with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=context) as server:
        # Log in to the sender's email account using the provided credentials
        server.login(sender_email, password)

        # Send the email to the recipient
        server.sendmail(sender_email, receiver_email, message.as_string())


def check_file_existence_and_readability(file_path):
    """
    Check if a file exists and is readable.

    Args:
        file_path (str): The path to the file.

    Returns:
        bool: True if the file exists and is readable, False otherwise.
    """
    if os.path.exists(file_path):
        # Check if the file is readable
        if os.access(file_path, os.R_OK):
            return True
        else:
            print(f"File '{file_path}' exists but is not readable.")
            return False
    else:
        print(f"File '{file_path}' does not exist.")
        return False


def render_html_template(template_path, context):
    """
    Renders an HTML template using Jinja2.

    Args:
        template_path (str): The path to the HTML template file.
        context (dict): A dictionary containing the variables and their values to be used in the template.

    Returns:
        str: The rendered HTML content as a string, or None if the template is not found.
    """
    try:
        template_dir = os.path.dirname(template_path)
        env = Environment(loader=FileSystemLoader(template_dir))

        template_filename = os.path.basename(template_path)
        template = env.get_template(template_filename)

        rendered_html = template.render(context)

        return rendered_html
    except TemplateNotFound:
        print(f"Template not found: {template_path}")
        return None


def read_excel_to_dict(file_path):
    """
    Read an Excel file and convert it into a dictionary.

    Args:
        file_path (str): The path to the Excel file.

    Returns:
        dict: A dictionary where keys are column names, and values are lists of data for each column.
    """
    try:
        # Read the Excel file into a DataFrame using pandas
        df = pd.read_excel(file_path)

        # Convert the DataFrame to a dictionary
        data_dict = df.to_dict(orient="records")

        return data_dict
    except Exception as e:
        print(f"Error reading Excel file: {e}")
        return None
