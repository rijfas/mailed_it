# MailedIt!

<center>
   <img src="assets/logo.png">
</center>
MailedIt! is a Mass Mailing Application built using Tkinter, allowing you to send personalized emails to recipients using an HTML template and data from an Excel file.

## How to Use

1. Enter your Gmail credentials:

   - Provide your Gmail email address and password(generate app specific password using this [link](https://myaccount.google.com/apppasswords)) in the respective entry fields.
   - Make sure to use a Gmail account or Gmail-compatible account to send emails.

2. Select an HTML Template:

   - Click the "Browse" button next to "Select HTML Template File" to choose an HTML file containing your email template.
   - The template should use Jinja2 syntax for variable placeholders, enclosed within double curly braces {{ }}.
   - Example: `<p>Hello, {{ name }}!</p>` - The {{ name }} will be replaced with the value of the "name" variable during rendering.

3. Select an Excel Data File:

   - Click the "Browse" button next to "Select Excel Data File" to choose an Excel file containing recipient data.
   - The Excel file should have a header row with column names corresponding to the variables used in the HTML template.
   - Each row represents a data entry, and the values in each cell will be used to replace the corresponding variables in the template.
   - Example:
     ```
     | name  | age |
     |-------|-----|
     | John  | 25  |
     | Alice | 30  |
     ```

4. Start Mailing:

   - Click the "Start Mailing" button to begin the mass mailing process.
   - The application will send personalized emails to each recipient in the Excel data file, using the selected HTML template.
   - A progress window will display the status of the mailing process.

5. Help:
   - Click the "Help" button to get useful information about obtaining app-specific passwords from Google account settings and using Jinja2 syntax for the HTML template.

## Important Notes

- This application uses Gmail's SMTP server to send emails. Ensure you have a valid Gmail account or a Gmail-compatible account.
- To avoid authentication issues, enable "Less Secure Apps" in your Google account settings, or create an "App Password" for this application.
- The provided Excel file should be properly formatted, containing the necessary data for the HTML template.
- The application will show error alerts for any issues related to invalid files, insufficient data, or SMTP errors during the mailing process.
- If the provided HTML template or Excel data file is not readable or invalid, the application will display an error alert.

## Dependencies

- Tkinter (Included in Python standard library)
- Pandas (For reading data from the Excel file)
- Jinja2 (For rendering the HTML template)

Make sure to install the required dependencies before running the application.
