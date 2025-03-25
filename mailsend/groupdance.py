import smtplib
import pandas as pd
import qrcode
from io import BytesIO
from email.message import EmailMessage
from email.mime.image import MIMEImage
import time

# Load participant details from Excel
file_path = "mailsend\mesent.xlsx"
df = pd.read_excel(file_path, sheet_name='Harmonicks(Beat Battle (Group D').fillna("N/A")

# Email credentials (Use an App Password if using Gmail)
EMAIL_ADDRESS = "kavi7010764469@gmail.com"
EMAIL_PASSWORD = "vunk efsl jkry yxne"
SMTP_SERVER = "smtp.gmail.com"
SMTP_PORT = 465

# Function to generate QR code
def generate_qr(data):
    qr = qrcode.QRCode(
        version=1,
        error_correction=qrcode.constants.ERROR_CORRECT_L,
        box_size=10,
        border=4,
    )
    qr.add_data(data)
    qr.make(fit=True)
    img = qr.make_image(fill="black", back_color="white")
    buffered = BytesIO()
    img.save(buffered, format="PNG")
    return buffered.getvalue()

# Function to send email
def send_email(to_email, details, qr_data):
    msg = EmailMessage()
    msg["Subject"] = f"🎉 Dakshaa Confirmation | {details['Preferred Hormonicks Events']} Participation Approved 🎟️"
    msg["From"] = EMAIL_ADDRESS
    msg["To"] = to_email
    
    # HTML Email Content with embedded QR code
    html_content = f"""
<html>
<head>
<style>
    body {{
        font-family: Arial, sans-serif;
        background-color: #f4f4f9;
        margin: 0;
        padding: 0;
        color: #333;
    }}
    .container {{
        max-width: 600px;
        margin: 20px auto;
        padding: 20px;
        background-color: #ffffff;
        border-radius: 12px;
        box-shadow: 0 4px 8px rgba(0,0,0,0.1);
        border: 1px solid #e1e1e1;
        text-align: center;
    }}
    .college-name {{
        font-size: 22px;
        font-weight: bold;
        color: #004085;
        background: #e3f2fd;
        padding: 10px;
        border-radius: 8px;
        display: inline-block;
        margin-bottom: 10px;
    }}
    h2 {{
        color: #007BFF;
        font-size: 24px;
        margin-bottom: 10px;
    }}
    p {{
        font-size: 16px;
        line-height: 1.6;
        color: #555;
        margin-bottom: 15px;
    }}
    ul {{
        padding-left: 0;
        list-style: none;
        margin-bottom: 20px;
        text-align: left;
        display: inline-block;
    }}
    li {{
        margin-bottom: 8px;
        font-size: 16px;
        color: #333;
    }}
    .highlight {{
        color: #007BFF;
        font-weight: bold;
    }}
    .qr-code {{
        text-align: center;
        margin: 20px 0;
    }}
    img {{
        border: 1px solid #ddd;
        border-radius: 8px;
        padding: 5px;
        box-shadow: 0 4px 8px rgba(0,0,0,0.1);
    }}
    .footer {{
        margin-top: 30px;
        font-size: 16px;
        color: #777;
        text-align: center;
        padding-top: 15px;
        border-top: 1px solid #eee;
    }}
    .logo {{
        display: block;
        margin: 0 auto 10px auto;
        width: 220px;
        padding: 10px;
        border-radius: 12px;
    }} 
    .footer-banner {{
        background-color: #002147; /* Dark blue background */
        color: #ffffff; /* White text */
        text-align: center;
        padding: 12px;
        font-size: 16px;
        font-weight: bold;
        border-radius: 8px;
        margin-top: 20px;
    }}
    .special-note{{
        background-color: #ffeb3b; /* Yellow highlight */
        padding: 15px;
        border-radius: 8px;
        font-weight: bold;
        color: #333;
        text-align: center;
        margin-top: 20px;
        box-shadow: 0 4px 8px rgba(0,0,0,0.1);
    }}
    .condition {{
        background-color: #f8d7da; /* Light red highlight */
        color: #721c24;
        padding: 10px;
        border-radius: 8px;
        font-weight: bold;
        text-align: center;
        margin-top: 10px;
        border: 1px solid #f5c6cb;
    }}
</style>
</head>
<body>
    <div class="container">
        <p class="college-name">K S Rangasamy College of Technology</p>
        <img src="cid:event_logo" class="logo" alt="Event Logo"/>
        <h2>Hello {details['Team Lead Name']}, 👋</h2>
        <p>Thank you for registering your team <span class="highlight">{details['Team Name']}</span> for the <span class="highlight">{details['Preferred Hormonicks Events']}</span> event!</p>
        <p><strong>Team Details:</strong></p>
        <ul>
            <li><strong>Team Name:</strong> {details['Team Name']}</li>
            <li><strong>Event Type:</strong> Harmonicks(Beat Battle (Group Dance))</li>
            <li><strong>Preferred Hormonicks Events:</strong> {details['Preferred Hormonicks Events']}</li>

            <li><strong>Team Lead:</strong> {details['Team Lead Name']} ({details['Team Lead Mail ID']})</li>
            <li><strong>Mobile Number:</strong> {details['Mobile Number']}</li>
            <li><strong>College Name:</strong> {details['College Name']}</li>
            <li><strong>Stream of Study:</strong> {details['Stream of Study']}</li>
            <li><strong>Department:</strong> {details['Department']}</li>
            <li><strong>Year:</strong> {details['Year']}</li>
            <li><strong>Team Member 2:</strong> {details.get('Team Member 2 Name', 'N/A')} ({details.get('Team Member 2 Mail ID', 'N/A')})</li>
            <li><strong>Team Member 3:</strong> {details.get('Team Member 3 Name', 'N/A')} ({details.get('Team Member 3 Mail ID', 'N/A')})</li>
            <li><strong>Team Member 4:</strong> {details.get('Team Member 4 Name', 'N/A')} ({details.get('Team Member 4 Mail ID', 'N/A')})</li>
            <li><strong>Team Member 5:</strong> {details.get('Team Member 5 Name', 'N/A')} ({details.get('Team Member 5 Mail ID', 'N/A')})</li>
            <li><strong>Team Member 6:</strong> {details.get('Team Member 6 Name', 'N/A')} ({details.get('Team Member 6 Mail ID', 'N/A')})</li>
            <li><strong>Team Member 7:</strong> {details.get('Team Member 7 Name', 'N/A')} ({details.get('Team Member 7 Mail ID', 'N/A')})</li>
            <li><strong>Team Member 8:</strong> {details.get('Team Member 8 Name', 'N/A')} ({details.get('Team Member 8 Mail ID', 'N/A')})</li>


        </ul>
        <p>📌 <strong>Please scan the QR code below at the entrance to mark your attendance:</strong></p>
        <div class="qr-code">
            <img src="cid:qr_code" width="200px" alt="QR Code"/>
        </div>
        <p>We look forward to seeing you at the event! 🎉</p>
        <p>For any queries, feel free to reach out to us at:</p>
        <p><strong>Contact: 9489243775</strong></p>
        <div class="footer">
            Best Regards,<br>
            <br>
            <strong>Event Team</strong>
        </div>
        <div class="footer-banner">
            <p>We're hyped to see you at DaKshaa T25! 🚀✨ Get ready to make memories that last a lifetime!</p>
        </div>
    </div>
</body>
</html>
"""

    msg.add_alternative(html_content, subtype='html')

    # Add Event Logo
    with open("mailsend/event_logo.png", "rb") as img:
        logo_image = MIMEImage(img.read())
        logo_image.add_header('Content-ID', '<event_logo>')
        logo_image.add_header('Content-Disposition', 'inline', filename="event_logo.png")
        msg.get_payload().append(logo_image)

    # Add QR Code
    qr_image = MIMEImage(qr_data)
    qr_image.add_header('Content-ID', '<qr_code>')
    qr_image.add_header('Content-Disposition', 'inline', filename="qr_code.png")
    msg.get_payload().append(qr_image)

    # Send email
    try:
        with smtplib.SMTP_SSL(SMTP_SERVER, SMTP_PORT) as server:
            server.login(EMAIL_ADDRESS, EMAIL_PASSWORD)
            server.send_message(msg)
            print(f"✅ Email sent to {to_email}")
            return True
    except Exception as e:
        print(f"❌ Failed to send email to {to_email}: {e}")
        return False

# Counter to track sent emails
emails_sent = 0

# Loop through each row and send email to team leads
for index, row in df.iterrows():
    email = row['Team Lead Mail ID']
    if pd.notna(email):
        qr_data = f"""
        Team Name: {row['Team Name']}
        Event Type: Harmonicks(Beat Battle (Group Dance))
        Preferred Hormonicks Events: {row['Hormonicks Events']}
        Team Lead: {row['Team Lead Name']} ({row['Team Lead Mail ID']})
        Mobile: {row['Team Lead Mobile Number']}
        College Name: {row['College Name']}
        Department: {row['Department']}
        Team Member 2: {row.get('Team Member 2 Name', 'N/A')} ({row.get('Team Member 2 Mail ID', 'N/A')})
        Team Member 3: {row.get('Team Member 3 Name', 'N/A')} ({row.get('Team Member 3 Mail ID', 'N/A')})
        Team Member 4: {row.get('Team Member 4 Name', 'N/A')} ({row.get('Team Member 4 Mail ID', 'N/A')})
        Team Member 5: {row.get('Team Member 5 Name', 'N/A')} ({row.get('Team Member 5 Mail ID', 'N/A')})
        Team Member 6: {row.get('Team Member 6 Name', 'N/A')} ({row.get('Team Member 6 Mail ID', 'N/A')})
        Team Member 7: {row.get('Team Member 7 Name', 'N/A')} ({row.get('Team Member 7 Mail ID', 'N/A')})
        Team Member 8: {row.get('Team Member 8 Name', 'N/A')} ({row.get('Team Member 8 Mail ID', 'N/A')})
       

        """
        qr_data_bytes = generate_qr(qr_data)

        details = {
    'Team Name': row['Team Name'],
    'Team Lead Name': row['Team Lead Name'],
    'Team Lead Mail ID': row['Team Lead Mail ID'],
    'Mobile Number': row['Team Lead Mobile Number'],
    'College Name': row['College Name'],
    'Stream of Study': row['Stream of Study'],
    'Department': row['Department'],
    'Year': row['Year'],
    'Event Type': 'Harmonicks(Beat Battle (Group Dance))',
    'Preferred Hormonicks Events': row['Hormonicks Events'],
    'Team Member 2 Name': row.get('Team Member 2 Name', 'N/A'),
    'Team Member 2 Mail ID': row.get('Team Member 2 Mail ID', 'N/A'),
    'Team Member 2 Mobile Number': row.get('Team Member 2 Mobile Number', 'N/A'),
    'Team Member 3 Name': row.get('Team Member 3 Name', 'N/A'),
    'Team Member 3 Mail ID': row.get('Team Member 3 Mail ID', 'N/A'),
    'Team Member 3 Mobile Number': row.get('Team Member 3 Mobile Number', 'N/A'),
    'Team Member 4 Name': row.get('Team Member 4 Name', 'N/A'),
    'Team Member 4 Mail ID': row.get('Team Member 4 Mail ID', 'N/A'),
    'Team Member 4 Mobile Number': row.get('Team Member 4 Mobile Number', 'N/A'),
    'Team Member 5 Name': row.get('Team Member 5 Name', 'N/A'),
    'Team Member 5 Mail ID': row.get('Team Member 5 Mail ID', 'N/A'),
    'Team Member 5 Mobile Number': row.get('Team Member 5 Mobile Number', 'N/A'),
    'Team Member 6 Name': row.get('Team Member 6 Name', 'N/A'),
    'Team Member 6 Mail ID': row.get('Team Member 6 Mail ID', 'N/A'),
    'Team Member 6 Mobile Number': row.get('Team Member 6 Mobile Number', 'N/A'),
    'Team Member 7 Name': row.get('Team Member 7 Name', 'N/A'),
    'Team Member 7 Mail ID': row.get('Team Member 7 Mail ID', 'N/A'),
    'Team Member 7 Mobile Number': row.get('Team Member 7 Mobile Number', 'N/A'),
    'Team Member 8 Name': row.get('Team Member 8 Name', 'N/A'),
    'Team Member 8 Mail ID': row.get('Team Member 8 Mail ID', 'N/A'),

}


        if send_email(email, details, qr_data_bytes):
            emails_sent += 1
        
        time.sleep(4)

print(f"\n🚀 All emails have been sent successfully! Total emails sent: {emails_sent}")
