from PIL import Image, ImageDraw, ImageFont
import pandas as pd
import gdown
import smtplib
import ssl
from email.mime.multipart import MIMEMultipart as MM
from email.mime.text import MIMEText 
from email.mime.image import MIMEImage
import datetime
import time
import os
import keyboard
import random
def add_img(text):
    image_path = 'bday.jpg'
    image = Image.open(image_path)
    font_path = 'Garogier.ttf' 
    font_size = 100
    font_color = (0, 0, 0)  
    font = ImageFont.truetype(font_path, font_size)
    draw = ImageDraw.Draw(image)
    text_bbox = draw.textbbox((0, 0), text, font=font)
    image_width, image_height = image.size
    text_position = ((image_width - text_bbox[2]) // 2, 1970)
    draw.text(text_position, text, font=font, fill=font_color)
    output_image_path = 'bdy1.jpg'
    image.save(output_image_path)
# Downloading file from Google Drive
file_url = 'https://drive.google.com/uc?id=1KokZsAka40dkQdCOMK7RUNB41GXbIfut'
gdown.download(file_url, 'data.xlsx', quiet=False)
# Reading the XLSX file into a pandas DataFrame
df = pd.read_excel('data.xlsx', engine='openpyxl')
# Checking the data type of 'DOB' column

# Converting 'DOB' column to datetime
df['DOB'] = pd.to_datetime(df['DOB'])
# Function to send email
def email_func(subject, birthday_receiver, name, institute_name, department_name):
    receiver = birthday_receiver
    sender = 'ghanshyamg.cs.22@nitj.ac.in'  # Email id (update)
    app_password = 'cqap xrel yghj xqjx'    # App password (update)
    msg = MM()
    msg['Subject'] = subject + ' ' + str(name) + '!'
    HTML = """
    <html>
    <body>
        <p>Dear {},</p>
        <p>On behalf of the alumni cell at {}, it gives me immense pleasure to extend our heartfelt birthday wishes to you on this special day.</p>
        <p>May this year bring you immense joy, success, and fulfillment in all your endeavors.</p>
        <p>If time permits, we would love to hear about your recent accomplishments and experiences since graduation. Feel free to share any updates or anecdotes you may have.</p>
        <p>As you reminisce about the memories forged during your time with us, remember that you will always have a special place in our hearts.</p>
        <p>Warm regards,</p>
        <p>{}<br>
        {}</p>
    </body>
    </html>
    """.format(name, institute_name, department_name, institute_name)
    msg.attach(MIMEText(HTML, 'html'))
    with open('bdy1.jpg', 'rb') as image_file:
        image = MIMEImage(image_file.read())
        image.add_header('Content-ID', '<bdy_image>')
        image_filename = name
        image.add_header('Content-Disposition', f'attachment; filename="{os.path.basename(image_filename)}"')
        msg.attach(image)
    SSL_context = ssl.create_default_context()
    server = smtplib.SMTP_SSL(host='smtp.gmail.com', port=465, context=SSL_context)
    server.login(sender, app_password)
    server.sendmail(sender, receiver, msg.as_string())


# Current date and time
def date_now():
    today = datetime.date.today()
    year = today.year
    return today, year
while 1:
    if keyboard.is_pressed('q'):
        break
    # Checking if it's 10:10
    if datetime.datetime.now().hour == 16 and datetime.datetime.now().minute == 32:
        # time.sleep(60)
        today, year = date_now() 
        for i in range(0, len(df)):
            day = df['DOB'].dt.day[i]
            month = df['DOB'].dt.month[i]
            name = df['Name'][i]
            email = df['Email'][i]
            birthdate = datetime.date(year, month, day)
            if birthdate == today:
                add_img(f'{name}')
                email_func('Happy Birthday', email, name, "Dr. B.R Ambedkar National Institute of Technology, Jalandhar", "Alumni Cell")
                print(f'{name}')
                print('Sent ')
                sleep_duration = random.randint(30, 50)
                time.sleep(sleep_duration)
            else:
                print("did not send") 
            print(i)
    
    
