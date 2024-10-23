import os
import sys
import smtplib
import csv
import getpass
import time
from rich.console import Console
from rich.progress import track
from rich.theme import Theme
from rich.panel import Panel
from rich.align import Align
from rich.prompt import Prompt
from rich.prompt import Confirm
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders

custom_theme = Theme({"important": "deep_sky_blue1 bold", "texts": "royal_blue1", "input": "gold3"})
console = Console(theme=custom_theme)

count = 0
console.print(Panel(Align.center("[texts]Welcome to the '[important]mass mailing[/important]' script by flush\n A file [input]README.txt[/input] is provided, and it contains ESSENTIAL information.\n All Right Reserved https://flushblog.fr")))

read_readme = Confirm.ask("[gold3]Would you like to read the README?[/gold3]", default="y")

if read_readme:
    console.print("\n[gold3]It is recommended to create a dedicated folder for the script to ensure its proper functioning.[/gold3]\n\n"
                  "[texts]Welcome to the '[important]mass mailing[/important]' script by flush.\n\n"
                  "To begin, please provide an HTML file named '[purple3]content.html[/purple3]' that will contain the content of the email to be sent. You can add the [gold3]{name}[/gold3] attribute in your HTML, which will automatically replace the target's name for each email.\n\n"
                  "Example:\n\n"
                  "[purple3]<p>Hello {name}, I’m sending you this email using 'mass mailing', an incredible mass email sending script!</p>[/purple3]\n\n"
                  "Next, please provide a file named '[purple3]mailing_list.csv[/purple3]' that will let the script know which addresses to target. The file should start with the headers '[gold3]mail[/gold3]' and '[gold3]name[/gold3]', otherwise, the script won’t work.\n\n"
                  "Example:\n\n"
                  "[gold3]mail,name\nmail@mail.com,mail\njackob@mail.com,jackob\nflush@outl.com,flush[/gold3]\n\n"
                  "If you want to send attachments, simply create a folder named '[purple3]join[/purple3]' and drop the files you want to send as attachments along with your emails.\n\n"
                  "After that, please create an app password for the email address that will send the emails. You can follow Google's instructions to create an app password: [link=https://support.google.com/accounts/answer/185833?hl=en]https://support.google.com/accounts/answer/185833?hl=en[/link]\n\n"
                  "Follow the instructions that will appear to configure your email.\n\n"
                  "Thank you for using 'mass mailing'!\n\n"
                  "IF YOUR LOAD STOPS BEFORE FINISHING, IT MEANS YOU INTERRUPTED IT BY CLICKING SOMEWHERE IN THE WINDOW. JUST PRESS ENTER AND THE LOADING WILL RESUME![/texts]\n\n"
                  "All rights reserved to flush, [link=https://flushblog.fr]https://flushblog.fr[/link]\n\n")

while True:
    mail_service = Prompt.ask("[gold3]Please choose the mail service to use (gmail or outlook):[/gold3]", choices=["gmail", "outlook"], default="gmail")
    from_addr = Prompt.ask("[gold3]Please enter the Google email address from which to send the emails: ")
    console.print("[gold3]Please enter the app password associated with this email address[grey93] \n(for security reasons, the entered text is not displayed):")
    password = getpass.getpass()
    from_name = Prompt.ask("[gold3]Enter the sender's name: ")
    subject = Prompt.ask("[gold3]Enter the subject of the email [grey93](writing '{name}' allows you to automatically add the name for each recipient)[/grey93]\n ")

    try:
        with open('mailing_list.csv', 'r') as file:
            reader = csv.reader(file)
            email_list = list(reader)  # Convert the reader to a list
            total_emails = len(email_list) - 1  # Subtract 1 for the header row
    except Exception as e:
        console.print(f"\n\n[red bold]Error opening the mailing_list.csv file.\n\n [blue_violet]{str(e)}")
        input("\n\nPress Enter to close the script.")
        sys.exit()

    attached_files = ', '.join(os.listdir('join')) if os.path.exists('join') else 'None'
    # Display a confirmation message
    console.print(f"\n[deep_sky_blue1]Send information\n"
          f"[gold3]Sender's email address:[grey93] {from_addr}\n"
          f"[gold3]Sender's name:[grey93] {from_name}\n"
          f"[gold3]Email subject:[grey93] {subject}\n"
          f"[gold3]Attached files:[royal_blue1] {attached_files}\n"
          f"[gold3]Email body:[royal_blue1] content.html\n"
          f"[gold3]Number of emails to send:[royal_blue1] {total_emails}\n")

    # Ask for confirmation before sending
    if Confirm.ask("[gold3]Do you confirm sending these emails?[/gold3]", default="n"):
        break

def send_email(subject, from_addr, to_addr, password, name, count):
    try:
        with open('content.html', 'r') as f:
            html_content = f.read()

        html_content = html_content.replace('{name}', name)
        subject = subject.replace('{name}', name)

        msg = MIMEMultipart()
        msg.attach(MIMEText(html_content, "html", "utf-8"))
        msg['Subject'] = subject
        msg['From'] = f"{from_name} <{from_addr}>"
        msg['To'] = to_addr

        if os.path.exists('join') and os.listdir('join'):
            for file in os.listdir('join'):
                with open(f'join/{file}', 'rb') as f:
                    part = MIMEBase('application', 'octet-stream')
                    part.set_payload(f.read())
                    encoders.encode_base64(part)
                    part.add_header('Content-Disposition', f"attachment; filename= {file}")
                    msg.attach(part)
        smtp_server = "smtp.gmail.com" if mail_service == "gmail" else "smtp-mail.outlook.com"
        with smtplib.SMTP_SSL(smtp_server, 465) as server:
            server.login(from_addr, password)
            server.sendmail(from_addr, to_addr, msg.as_string())
        print(f"\033[92m Successfully sent the email to {name}!")
        return count + 1
    except Exception as e:
        console.print(f"\n\n[red bold]Error sending emails from the address [gold3]{from_addr}[/gold3].\n\n [blue_violet]{str(e)}")
        input("\n\nPress Enter to close the script.")
        sys.exit()

for row in track(email_list[1:], total=total_emails, description="Sending in progress.."):  # Skip the header row
    to_addr = row[0]  # Assuming the email address is in the first column
    name = row[1]  # Assuming the name is in the second column
    count = send_email(subject, from_addr, to_addr, password, name, count)

console.print(f"[green4]All emails have been successfully sent, thank you for using '[deep_sky_blue1 bold]mass mailing[/deep_sky_blue1 bold]'.\n sent by the email: [gold3]{from_addr}[/gold3]\n sender's name: [gold3]{from_name}[/gold3]\n Email subject: [gold3]{subject}[/gold3]\n Email content: [purple3]content.html[/purple3]\n [gold3]{count}[/gold3] emails were sent!")
input("\n\nPress Enter to close the script.")
sys.exit()
