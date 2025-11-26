
r"""

C:\analytics\projects\git\lexi\demos\venv\Scripts\python.exe

import runpy ; temp = runpy._run_module_as_main("junk")

-m pip install extract-msg


"""


import os
import sys
import re

file__fullPath = os.path.abspath(__file__)
file__baseName = os.path.basename(file__fullPath)
file__parentDr = os.path.dirname(file__fullPath)
file__fileSysD = (lambda a:lambda v:a(a, v, v))(lambda s, v, x:x if os.path.isdir(x) else (_ for _ in ()).throw(Exception(f"Argument not a directory:'{v}'")) if x==os.path.dirname(x) else s(s, v, os.path.dirname(x)))(file__parentDr)


from pathlib import Path



def fix_to_field(raw_to: str) -> str:
    """
    Extracts {email, friendly} from strings like:
        'user@example.com <None>'
        'User Name <user@example.com>'
        'user@example.com'
        'User <None>'
    and returns a clean RFC822 To header.
    """
    import email.utils 

    raw_to = raw_to.strip()

    # CASE 1 — Has angle brackets of the form:   something <something>
    m = re.match(r"^(.*?)\s*<\s*(.*?)\s*>$", raw_to)
    if m:
        name, addr = m.group(1).strip(), m.group(2).strip()

        # If the address is "None" but the name looks like an email
        if addr.lower() == "none":
            # We assume 'name' was the real email
            email = name
            return email

        # If the name is blank, or literally "None", fix it
        if name == "" or name.lower() == "none":
            return addr

        # Else: both name + email look valid
        return email.utils.formataddr((name, addr))

    # CASE 2 — No angle brackets, but looks like "None"
    if raw_to.lower() == "none":
        return ""

    # CASE 3 — Just an email
    return raw_to
    
if False and __name__ == "__main__":
    ss = [
    "davidnoz@yahoo.com <None>",
    "davidnoz@yahoo.com <Dave>",
    "Dave <david@example.com>",
    "None <david@example.com>",
    "user@example.com",
    ]
    m = max(map(len, ss))
    for s in ss:
        print("%-*s %s" % (m, s, fix_to_field(s)))
    raise Exception("OK")
    


def build_email_from_msg(msg_path: str):
    import extract_msg
    import email.message
    import email.utils
    
    m = extract_msg.Message(msg_path)

    em = email.message.EmailMessage()

    # Headers
    em["From"] = YAHOO_USER
    if m.to:
        em["To"] = ",".join(
            [fix_to_field(addr.strip()) for addr in m.to.split(";") if addr.strip()]
        )
    if m.cc:
        em["Cc"] = ",".join(
            [fix_to_field(addr.strip()) for addr in m.cc.split(";") if addr.strip()]
        )
    em["Subject"] = m.subject or ""

    # Bodies
    text_body = m.body or "See HTML version"
    html_body = m.htmlBody

    # m.body might also be bytes on some messages
    if isinstance(text_body, bytes):
        text_body = text_body.decode("utf-8", "replace")

    em.set_content(text_body)

    if html_body:
        if isinstance(html_body, bytes):
            html_body = html_body.decode("utf-8", "replace")

        # Now html_body is a str, so this works fine
        em.add_alternative(html_body, subtype="html")

    # Attachments
    for att in m.attachments:
        filename = att.longFilename or att.shortFilename or "attachment.bin"
        data = att.data  # bytes

        em.add_attachment(
            data,
            maintype="application",
            subtype="octet-stream",
            filename=filename,
        )

    return em

def send_via_yahoo(msg):
    import smtplib
    with smtplib.SMTP_SSL("smtp.mail.yahoo.com", 465) as smtp:
        smtp.login(YAHOO_USER, YAHOO_APP_PASSWORD)
        smtp.send_message(msg)


if __name__ == "__main__":
    YAHOO_USER = os.environ["YAHOO_USER"] if "YAHOO_USER" in os.environ else "davidnoz@yahoo.com"
    MSG_PATH = os.environ["MSG_PATH"] if "MSG_PATH" in os.environ else os.path.join(file__fileSysD, "asxasx.msg") # same path you used in VBA    
    if "YAHOO_APP_PASSWORD" in os.environ:
        YAHOO_APP_PASSWORD = os.environ["YAHOO_APP_PASSWORD"]
    else:
        1/0
        with open(os.path.join(file__fileSysD, "YAHOO_APP_PASSWORD.txt"), "r") as ff:
            YAHOO_APP_PASSWORD = ff.read().strip()    
    email_msg = build_email_from_msg(MSG_PATH)
    send_via_yahoo(email_msg)
