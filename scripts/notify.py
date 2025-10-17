import os, smtplib, csv
from email.message import EmailMessage
from dotenv import load_dotenv
load_dotenv()

HOST = os.getenv("SMTP_HOST")
PORT = int(os.getenv("SMTP_PORT","25"))
FROM = os.getenv("MAIL_FROM")

def _email_map():
    m={}
    with open("sql/researchers.csv", encoding="utf-8") as f:
        for r in csv.DictReader(f):
            if r.get("active","1")=="1":
                m[r["name"]] = r["email"]
    return m

def notify(assignments):
    by = {}
    for a in assignments:
        by.setdefault(a["researcher"], []).append(a)

    email_map = _email_map()
    with smtplib.SMTP(HOST, PORT) as s:
        for name, rows in by.items():
            to = email_map.get(name)
            if not to: 
                continue
            items = "".join(f"<li>{r['sample_no']} - {r['item']}</li>" for r in rows)
            msg = EmailMessage()
            msg["Subject"] = "[EIMS 배정] 금일 할당 건"
            msg["From"] = FROM
            msg["To"] = to
            msg.set_content("HTML 전용")
            msg.add_alternative(f"<h3>{len(rows)}건</h3><ul>{items}</ul>", subtype="html")
            s.send_message(msg)
