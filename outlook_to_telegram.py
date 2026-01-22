import os, imaplib, email, requests
from bs4 import BeautifulSoup
from email.header import decode_header

BOT_TOKEN = os.environ["BOT_TOKEN"]
CHAT_ID = os.environ["CHAT_ID"]
OUTLOOK_USER = os.environ["OUTLOOK_USER"]
OUTLOOK_APP_PASSWORD = os.environ["OUTLOOK_APP_PASSWORD"]

# Recomendado: filtra SOLO correos del portal (ej: noreply@idealista.com)
FROM_FILTER = os.environ.get("FROM_FILTER", "").strip()

def tg_send(text: str):
    url = f"https://api.telegram.org/bot{BOT_TOKEN}/sendMessage"
    requests.post(url, json={
        "chat_id": CHAT_ID,
        "text": text,
        "disable_web_page_preview": False
    }, timeout=20).raise_for_status()

def dec(s):
    if not s:
        return ""
    out = ""
    for part, enc in decode_header(s):
        out += part.decode(enc or "utf-8", "replace") if isinstance(part, bytes) else part
    return out

def extract_text_and_links(msg):
    text = ""
    links = []

    def parse_html(h):
        soup = BeautifulSoup(h, "html.parser")
        for a in soup.select("a[href]"):
            href = (a.get("href") or "").strip()
            if href.startswith("http"):
                links.append(href)
        return soup.get_text("\n", strip=True)

    if msg.is_multipart():
        for p in msg.walk():
            if "attachment" in str(p.get("Content-Disposition", "")).lower():
                continue
            ctype = p.get_content_type()
            payload = p.get_payload(decode=True) or b""
            charset = p.get_content_charset() or "utf-8"
            content = payload.decode(charset, errors="replace")

            if ctype == "text/plain" and not text and content.strip():
                text = content.strip()
            if ctype == "text/html":
                # si no tenemos texto todavía, sacamos el texto del html
                if not text:
                    text = parse_html(content)
                else:
                    # aunque ya tengamos texto, extraemos links
                    parse_html(content)
    else:
        payload = msg.get_payload(decode=True) or b""
        charset = msg.get_content_charset() or "utf-8"
        content = payload.decode(charset, errors="replace")
        if msg.get_content_type() == "text/html":
            text = parse_html(content)
        else:
            text = content.strip()

    # limpia links duplicados manteniendo orden
    uniq_links = list(dict.fromkeys(links))

    if len(text) > 700:
        text = text[:700] + "…"

    return text, uniq_links[:5]

def main():
    imap = imaplib.IMAP4_SSL("outlook.office365.com", 993)
    imap.login(OUTLOOK_USER, OUTLOOK_APP_PASSWORD)
    imap.select("INBOX")

    query = f'(UNSEEN FROM "{FROM_FILTER}")' if FROM_FILTER else "(UNSEEN)"
    status, data = imap.search(None, query)
    if status != "OK":
        imap.logout()
        return

    ids = data[0].split()
    for eid in ids[-10:]:  # máximo 10 para evitar spam
        _, msg_data = imap.fetch(eid, "(RFC822)")
        msg = email.message_from_bytes(msg_data[0][1])

        subject = dec(msg.get("Subject"))
        sender = dec(msg.get("From"))
        body, links = extract_text_and_links(msg)

        out = f"🏠 Alerta pisos\n🧾 {subject}\n👤 {sender}\n\n{body}"
        if links:
            out += "\n\n🔗 Enlaces:\n" + "\n".join(links)

        tg_send(out)
        imap.store(eid, "+FLAGS", "\\Seen")

    imap.logout()

if __name__ == "__main__":
    main()
