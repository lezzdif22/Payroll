# mailer.py
import os, mimetypes, platform, ssl, smtplib, base64
from email.message import EmailMessage
import json
import requests
try:
    import msal
except Exception:
    msal = None

# Load .env
try:
    from dotenv import load_dotenv, find_dotenv
    load_dotenv(find_dotenv())
except Exception:
    # tiny fallback loader
    env = os.path.join(os.getcwd(), ".env")
    if os.path.exists(env):
        with open(env, encoding="utf-8") as f:
            for line in f:
                line=line.strip()
                if not line or line.startswith("#") or "=" not in line: continue
                k,v = line.split("=",1); os.environ.setdefault(k.strip(), v.strip())

def _bool_env(name, default=False):
    v = (os.getenv(name) or "").strip().lower()
    if v in ("1","true","yes","y","on"): return True
    if v in ("0","false","no","n","off"): return False
    return default

class Mailer:
    """
    Unified mailer:
      - MAIL_BACKEND=smtp   -> use SMTP (works on Linux & Windows)
      - MAIL_BACKEND=outlook-> use Outlook desktop (Windows only)
    """
    def __init__(self):
        self.backend = (os.getenv("MAIL_BACKEND") or "smtp").strip().lower()

        # Common sender info
        self.from_name  = os.getenv("SMTP_FROM_NAME", "").strip()
        self.from_email = os.getenv("SMTP_FROM_EMAIL", "").strip()
        self.bcc = [e.strip() for e in (os.getenv("SMTP_BCC","").split(",")) if e.strip()]

        # SMTP config (used if backend == smtp)
        self.host   = os.getenv("SMTP_HOST", "").strip()
        self.port   = int(os.getenv("SMTP_PORT", "0") or 0)
        self.user   = os.getenv("SMTP_USER", "").strip()
        self.pw     = os.getenv("SMTP_PASS", "").strip()
        self.use_ssl = _bool_env("SMTP_USE_SSL", True)

        # autodetect host for smtp if empty
        if self.backend == "smtp" and not self.host:
            dom = (self.from_email.split("@")[-1] if self.from_email else "").lower()
            if dom in ("gmail.com","googlemail.com"):
                self.host, self.port, self.use_ssl = "smtp.gmail.com", 465, True
            else:
                self.host, self.port, self.use_ssl = "smtp.office365.com", 587, False

        # Outlook config (used if backend == outlook)
        self.outlook_sender = (os.getenv("OUTLOOK_SENDER") or self.from_email).strip()
        # Graph API config (used if backend == graph)
        self.graph_tenant = (os.getenv("GRAPH_TENANT_ID") or os.getenv("AZURE_TENANT_ID") or "").strip()
        self.graph_client_id = (os.getenv("GRAPH_CLIENT_ID") or "").strip()
        self.graph_client_secret = (os.getenv("GRAPH_CLIENT_SECRET") or "").strip()
        # scopes: for client credentials use /.default, for delegated flows use Mail.Send
        self.graph_scope = os.getenv("GRAPH_SCOPE") or "https://graph.microsoft.com/.default"

        # sane defaults
        if not self.from_email:
            self.from_email = self.user

    # ---------- Public API ----------
    def send(self, to_email, subject, body_text, attachments=None, body_html=None, cc=None):
        if self.backend == "outlook":
            if platform.system() != "Windows":
                raise RuntimeError("MAIL_BACKEND=outlook requires Windows + Outlook desktop.")
            return self._send_outlook(to_email, subject, body_text, attachments or [], body_html, cc or [])
        elif self.backend == "graph":
            return self._send_graph(to_email, subject, body_text, attachments or [], body_html, cc or [])
        else:
            return self._send_smtp(to_email, subject, body_text, attachments or [], body_html, cc or [])

    # ---------- Microsoft Graph backend ----------
    def _acquire_graph_token(self):
        if not msal:
            raise RuntimeError("msal library not installed; add 'msal' to requirements.txt")
        # Prefer client credentials flow when client secret present
        if self.graph_client_id and self.graph_client_secret and self.graph_tenant:
            app = msal.ConfidentialClientApplication(
                client_id=self.graph_client_id,
                client_credential=self.graph_client_secret,
                authority=f'https://login.microsoftonline.com/{self.graph_tenant}'
            )
            scopes = [self.graph_scope] if self.graph_scope else ["https://graph.microsoft.com/.default"]
            token = app.acquire_token_silent(scopes, account=None)
            if not token:
                token = app.acquire_token_for_client(scopes=scopes)
            if not token or 'access_token' not in token:
                raise RuntimeError(f"Failed to acquire Graph token: {token}")
            return token['access_token']
        # Fallback to device code flow (interactive)
        if self.graph_client_id:
            app = msal.PublicClientApplication(client_id=self.graph_client_id, authority=f'https://login.microsoftonline.com/{self.graph_tenant or "common"}')
            flow = app.initiate_device_flow(scopes=["Mail.Send"])
            if 'user_code' not in flow:
                raise RuntimeError('Failed to start device code flow')
            print(flow['message'])
            token = app.acquire_token_by_device_flow(flow)
            if not token or 'access_token' not in token:
                raise RuntimeError(f"Failed to acquire token via device flow: {token}")
            return token['access_token']
        raise RuntimeError('Graph configuration missing. Set GRAPH_CLIENT_ID (+ client secret) in environment')

    def _send_graph(self, to_email, subject, body_text, attachments, body_html, cc):
        # Acquire token
        token = self._acquire_graph_token()
        headers = {
            'Authorization': f'Bearer {token}',
            'Content-Type': 'application/json'
        }

        # Build recipients
        def recip_obj(addr):
            return {"emailAddress": {"address": addr}}

        message = {
            "subject": subject or "",
            "body": {"contentType": "html" if body_html else "text", "content": body_html or body_text or ""},
            "toRecipients": [recip_obj(to_email)] if to_email else [],
        }
        if cc:
            message["ccRecipients"] = [recip_obj(c) for c in cc]

        # Attach files as fileAttachment with base64 content
        if attachments:
            atts = []
            for path in attachments:
                if not os.path.exists(path):
                    continue
                with open(path, 'rb') as f:
                    data = f.read()
                content_b64 = base64.b64encode(data).decode('ascii')
                ctype, _ = mimetypes.guess_type(path)
                atts.append({
                    "@odata.type": "#microsoft.graph.fileAttachment",
                    "name": os.path.basename(path),
                    "contentBytes": content_b64,
                })
            if atts:
                message['attachments'] = atts

        payload = {"message": message, "saveToSentItems": True}

        # Send as the configured sender
        sender = (self.from_email or self.outlook_sender or '').strip()
        if sender:
            url = f'https://graph.microsoft.com/v1.0/users/{sender}/sendMail'
        else:
            url = 'https://graph.microsoft.com/v1.0/me/sendMail'

        resp = requests.post(url, headers=headers, data=json.dumps(payload), timeout=30)
        if resp.status_code >= 400:
            raise RuntimeError(f'Graph sendMail failed: {resp.status_code} {resp.text}')

    # ---------- SMTP backend ----------
    def _connect_smtp(self):
        if not self.user or not self.pw:
            raise ValueError("SMTP_USER/SMTP_PASS not set. Check your .env.")
        if self.use_ssl:
            ctx = ssl.create_default_context()
            srv = smtplib.SMTP_SSL(self.host, self.port or 465, context=ctx, timeout=30)
        else:
            srv = smtplib.SMTP(self.host, self.port or 587, timeout=30)
            srv.ehlo(); srv.starttls(context=ssl.create_default_context())
        srv.login(self.user, self.pw)
        return srv

    def _send_smtp(self, to_email, subject, body_text, attachments, body_html, cc):
        msg = EmailMessage()
        msg["Subject"] = subject
        msg["From"] = f"{self.from_name} <{self.from_email}>" if self.from_name else self.from_email
        # Ensure to_email is a valid string
        if not to_email:
            raise ValueError("to_email must be provided for SMTP send")
        msg["To"] = str(to_email)
        # Safely handle cc list (filter None / empty and coerce to str)
        cc_list = [str(c) for c in (cc or []) if c]
        if cc_list:
            try:
                msg["Cc"] = ",".join(cc_list)
            except Exception:
                # Fallback: join after coercion
                msg["Cc"] = ",".join([str(x) for x in cc_list if x is not None])

        if body_html:
            msg.set_content(body_text or "")
            msg.add_alternative(body_html, subtype="html")
        else:
            msg.set_content(body_text or "")

        for path in attachments:
            try:
                if not path or not os.path.exists(path):
                    continue
                ctype, _ = mimetypes.guess_type(path)
                if not ctype:
                    ctype = "application/octet-stream"
                maintype, subtype = ctype.split("/", 1)
                with open(path, "rb") as f:
                    data = f.read()
                msg.add_attachment(data, maintype=maintype, subtype=subtype, filename=os.path.basename(path) or "attachment")
            except Exception:
                # skip problematic attachment
                continue

        # Build recipients (filter empty/None)
        recipients = [str(to_email)] + [str(b) for b in (self.bcc or []) if b] + cc_list
        recipients = [r for r in recipients if r]
        with self._connect_smtp() as s:
            s.send_message(msg, from_addr=self.from_email, to_addrs=list(dict.fromkeys(recipients)))

    # ---------- Outlook desktop backend (Windows only) ----------
    def _send_outlook(self, to_email, subject, body_text, attachments, body_html, cc):
        import win32com.client as win32
        import pythoncom, time
        pythoncom.CoInitialize()
        app = win32.Dispatch("Outlook.Application")
        mail = app.CreateItem(0)  # olMailItem

        # Choose account if requested
        desired = (self.outlook_sender or "").lower()
        try:
            if desired:
                accounts = app.Session.Accounts
                for i in range(1, accounts.Count + 1):
                    acct = accounts.Item(i)
                    if acct.SmtpAddress and acct.SmtpAddress.lower() == desired:
                        # SendUsingAccount
                        mail._oleobj_.Invoke(0x0000F006, 0, 8, 0, acct)
                        break
        except Exception:
            # fallback: default account
            pass

        mail.To = to_email
        if cc:
            mail.CC = ";".join(cc)
        mail.Subject = subject
        if body_html:
            # Outlook needs <html> content; simplest is wrap a minimal html if needed
            mail.HTMLBody = body_html
        else:
            mail.Body = body_text or ""

        for path in attachments:
            try:
                mail.Attachments.Add(os.path.abspath(path))
            except Exception:
                pass

        # To *view* before sending, use .Display(); to send immediately, .Send()
        mail.Send()
        time.sleep(0.1)
