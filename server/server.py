"""
DES Comparator - Servidor Local para envio de emails
Ejecutar: python server.py
"""

from flask import Flask, request, jsonify, Response
from flask_cors import CORS
import win32com.client
import time
import requests

try:
    from playwright.sync_api import sync_playwright
    PLAYWRIGHT_AVAILABLE = True
except ImportError:
    PLAYWRIGHT_AVAILABLE = False

import os
import threading
import uuid
import base64
import tempfile
SP_FORM_URL = ('https://ts.accenture.com/sites/AcnEntArch/DevOps/_layouts/15/listform.aspx'
               '?PageType=8&ListId=%7BF056F064%2D869C%2D43B9%2DBB0C%2D9B7AEECE31C0%7D&RootFolder=')
# Perfil dedicado para automatización (Edge, directorio separado para no bloquear el Edge del usuario)
SP_AUTO_PROFILE = os.path.join(os.path.dirname(__file__), 'edge_auto_profile')
sp_jobs = {}  # job_id -> {progress, message, done, error, filled}


def _has_session():
    """Retorna True si el perfil de automatización ya fue creado (sesión previa)."""
    return os.path.isdir(os.path.join(SP_AUTO_PROFILE, 'Default'))


def _launch_managed_browser(playwright):
    """Lanza Edge con perfil dedicado de automatización (necesario para Accenture SSO)."""
    try:
        context = playwright.chromium.launch_persistent_context(
            user_data_dir=SP_AUTO_PROFILE,
            channel='msedge',
            headless=False,
            ignore_default_args=['--enable-automation', '--no-sandbox'],
            timeout=25000
        )
        return context, None
    except Exception as e:
        err = str(e)
        if ('already in use' in err.lower() or 'singleton' in err.lower()
                or 'user data' in err.lower() or 'timeout' in err.lower()):
            return None, 'BROWSER_BUSY'
        return None, err

app = Flask(__name__)
CORS(app)  # Permite requests desde comparator.html

# Links
LINKS = {
    "learning_pm": "https://in.accenture.com/cioorganization/cio-bootcamp/cio-bootcamp-roles/",
    "learning_pam": "https://in.accenture.com/cioorganization/cio-bootcamp/cio-bootcamp-roles/",
    "learning_po": "https://in.accenture.com/cioorganization/cio-bootcamp/cio-bootcamp-roles/",
    "learning_sm": "https://in.accenture.com/cioorganization/cio-bootcamp/cio-bootcamp-roles/",
    "learning_rte": "https://in.accenture.com/cioorganization/cio-bootcamp/cio-bootcamp-roles/",
    "learning_ste": "https://in.accenture.com/cioorganization/cio-bootcamp/cio-bootcamp-roles/",
    "learning_default": "https://in.accenture.com/cioorganization/cio-bootcamp/cio-bootcamp-roles/",
    "new_ado_request": "https://ts.accenture.com/sites/AcnEntArch/DevOps/Lists/CIO%20Azure%20DevOps%20Request%20Form%2020/AllItems.aspx?InitialTabId=Ribbon%2EListItem&VisibilityContext=WSSTabPersistence",
    "ad_group": "https://directory.accenture.com/WebAdmin/search.aspx",
    "ja_onboarding": "https://ts.accenture.com/sites/JiraAlignRollout/SitePages/Request%20Access%20Onboarding.aspx",
    "roles_itba": "https://in.accenture.com/cioorganization/cio-operating-model/it-business-agility/de-hub/roles/",
    "agile_home": "https://in.accenture.com/cioorganization/agile/",
    "itba": "https://in.accenture.com/cioorganization/cio-operating-model/it-business-agility/",
    "ado_guidelines": "https://in.accenture.com/cioorganization/agile/guidelines-ado/",
    "jira_align_site": "https://ts.accenture.com/sites/JiraAlignRollout/SitePages/Welcome-to-Jira-Align.aspx",
}

# Email de contacto
CONTACT_EMAIL = "cio.itba.des.cell.powerpuff@accenture.com"
DEFAULT_CC = "CIO.ITBA.DES.Cell.PowerPuff@accenture.com"


def get_learning_link(role):
    """Obtiene el link de learning path segun el rol"""
    role_lower = role.lower()
    if "product manager" in role_lower and "area" not in role_lower:
        return LINKS["learning_pm"]
    elif "product area manager" in role_lower:
        return LINKS["learning_pam"]
    elif "product owner" in role_lower:
        return LINKS["learning_po"]
    elif "scrum master" in role_lower:
        return LINKS["learning_sm"]
    elif "release train engineer" in role_lower:
        return LINKS["learning_rte"]
    elif "solution train engineer" in role_lower:
        return LINKS["learning_ste"]
    else:
        return LINKS["learning_default"]


def get_html_template_new_product_pm(name, product_id, product_name=""):
    """Template para PM de producto NUEVO"""
    product_display = f"{product_id} ({product_name})" if product_name else product_id
    return f"""
<html>
<body style="font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; font-size: 14px; color: #333; line-height: 1.6;">
    <p>Dear {name},</p>

    <p>Congratulations on your assignment to the role of <strong>Product Manager</strong> for the product <strong>{product_display}</strong>!</p>

    <p>To help you get started, we've developed a customized learning path to help you get familiar with your new responsibilities. You can start your learning journey <a href="{LINKS['learning_pm']}" style="color: #A100FF;">here</a>.</p>

    <p>As part of your new role position you may need JA access. Please review and follow the instructions<br>
    <a href="{LINKS['ja_onboarding']}" style="color: #A100FF;">Jira Align Product Onboarding Job Aid</a></p>

    <p>For more information on ITBA roles and responsibilities<br>
    <a href="{LINKS['roles_itba']}" style="color: #A100FF;">Roles | ITBA | CIO Organization</a></p>

    <p>For additional training and information please review the following sites:</p>
    <ul style="margin: 10px 0;">
        <li><a href="{LINKS['agile_home']}" style="color: #A100FF;">Agile Home | CIO Organization</a></li>
        <li><a href="{LINKS['itba']}" style="color: #A100FF;">IT Business Agility | CIO Organization</a></li>
        <li><a href="{LINKS['ado_guidelines']}" style="color: #A100FF;">Azure DevOps Guidelines</a></li>
        <li><a href="{LINKS['jira_align_site']}" style="color: #A100FF;">Jira Align Site</a></li>
    </ul>

    <p>If you are receiving this email, it is because someone responsible for the designated product has added your name in AIR for the specified role. Should you have any questions about the assignment, please reach out to the Product Manager (copied) or, alternatively, to the Product Area Manager as appropriate. For any other questions or inquiries, please submit a ticket to <a href="https://support.accenture.com/support_portal?id=it_services_order&articleNumber=KB0072225&sys_id=cafd77e3dbab3780b3535eea4b96198d&sysold=7ee09b61131aa600380ddbf18144b032&sysparm_load=181285" style="color: #A100FF;">Accenture Support</a>.</p>

    <p>Best Regards,<br>
    <strong>IT Business Agility – Delivery Enablement Solutions</strong></p>
</body>
</html>
"""


def get_html_template_standard(name, role, product_id, product_name=""):
    """Template estandar para PM existente, PAM, PO, SM, RTE, STE, etc."""
    learning_link = get_learning_link(role)
    role_display = role.replace("(s)", "").strip()
    product_display = f"{product_id} ({product_name})" if product_name else product_id

    return f"""
<html>
<body style="font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; font-size: 14px; color: #333; line-height: 1.6;">
    <p>Dear {name},</p>

    <p>Congratulations on your assignment to the role of <strong>{role_display}</strong> for the product <strong>{product_display}</strong>!</p>

    <p>To help you get started, we've developed a customized learning path to help you get familiar with your new responsibilities. You can start your learning journey <a href="{learning_link}" style="color: #A100FF;">here</a>.</p>

    <p>As part of your new role position you may need JA access. Please review and follow the instructions<br>
    <a href="{LINKS['ja_onboarding']}" style="color: #A100FF;">Jira Align Product Onboarding Job Aid</a></p>

    <p>For more information on ITBA roles and responsibilities<br>
    <a href="{LINKS['roles_itba']}" style="color: #A100FF;">Roles | ITBA | CIO Organization</a></p>

    <p>For additional training and information please review the following sites:</p>
    <ul style="margin: 10px 0;">
        <li><a href="{LINKS['agile_home']}" style="color: #A100FF;">Agile Home | CIO Organization</a></li>
        <li><a href="{LINKS['itba']}" style="color: #A100FF;">IT Business Agility | CIO Organization</a></li>
        <li><a href="{LINKS['ado_guidelines']}" style="color: #A100FF;">Azure DevOps Guidelines</a></li>
        <li><a href="{LINKS['jira_align_site']}" style="color: #A100FF;">Jira Align Site</a></li>
    </ul>

    <p>If you are receiving this email, it is because someone responsible for the designated product has added your name in AIR for the specified role. Should you have any questions about the assignment, please reach out to the Product Manager (copied) or, alternatively, to the Product Area Manager as appropriate. For any other questions or inquiries, please submit a ticket to <a href="https://support.accenture.com/support_portal?id=it_services_order&articleNumber=KB0072225&sys_id=cafd77e3dbab3780b3535eea4b96198d&sysold=7ee09b61131aa600380ddbf18144b032&sysparm_load=181285" style="color: #A100FF;">Accenture Support</a>.</p>

    <p>Best Regards,<br>
    <strong>IT Business Agility – Delivery Enablement Solutions</strong></p>
</body>
</html>
"""


def get_html_template(name, role, product_id, product_name="", is_new_product=False):
    """Selecciona el template correcto segun rol y si es producto nuevo"""
    role_lower = role.lower()
    is_pm = "product manager" in role_lower and "area" not in role_lower

    if is_pm and is_new_product:
        return get_html_template_new_product_pm(name, product_id, product_name)
    else:
        return get_html_template_standard(name, role, product_id, product_name)



def send_single_email(email_data):
    """Envia un email usando Outlook COM"""
    import pythoncom
    tmp_path = None
    pythoncom.CoInitialize()
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        mail = outlook.CreateItem(0)

        # Use Recipients.Add() instead of mail.To to force SMTP address resolution
        # and avoid Outlook resolving the address against GAL/autocomplete cache,
        # which can silently redirect to a different person with a similar name.
        def add_smtp_recipients(addresses_str, recipient_type):
            for addr in [a.strip() for a in addresses_str.split(';') if a.strip()]:
                recip = mail.Recipients.Add(addr)
                recip.Type = recipient_type
                recip.Resolve()
                # Verify Outlook didn't silently remap to a different address
                try:
                    resolved = recip.AddressEntry.GetExchangeUser().PrimarySmtpAddress
                    if resolved.lower() != addr.lower():
                        raise Exception(
                            f"Outlook resolved '{addr}' to '{resolved}' — aborting to prevent wrong delivery"
                        )
                except AttributeError:
                    pass  # Non-Exchange address, trust as-is

        add_smtp_recipients(email_data["to"], 1)   # 1 = olTo
        if email_data.get("cc"):
            add_smtp_recipients(email_data["cc"], 2)  # 2 = olCC
        mail.Subject = email_data["subject"]

        html_body = email_data["html_body"]
        # Wrap partial HTML snippets in a proper document structure so Outlook
        # renders paragraph spacing and fonts correctly.
        if html_body and '<html' not in html_body.lower():
            html_body = (
                '<html><body style="font-family:\'Segoe UI\',Tahoma,Geneva,Verdana,sans-serif;'
                'font-size:14px;color:#333;line-height:1.6;">'
                + html_body +
                '</body></html>'
            )

        mail.BodyFormat = 2  # olFormatHTML – must be set before HTMLBody
        mail.HTMLBody = html_body

        # Attachment (optional) — base64-encoded file sent from the frontend
        att = email_data.get("attachment")
        if att and att.get("data") and att.get("filename"):
            tmp_path = os.path.join(tempfile.gettempdir(), att["filename"])
            with open(tmp_path, "wb") as f:
                f.write(base64.b64decode(att["data"]))
            mail.Attachments.Add(tmp_path)

        mail.Send()
        return {"success": True, "to": email_data["to"]}

    except Exception as e:
        return {"success": False, "to": email_data["to"], "error": str(e)}
    finally:
        if tmp_path and os.path.exists(tmp_path):
            try:
                os.remove(tmp_path)
            except Exception:
                pass
        pythoncom.CoUninitialize()


@app.route("/health", methods=["GET"])
def health():
    """Endpoint para verificar que el servidor esta corriendo"""
    return jsonify({"status": "ok", "message": "DES Email Server running"})


@app.route("/send", methods=["POST"])
def send_emails():
    """
    Endpoint para enviar emails

    Espera JSON con formato:
    {
        "emails": [
            {
                "to": "email@accenture.com",
                "cc": "manager@accenture.com",
                "name": "Carlos",
                "role": "Product Manager",
                "productId": "PR1291",
                "isNewProduct": true
            },
            ...
        ]
    }
    """
    try:
        data = request.get_json()

        if not data or "emails" not in data:
            return jsonify({"success": False, "error": "No emails provided"}), 400

        emails = data["emails"]
        results = {"sent": 0, "failed": 0, "details": []}

        print(f"\n{'='*50}")
        print(f"Recibidos {len(emails)} emails para enviar")
        print(f"{'='*50}\n")

        for idx, email_info in enumerate(emails):
            # Extraer datos
            to = email_info.get("to", "")
            cc = email_info.get("cc", DEFAULT_CC)
            name = email_info.get("name", "")
            role = email_info.get("role", "")
            product_id = email_info.get("productId", "")
            product_name = email_info.get("productName", "")
            is_new_product = email_info.get("isNewProduct", False)

            subject = email_info.get("subject") or "Welcome to your new role!"
            raw_html = email_info.get("htmlBody", "")
            html_body = raw_html or get_html_template(name, role, product_id, product_name, is_new_product)
            template_source = "FRONTEND" if raw_html else "PYTHON_FALLBACK"

            email_data = {
                "to": to,
                "cc": cc,
                "subject": subject,
                "html_body": html_body,
                "attachment": email_info.get("attachment")
            }

            print(f"[{idx+1}/{len(emails)}] To: {to} | Subject: {subject} | Template: {template_source} ({len(html_body)} chars)")

            result = send_single_email(email_data)

            if result["success"]:
                results["sent"] += 1
                results["details"].append({"to": to, "status": "sent"})
                print(f"         OK")
            else:
                error_msg = result.get("error", "Unknown")
                results["failed"] += 1
                results["details"].append({"to": to, "status": "failed", "error": error_msg})
                print(f"         ERROR: {error_msg}")

            # Pequeno delay para no saturar Outlook
            time.sleep(0.3)

        print(f"\n{'='*50}")
        print(f"RESUMEN: {results['sent']} enviados, {results['failed']} fallidos")
        print(f"{'='*50}\n")

        return jsonify({
            "success": True,
            "sent": results["sent"],
            "failed": results["failed"],
            "details": results["details"]
        })

    except Exception as e:
        print(f"ERROR: {str(e)}")
        return jsonify({"success": False, "error": str(e)}), 500


@app.route("/send-single", methods=["POST"])
def send_single():
    """
    Endpoint para enviar UN solo email
    """
    try:
        email_info = request.get_json()

        if not email_info or "to" not in email_info:
            return jsonify({"success": False, "error": "No email data provided"}), 400

        to = email_info.get("to", "")
        cc = email_info.get("cc", DEFAULT_CC)
        name = email_info.get("name", "")
        role = email_info.get("role", "")
        product_id = email_info.get("productId", "")
        product_name = email_info.get("productName", "")
        is_new_product = email_info.get("isNewProduct", False)

        subject = "Welcome to your new role!"
        html_body = email_info.get("htmlBody") or get_html_template(name, role, product_id, product_name, is_new_product)

        email_data = {
            "to": to,
            "cc": cc,
            "subject": subject,
            "html_body": html_body
        }

        print(f"Enviando email a {to}...")

        result = send_single_email(email_data)

        if result["success"]:
            print(f"OK - Email enviado a {to}")
            return jsonify({"success": True, "to": to})
        else:
            print(f"ERROR - {result.get('error', 'Unknown')}")
            return jsonify({"success": False, "to": to, "error": result.get("error", "Unknown")}), 500

    except Exception as e:
        print(f"ERROR: {str(e)}")
        return jsonify({"success": False, "error": str(e)}), 500


@app.route("/preview", methods=["POST"])
def preview_email():
    """
    Endpoint para obtener preview del email sin enviarlo

    Espera JSON con formato:
    {
        "to": "email@accenture.com",
        "cc": "manager@accenture.com",
        "name": "Carlos",
        "role": "Product Manager",
        "productId": "PR1291",
        "isNewProduct": true
    }

    Retorna:
    {
        "success": true,
        "to": "email@accenture.com",
        "cc": "manager@accenture.com",
        "subject": "Welcome to your new role!",
        "htmlBody": "<html>..."
    }
    """
    try:
        email_info = request.get_json()

        if not email_info:
            return jsonify({"success": False, "error": "No email data provided"}), 400

        to = email_info.get("to", "")
        cc = email_info.get("cc", DEFAULT_CC)
        name = email_info.get("name", "")
        role = email_info.get("role", "")
        product_id = email_info.get("productId", "")
        product_name = email_info.get("productName", "")
        is_new_product = email_info.get("isNewProduct", False)

        subject = "Welcome to your new role!"
        # Usar el htmlBody del frontend si viene; sino generarlo con el template Python
        html_body = email_info.get("htmlBody") or get_html_template(name, role, product_id, product_name, is_new_product)

        return jsonify({
            "success": True,
            "to": to,
            "cc": cc,
            "subject": subject,
            "htmlBody": html_body
        })

    except Exception as e:
        print(f"ERROR: {str(e)}")
        return jsonify({"success": False, "error": str(e)}), 500


def _get_frame(page):
    """Retorna el frame con mas inputs de texto (donde vive el formulario InfoPath)."""
    frames = page.frames
    print(f"[SP Form] Frames disponibles: {[f.url for f in frames]}")
    best_frame, best_count = page, 0
    for frame in frames:
        try:
            count = frame.evaluate("() => document.querySelectorAll('input[type=\"text\"]').length")
            print(f"[SP Form] Frame {frame.url[:80]} — {count} text inputs")
            if count > best_count:
                best_count = count
                best_frame = frame
        except Exception:
            pass
    print(f"[SP Form] Usando frame con {best_count} inputs: {best_frame.url[:80]}")
    return best_frame


def _fill_people_picker(frame, page, label_text, value):
    """Finds a SharePoint People Picker div by adjacent label and types the value to resolve it."""
    FIND_PP_JS = """(labelText) => {
        const divs = document.querySelectorAll('div[contenteditable="true"][title="People Picker"]');
        for (const div of divs) {
            let node = div.parentElement;
            for (let d = 0; d < 8; d++) {
                if (!node) break;
                const prev = node.previousElementSibling;
                if (prev && prev.textContent.indexOf(labelText) !== -1) return div;
                const par = node.parentElement;
                if (par) {
                    const kids = Array.from(par.children);
                    const idx = kids.indexOf(node);
                    if (idx > 0 && kids[idx-1].textContent.indexOf(labelText) !== -1) return div;
                }
                node = node.parentElement;
            }
        }
        return divs[0] || null;
    }"""
    try:
        handle = frame.evaluate_handle(FIND_PP_JS, label_text)
        el = handle.as_element()
        if not el:
            print(f"[SP Form] --  People Picker not found (label='{label_text}')")
            return False
        el.click()
        el.type(value)          # character-by-character → triggers onkeyup (autocomplete)
        page.wait_for_timeout(1000)
        # Try to click first autocomplete suggestion
        suggestion = frame.query_selector(
            '.ms-PeoplePicker-result, .sp-peoplepicker-suggestion, '
            '[class*="peoplepicker"] li:first-child, .ms-suggestionItem, '
            'div[class*="Suggestions"] li:first-child'
        )
        if suggestion:
            suggestion.click()
            print(f"[SP Form] OK  People Picker '{label_text}' — selected suggestion")
        else:
            # No dropdown visible — press Enter to attempt resolution (cap wait at 10s)
            page.set_default_timeout(10000)
            try:
                page.keyboard.press('Enter')
            except Exception:
                pass
            finally:
                page.set_default_timeout(30000)
            print(f"[SP Form] OK  People Picker '{label_text}' — pressed Enter to resolve")
        return True
    except Exception as e:
        print(f"[SP Form] --  People Picker '{label_text}': {e}")
        return False


def _sp_fill_fields(page, fields):
    """Completa los campos del formulario usando label adyacente (InfoPath no tiene title/aria-label)."""
    frame = _get_frame(page)
    filled = 0

    # Fields where the key doesn't match the label text in the form
    LABEL_OVERRIDE = {
        'Shared with a Sarbanes-Oxley governed application?': ['SOX', 'Sarbanes-Oxley', 'Sarbanes'],
        'Active Directory Group Names': ['Active Directory Group Names', 'AD Group Names', 'Active Directory Groups', 'AD Groups'],
    }

    # Mirrors findFieldByLabel from content.js
    FIND_BY_LABEL_JS = """(labelText) => {
        const candidates = document.querySelectorAll(
            'input[type="text"], input[type="search"], textarea:not([name*="downlevel"]), select'
        );
        for (const el of candidates) {
            let node = el.parentElement;
            for (let d = 0; d < 8; d++) {
                if (!node) break;
                const prev = node.previousElementSibling;
                if (prev && prev.textContent.indexOf(labelText) !== -1) return el;
                const par = node.parentElement;
                if (par) {
                    const kids = Array.from(par.children);
                    const idx = kids.indexOf(node);
                    if (idx > 0 && kids[idx - 1].textContent.indexOf(labelText) !== -1) return el;
                }
                node = node.parentElement;
            }
        }
        return null;
    }"""

    PEOPLE_PICKERS = {'Project Contact Person'}
    SKIP_KEYS = {'Shared with a Sarbanes-Oxley governed application?'}
    deferred_pp = {}

    for key, value in fields.items():
        if not value or key in SKIP_KEYS:
            continue
        if key in PEOPLE_PICKERS:
            deferred_pp[key] = value
            continue
        # Product ID: numeric only
        if key == 'Product ID' and str(value).upper().startswith('PR'):
            value = str(value)[2:]

        labels = LABEL_OVERRIDE.get(key, [key])
        if isinstance(labels, str):
            labels = [labels]

        try:
            el = None
            for label in labels:
                # 1. Standard attribute selectors
                for sel in [f'input[title="{label}"]', f'textarea[title="{label}"]',
                            f'select[title="{label}"]', f'input[aria-label="{label}"]',
                            f'textarea[aria-label="{label}"]', f'select[aria-label="{label}"]']:
                    el = frame.query_selector(sel)
                    if el:
                        break
                # 2. Label-based JS lookup
                if not el:
                    handle = frame.evaluate_handle(FIND_BY_LABEL_JS, label)
                    el = handle.as_element()
                if el:
                    break

            if el:
                tag = el.evaluate("el => el.tagName.toLowerCase()")
                if tag == 'select':
                    try:
                        el.select_option(label=str(value))
                    except Exception:
                        el.select_option(value=str(value))
                else:
                    try:
                        el.fill(str(value))
                    except Exception:
                        el.click()
                        el.press_sequentially(str(value))
                filled += 1
                print(f"[SP Form] OK  {key} = {str(value)[:60]}")
            else:
                print(f"[SP Form] --  {key}: field not found (tried labels: {labels})")
        except Exception as fe:
            print(f"[SP Form] --  {key}: {fe}")

    # Fill People Pickers last
    for key, value in deferred_pp.items():
        email = str(value) if '@' in str(value) else str(value) + '@accenture.com'
        if _fill_people_picker(frame, page, key, email):
            filled += 1
    return filled


def _sp_save_form(page):
    """Hace click en el boton Save del formulario (ribbon de SharePoint clasico)."""
    frame = _get_frame(page)

    selectors = [
        'input[id*="SaveButton"]',
        'input[value="Save"]',
        'a[id*="SaveButton"]',
        'a[id*="diidIOSaveItem"]',
        '[id*="Ribbon.ListForm.Edit.Commit.Publish"]',
        '[aria-label="Save"]',
        '[title="Save"]',
    ]
    for sel in selectors:
        el = frame.query_selector(sel)
        if el:
            print(f"[SP Form] Save con selector: {sel}")
            el.click()
            page.wait_for_load_state('networkidle', timeout=30000)
            return

    for tag in ['a', 'input', 'button', 'span']:
        el = frame.query_selector(f'{tag}:text-is("Save")')
        if el:
            print(f"[SP Form] Save por texto en <{tag}>")
            el.click()
            page.wait_for_load_state('networkidle', timeout=30000)
            return

    # Debug: screenshot + listar todos los IDs
    screenshot_path = os.path.join(os.path.dirname(__file__), 'debug_form.png')
    page.screenshot(path=screenshot_path)
    print(f"[SP Form] Screenshot guardado en {screenshot_path}")
    all_ids = [e.get_attribute('id') for e in frame.query_selector_all('[id]') if e.get_attribute('id')]
    print(f"[SP Form] Todos los IDs en el frame ({len(all_ids)}): {all_ids[:20]}")
    raise Exception(f"No se encontró el botón Save. Ver screenshot: {screenshot_path}")


@app.route("/sp-login", methods=["POST"])
def sp_login():
    """
    Abre Edge con perfil dedicado para que el usuario inicie sesion en SharePoint.
    La sesion queda guardada en el perfil (edge_auto_profile/).
    """
    if not PLAYWRIGHT_AVAILABLE:
        return jsonify({"success": False, "error": "Falta Playwright."}), 500

    try:
        print(f"\n[SP Login] Abriendo Edge dedicado para login...")
        with sync_playwright() as p:
            context, err = _launch_managed_browser(p)
            if err == 'BROWSER_BUSY':
                return jsonify({"success": False, "error": "BROWSER_BUSY"}), 409
            if err:
                return jsonify({"success": False, "error": err}), 500

            page = context.new_page()
            print(f"[SP Login] Navegando a SharePoint...")
            page.goto(SP_FORM_URL, wait_until='load', timeout=60000)
            print(f"[SP Login] URL: {page.url}")

            if 'login.microsoftonline' in page.url or 'login.live' in page.url:
                print(f"[SP Login] Login requerido, esperando hasta 3 min...")
                try:
                    page.wait_for_selector('input[title="Product Name"], input[title="Product ID"]', timeout=180000)
                except Exception:
                    print(f"[SP Login] Timeout, URL actual: {page.url}")

            print(f"[SP Login] Sesion guardada en perfil dedicado.")
            context.close()

        return jsonify({"success": True, "message": "Sesion guardada correctamente."})
    except Exception as e:
        print(f"[SP Login] ERROR: {str(e)}")
        return jsonify({"success": False, "error": str(e)}), 500


def _run_fill_form(job_id, fields):
    """Ejecuta el llenado del formulario en un hilo separado."""
    job = sp_jobs[job_id]

    def update(progress, message):
        job['progress'] = progress
        job['message'] = message
        print(f"[SP Form] ({progress}%) {message}")

    def start_ticker(from_pct, to_pct, seconds):
        """Avanza el progreso lentamente en background mientras se espera una operacion larga."""
        cancel_event = threading.Event()
        def _tick():
            steps = max(2, int(seconds * 2))
            for i in range(steps):
                if cancel_event.is_set():
                    return
                pct = from_pct + int((to_pct - from_pct) * (i + 1) / steps)
                job['progress'] = pct
                time.sleep(0.5)
        threading.Thread(target=_tick, daemon=True).start()
        return cancel_event

    def chrome_status(page, pct, msg):
        """Inyecta/actualiza un banner de progreso en la ventana de Chrome."""
        try:
            page.evaluate(f"""() => {{
                let el = document.getElementById('_des_ovl');
                if (!el) {{
                    el = document.createElement('div');
                    el.id = '_des_ovl';
                    el.style.cssText = 'position:fixed;top:0;left:0;right:0;z-index:2147483647;'
                        + 'background:rgba(20,20,40,0.97);color:white;padding:10px 18px;'
                        + 'font-family:Segoe UI,sans-serif;font-size:13px;'
                        + 'display:flex;align-items:center;gap:10px;box-shadow:0 2px 8px rgba(0,0,0,.5);';
                    document.body && document.body.appendChild(el);
                }}
                el.innerHTML = '<b style="color:#A100FF;margin-right:4px;">DES</b>'
                    + '<span style="background:#A100FF;padding:1px 8px;border-radius:10px;font-size:11px;">{pct}%</span>'
                    + '<span>{msg}</span>';
            }}""")
        except Exception:
            pass

    try:
        update(5, 'Iniciando proceso...')
        with sync_playwright() as p:
            update(12, 'Lanzando Edge con perfil de automatización...')
            ticker = start_ticker(12, 22, 8)
            context, err = _launch_managed_browser(p)
            ticker.set()

            if err == 'BROWSER_BUSY':
                job['error'] = 'BROWSER_BUSY'
                job['done'] = True
                return
            if err:
                job['error'] = err
                job['done'] = True
                return

            update(23, 'Edge abierto, abriendo pestaña...')
            page = context.new_page()

            update(28, 'Navegando al formulario de SharePoint...')
            # 'load' evita colgar esperando network-idle (SharePoint tiene requests continuos)
            ticker = start_ticker(28, 62, 60)
            page.goto(SP_FORM_URL, wait_until='load', timeout=90000)
            ticker.set()

            update(63, 'Página cargada, verificando sesión...')

            if 'login.microsoftonline' in page.url or 'login.live' in page.url:
                context.close()
                job['error'] = 'SESSION_EXPIRED'
                job['done'] = True
                return

            # Wait for InfoPath to finish rendering dropdowns and custom controls
            update(63, 'Esperando que cargue InfoPath...')
            page.wait_for_timeout(4000)

            update(65, 'Completando campos del formulario...')
            chrome_status(page, 65, 'Completando campos...')
            filled = _sp_fill_fields(page, fields)

            chrome_status(page, 100, f'Done — {filled} fields filled. Save the form and close this window.')
            job['filled'] = filled
            job['progress'] = 100
            job['message'] = f'Fields filled ({filled}). Save the form manually.'
            job['done'] = True
            print(f"[SP Form] Listo! {filled} campos. Ventana abierta — esperando cierre manual.")
            # Keep browser open until user closes it manually (max 15 min)
            time.sleep(900)

    except Exception as e:
        print(f"[SP Form] ERROR: {str(e)}")
        job['error'] = str(e)
        job['done'] = True


@app.route("/fill-form", methods=["POST"])
def fill_sp_form():
    """Inicia el llenado del formulario en background y retorna un job_id."""
    if not PLAYWRIGHT_AVAILABLE:
        return jsonify({
            "success": False,
            "error": "Falta Playwright. Ejecuta: pip install playwright && playwright install chromium"
        }), 500

    if not _has_session():
        return jsonify({"success": False, "error": "NO_SESSION"}), 401

    data = request.get_json()
    if not data or "fields" not in data:
        return jsonify({"success": False, "error": "No fields provided"}), 400

    job_id = str(uuid.uuid4())
    sp_jobs[job_id] = {'progress': 0, 'message': 'Iniciando...', 'done': False, 'error': None, 'filled': 0}

    thread = threading.Thread(target=_run_fill_form, args=(job_id, data["fields"]))
    thread.daemon = True
    thread.start()

    return jsonify({"success": True, "jobId": job_id})


@app.route("/fill-form-status/<job_id>", methods=["GET"])
def fill_form_status(job_id):
    """Retorna el estado actual de un job de llenado de formulario."""
    job = sp_jobs.get(job_id)
    if not job:
        return jsonify({"error": "Job not found"}), 404
    return jsonify(job)


AIR_EXPORT_URL = (
    'https://support.accenture.com/u_business_service_attributes_list.do?EXCEL'
    '&sysparm_query=u_product_status!=Withdrawn^ORu_product_status=NULL'
    '&sysparm_view=air_product_detailed_extract'
    '&sysparm_fields=u_product_id,u_product_name.u_product_name,u_product_status,'
    'u_registration_status,u_product_manager,u_portfolio,u_portfolio_lead,u_component,'
    'u_service_area,u_comp_lead,u_service_area_lead,u_reliability_director,u_business_owner,'
    'u_product_area_managers,u_solution_train_engineers,u_release_train_engineers,'
    'u_optional_editors,u_scrum_masters,u_product_owners,u_crossport_ste,u_product_vision,'
    'u_higher_purpose,u_value_proposition,u_how_we_measure_value,u_stakeholder_s,'
    'u_stakeholder_groups,sys_created_by,sys_created_on,u_last_modified_by,u_last_modified_on,'
    'u_problems_to_solve,u_ways_we_solve,u_owned_application_s,u_owned_application_ids,'
    'u_funding_sources_and_structure_aim_id,u_consumed_products,u_user_segment,'
    'u_non_air_solutions,u_product_obstacles,u_delivery_methodology,'
    'u_delivery_methodology_explanation,u_itba_maturity,u_'
)


def _is_excel_content(content, content_type=''):
    """Verifica que el contenido recibido sea un archivo Excel y no una página de login."""
    ct = content_type.lower()
    if 'html' in ct:
        return False
    # Firma magic de ZIP/XLSX: PK\x03\x04
    if content and content[:4] == b'PK\x03\x04':
        return True
    # Firma de XLS (BIFF8): D0 CF 11 E0
    if content and content[:4] == b'\xd0\xcf\x11\xe0':
        return True
    # Si el content-type dice Excel explícitamente
    if any(x in ct for x in ['excel', 'spreadsheet', 'octet-stream']):
        return True
    return False


def _download_air_via_sspi():
    """Intenta descargar el archivo AIR usando Windows SSO (Kerberos/NTLM)."""
    auth = None
    try:
        from requests_negotiate_sspi import HttpNegotiateAuth
        auth = HttpNegotiateAuth()
        print('[AIR] Usando HttpNegotiateAuth (Kerberos/NTLM SSPI)')
    except ImportError:
        try:
            from requests_ntlm import HttpNtlmAuth
            auth = HttpNtlmAuth('', '')
            print('[AIR] Usando HttpNtlmAuth (NTLM con usuario Windows actual)')
        except ImportError:
            raise Exception('No hay módulos SSPI instalados (requests-negotiate-sspi / requests-ntlm)')

    resp = requests.get(AIR_EXPORT_URL, auth=auth, timeout=120, allow_redirects=True)
    ct = resp.headers.get('Content-Type', '')
    print(f'[AIR] SSPI response: {resp.status_code}, Content-Type: {ct}, URL: {resp.url}')
    resp.raise_for_status()

    if not _is_excel_content(resp.content, ct):
        raise Exception(f'SSPI devolvió contenido no-Excel (Content-Type: {ct}). Sesión inválida.')

    return resp.content


def _download_air_via_edge_cookies():
    """
    Descarga el archivo AIR leyendo las cookies del Edge real del usuario.
    No necesita abrir ninguna ventana — usa la sesión activa del browser.
    Requiere: pip install browser-cookie3
    """
    try:
        import browser_cookie3
    except ImportError:
        raise Exception('Falta browser-cookie3. Instalá con: pip install browser-cookie3')

    print('[AIR] Leyendo cookies de Edge del usuario...')
    cookiejar = browser_cookie3.edge(domain_name='accenture.com')

    session = requests.Session()
    session.cookies.update(cookiejar)

    resp = session.get(AIR_EXPORT_URL, timeout=120, allow_redirects=True)
    ct = resp.headers.get('Content-Type', '')
    print(f'[AIR] Edge cookies → {resp.status_code}, Content-Type: {ct}, URL: {resp.url}')

    if not _is_excel_content(resp.content, ct):
        raise Exception(
            f'Respuesta no-Excel (Content-Type: {ct}). '
            'Asegurate de estar logueado en AIR en tu Edge principal.'
        )
    return resp.content


ONEDRIVE_FOLDER = r'C:\Users\c.san.emeterio\OneDrive - Accenture\ITBA CoE - FilesToCompare'

@app.route('/local-files', methods=['GET'])
def local_files():
    """
    Returns the two most recent Excel files from the OneDrive sync folder as base64.
    Newest file = file2 (current), second newest = file1 (previous).
    """
    if not os.path.isdir(ONEDRIVE_FOLDER):
        return jsonify({'success': False, 'error': f'Folder not found: {ONEDRIVE_FOLDER}'}), 404

    entries = []
    for fname in os.listdir(ONEDRIVE_FOLDER):
        if fname.lower().endswith(('.xlsx', '.xls')) and not fname.startswith('~$'):
            fpath = os.path.join(ONEDRIVE_FOLDER, fname)
            entries.append({'name': fname, 'path': fpath, 'mtime': os.path.getmtime(fpath)})

    entries.sort(key=lambda x: x['mtime'], reverse=True)

    if len(entries) < 2:
        return jsonify({'success': False, 'error': f'Need at least 2 Excel files, found {len(entries)}'}), 404

    f2, f1 = entries[0], entries[1]  # newest = current, second = previous

    with open(f1['path'], 'rb') as fh:
        f1_b64 = base64.b64encode(fh.read()).decode()
    with open(f2['path'], 'rb') as fh:
        f2_b64 = base64.b64encode(fh.read()).decode()

    print(f'[LocalFiles] file1={f1["name"]}  file2={f2["name"]}')
    return jsonify({
        'success': True,
        'file1': {'name': f1['name'], 'data': f1_b64},
        'file2': {'name': f2['name'], 'data': f2_b64}
    })


@app.route('/air-download', methods=['GET'])
def air_download():
    """
    Proxy para descargar el archivo AIR.
    Intenta primero con Windows SSO (SSPI/NTLM), luego con Playwright.
    Retorna el archivo Excel crudo para que el frontend lo procese.
    """
    print('\n[AIR] Solicitud de descarga recibida')
    content = None
    last_error = None

    # Intento 1: cookies del Edge real del usuario (más confiable)
    try:
        content = _download_air_via_edge_cookies()
        print(f'[AIR] OK via Edge cookies — {len(content)} bytes')
    except Exception as e:
        last_error = str(e)
        print(f'[AIR] Edge cookies falló: {e}')

    # Intento 2: Windows SSPI/NTLM (si hay módulos instalados)
    if content is None:
        try:
            content = _download_air_via_sspi()
            print(f'[AIR] OK via SSPI — {len(content)} bytes')
        except Exception as e:
            last_error = str(e)
            print(f'[AIR] SSPI falló: {e}')

    if content is None:
        print(f'[AIR] ERROR: Todos los métodos fallaron. Último error: {last_error}')
        return jsonify({'success': False, 'error': last_error}), 502

    return Response(
        content,
        status=200,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )


@app.route('/air-debug', methods=['GET'])
def air_debug():
    """Diagnóstico: muestra qué cookies encuentra y qué respuesta da AIR."""
    info = {}
    try:
        import browser_cookie3
        info['browser_cookie3'] = 'instalado'
        jar = browser_cookie3.edge(domain_name='accenture.com')
        cookies = {c.name: c.value[:10] + '...' for c in jar}
        info['cookies_encontradas'] = len(cookies)
        info['cookie_names'] = list(cookies.keys())[:10]
    except Exception as e:
        info['browser_cookie3_error'] = str(e)

    try:
        import browser_cookie3
        jar = browser_cookie3.edge(domain_name='accenture.com')
        session = requests.Session()
        session.cookies.update(jar)
        session.headers['User-Agent'] = (
            'Mozilla/5.0 (Windows NT 10.0; Win64; x64) '
            'AppleWebKit/537.36 (KHTML, like Gecko) '
            'Chrome/136.0.0.0 Safari/537.36 Edg/136.0.0.0'
        )
        resp = session.get(AIR_EXPORT_URL, timeout=30, allow_redirects=True)
        ct = resp.headers.get('Content-Type', '')
        info['status'] = resp.status_code
        info['content_type'] = ct
        info['final_url'] = resp.url
        info['content_start'] = resp.content[:200].hex()
        info['is_excel'] = _is_excel_content(resp.content, ct)
    except Exception as e:
        info['request_error'] = str(e)

    return jsonify(info)


if __name__ == "__main__":
    print("\n" + "="*50)
    print("DES Comparator - Email Server")
    print("="*50)
    print("Servidor corriendo en: http://localhost:5000")
    print("Endpoints:")
    print("  GET  /health      - Verificar estado")
    print("  POST /send        - Enviar multiples emails")
    print("  POST /send-single - Enviar un email")
    print("  POST /fill-form   - Completar formulario Azure DevOps en SharePoint")
    print("  GET  /air-download - Descargar archivo AIR (SSO automático via Edge)")
    print("="*50)
    if not PLAYWRIGHT_AVAILABLE:
        print("AVISO: Para usar /fill-form instala:")
        print("  pip install playwright && playwright install chromium")
        print("="*50)
    print("Templates:")
    print("  - PM + Producto Nuevo: template especial con ADO Request Flow")
    print("  - Resto: template estandar con learning path por rol")
    print("="*50)
    print("Presiona Ctrl+C para detener\n")

    app.run(host="localhost", port=5000, debug=False)
