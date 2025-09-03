

from __future__ import unicode_literals
import os
import re
import fnmatch
import zipfile
import xml.etree.ElementTree as ET

import frappe
from frappe.model.document import Document
from erpnext.accounts.party import set_taxes as party_st

try:
    from erpnext.stock.get_item_details import get_item_details
except Exception:
    get_item_details = None


# --------------------------- Helpers ---------------------------


def _report_kind_from_root(root):
    if root is None:
        return "UNKNOWN"
    if root.tag == "MesReport":
        return "MES"
    if root.tag == "MxdrMspReport":
        return "MXDR"
    if root.find("ReportEntry") is not None:
        return "MES-DETAIL"
    return root.tag  # fallback

def _sales_invoice_has_field(fieldname):
    try:
        return frappe.get_meta("Sales Invoice").has_field(fieldname)
    except Exception:
        return False

def _get_intro_base_text(settings, report_kind):
    """MXDR bevorzugt introduction_text_mxdr, sonst introduction_text; MES nutzt introduction_text."""
    if report_kind == "MXDR":
        return (getattr(settings, "introduction_text_mxdr", None)
                or getattr(settings, "introduction_text", "")
                or "")
    return getattr(settings, "introduction_text", "") or ""

def _set_intro_field(sinv, fieldname, html):
    """In gewünschtes Intro-Feld schreiben; Fallback auf 'introduction_text'; Notfall → remarks (plain)."""
    target = fieldname if _sales_invoice_has_field(fieldname) else "introduction_text"
    try:
        setattr(sinv, target, html)
    except Exception:
        try:
            sinv.remarks = (sinv.remarks or "") + ("\n" if sinv.remarks else "") + frappe.utils.strip_html(html)
        except Exception:
            pass


def _safe_int(val, default=0):
    try:
        return int(val)
    except Exception:
        return default


def _get_default_selling_price_list(customer_doc):
    """Kunde → Selling Settings → 'Standard Selling' (Fallback)."""
    pl = getattr(customer_doc, "default_price_list", None) or None
    if not pl:
        try:
            pl = frappe.db.get_single_value("Selling Settings", "selling_price_list")
        except Exception:
            pl = None
    if not pl:
        try:
            if frappe.db.exists("Price List", {"price_list_name": "Standard Selling"}):
                pl = "Standard Selling"
        except Exception:
            pass
    return pl


def _apply_item_price(row, sinv, item_doc, qty, log_messages):
    """Preis setzen:
    1) get_item_details (falls vorhanden)
    2) Fallback relaxed: erst Item Price in Price List (ohne selling=1), sonst irgendein Item Price zum Item
    """
    # 1) Standardweg
    if get_item_details:
        try:
            args = frappe._dict({
                "doctype": "Sales Invoice",
                "item_code": item_doc.name,
                "company": sinv.company,
                "customer": sinv.customer,
                "price_list": sinv.selling_price_list,
                "transaction_date": sinv.posting_date,
                "qty": qty,
                "uom": getattr(item_doc, "sales_uom", None) or getattr(item_doc, "stock_uom", None),
                "currency": sinv.currency,
                "conversion_rate": 1,
            })
            details = get_item_details(args)
            for k in ("uom", "conversion_factor", "price_list_rate", "rate",
                      "discount_percentage", "discount_amount",
                      "base_price_list_rate", "base_rate"):
                if k in details and details[k] is not None:
                    row.set(k, details[k])
            if row.get("rate"):
                return
        except Exception:
            pass

    # 2) Fallback relaxed
    price_doc = None
    try:
        price_doc = frappe.db.get_value(
            "Item Price",
            {"item_code": item_doc.name, "price_list": sinv.selling_price_list},
            ["price_list_rate", "currency", "price_list"],
            as_dict=True
        )
    except Exception:
        price_doc = None

    if not price_doc:
        try:
            price_doc = frappe.db.get_value(
                "Item Price",
                {"item_code": item_doc.name},
                ["price_list_rate", "currency", "price_list"],
                as_dict=True
            )
            if price_doc:
                log_messages.append(
                    f"Hinweis: Preis für Item '{item_doc.name}' aus Price List '{price_doc.get('price_list')}' übernommen (Fallback)."
                )
        except Exception:
            price_doc = None

    if price_doc and price_doc.get("price_list_rate") is not None:
        row.price_list_rate = price_doc["price_list_rate"]
        row.rate = price_doc["price_list_rate"]
        try:
            if not getattr(sinv, "currency", None) and price_doc.get("currency"):
                sinv.currency = price_doc["currency"]
        except Exception:
            pass
    else:
        log_messages.append(
            f"Hinweis: Kein Item Price für Item '{item_doc.name}' gefunden (Price List: {sinv.selling_price_list or '—'})."
        )


def _infer_month_year(zip_path, xml_name, root):
    """Ermittle Monat/Jahr aus MesReport/ReportEntry oder Dateinamen."""
    # 1) MesReport
    try:
        if root is not None and root.tag == "MesReport":
            m = _safe_int(root.attrib.get("Month"))
            y = _safe_int(root.attrib.get("Year"))
            if 1 <= m <= 12 and y > 0:
                return m, y, "MesReport attributes"
    except Exception:
        pass

    # 2) ReportEntry
    try:
        rep = root.find("ReportEntry")
        if rep is not None:
            m = _safe_int(rep.attrib.get("Month"))
            y = _safe_int(rep.attrib.get("Year"))
            if 1 <= m <= 12 and y > 0:
                return m, y, "ReportEntry attributes"
    except Exception:
        pass

    # 3) Dateinamen scannen
    def scan_name(name):
        if not name:
            return None
        s = os.path.basename(name).lower()
        patterns = [
            r'(?P<m>\d{1,2})[_\-.](?P<y>\d{4})',
            r'(?P<y>\d{4})[_\-.](?P<m>\d{1,2})',
        ]
        for pat in patterns:
            mm = re.search(pat, s)
            if mm:
                m = _safe_int(mm.group('m'))
                y = _safe_int(mm.group('y'))
                if 1 <= m <= 12 and y >= 2000:
                    return m, y
        return None

    for candidate, label in ((zip_path, "zip filename"), (xml_name, "xml filename")):
        res = scan_name(candidate)
        if res:
            return res[0], res[1], label

    return None, None, None

def _prepare_invoice_common(customer_doc, settings, month, year, prefix, marker_label, marker_value, report_kind, intro_field):
    """Baut Grundstruktur der Sales Invoice (ohne insert) inkl. richtigem Intro-Feld."""
    posting_dt = frappe.utils.today()

    base_txt = _get_intro_base_text(settings, report_kind)
    intro_html = (base_txt or "") + (
        f"<div><br></div><div>Leistungszeitraum {month}.{year}<br>"
        f"{marker_label}: {marker_value}</div>"
    )

    sinv = frappe.get_doc({
        "doctype": "Sales Invoice",
        "title": f"{prefix} {month}.{year} {customer_doc.customer_name}",
        "customer": customer_doc.name,
        "status": "Draft",
        "tc_name": settings.terms_and_conditions,
        "company": frappe.get_doc("Global Defaults").default_company,
        "posting_date": posting_dt,
        "set_posting_time": 1,
    })

    # Intro in das gewünschte Feld (mit Fallback) schreiben
    _set_intro_field(sinv, intro_field, intro_html)

    # Price List / Currency
    sinv.selling_price_list = _get_default_selling_price_list(customer_doc)
    if sinv.selling_price_list:
        try:
            sinv.price_list_currency = frappe.db.get_value("Price List", sinv.selling_price_list, "currency")
        except Exception:
            sinv.price_list_currency = None
    try:
        company_currency = frappe.db.get_value("Company", sinv.company, "default_currency")
        sinv.currency = sinv.price_list_currency or company_currency
        sinv.conversion_rate = 1
        sinv.plc_conversion_rate = 1
    except Exception:
        pass

    # Payment Terms
    sinv.payment_terms_template = (customer_doc.payment_terms or settings.payment_terms_template)

    return sinv






def _finalize_terms_and_totals(sinv):
    """Fälligkeit nach posting_date neu berechnen + Summen/Steuern kalkulieren."""
    try:
        if hasattr(sinv, "due_date"):
            sinv.due_date = None
        if getattr(sinv, "payment_schedule", None):
            sinv.set("payment_schedule", [])

        if hasattr(sinv, "set_missing_values"):
            sinv.set_missing_values()
        if hasattr(sinv, "set_payment_schedule"):
            sinv.set_payment_schedule()
        if hasattr(sinv, "set_due_date"):
            sinv.set_due_date()
    except Exception:
        pass

    # Fallback-Fälligkeit 7/14 Tage
    try:
        payment_schedule_empty = (not getattr(sinv, "payment_schedule", None)) or len(sinv.payment_schedule) == 0
        no_due_date_field = (not hasattr(sinv, "due_date")) or (hasattr(sinv, "due_date") and not sinv.due_date)
        if payment_schedule_empty and no_due_date_field:
            tmpl_name = (sinv.payment_terms_template or "").lower()
            delta = 7 if "7" in tmpl_name else 14
            due = frappe.utils.add_days(sinv.posting_date, delta)
            if hasattr(sinv, "due_date"):
                sinv.due_date = due
            else:
                sinv.append("payment_schedule", {
                    "due_date": due,
                    "invoice_portion": 100,
                })
    except Exception:
        pass

    try:
        if hasattr(sinv, "calculate_taxes_and_totals"):
            sinv.calculate_taxes_and_totals()
    except Exception:
        pass


# --------------------------- Doctype ---------------------------

class Abrechungen(Document):
    @frappe.whitelist()
    def start_processing_zip(self):
        # 1) ZIP finden
        attached_files = frappe.get_all(
            'File',
            filters={"attached_to_doctype": "Abrechungen", "attached_to_name": self.name},
            fields=["name"],
            limit=1,
        )
        if not attached_files:
            frappe.throw("Keine Datei an diesem Dokument gefunden.")
        zip_path = frappe.utils.file_manager.get_file_path(attached_files[0]["name"])

        # 2) XMLs aus ZIP
        xml_files = self.extract_xml_from_zip(zip_path)
        if not xml_files:
            frappe.throw("Keine XML-Datei in der ZIP gefunden.")

        # 3) Auswahl: short -> detailed -> erste
        pick_short = next((x for x in xml_files if 'short' in x['name'].lower()), None)
        pick_detailed = next((x for x in xml_files if 'detailed' in x['name'].lower()), None)
        chosen = pick_short or pick_detailed or xml_files[0]

        self.xml_data = chosen['content'].decode('utf-8', errors='replace')

        # 4) XML parsen
        try:
            root = ET.fromstring(self.xml_data)
        except ET.ParseError:
            frappe.throw("XML konnte nicht geparst werden. Bitte die Datei prüfen.")

        # 5) Monat/Jahr robust bestimmen
        invoice_month, invoice_year, src = _infer_month_year(zip_path, chosen['name'], root)
        if not (invoice_month and invoice_year):
            frappe.throw('Abrechnungsmonat/-jahr konnte nicht ermittelt werden.')
        rep_kind = _report_kind_from_root(root)

        # Wenn du auch den XML-Dateinamen in den Schlüssel aufnehmen willst (um Re-Runs unterschiedlicher Dateien im selben Monat zu erlauben),
        # setze diese Variable auf True:
        USE_FILENAME_IN_KEY = False

        if USE_FILENAME_IN_KEY:
            self.month = f"{invoice_month}.{invoice_year} ({rep_kind}) | {os.path.basename(chosen['name'])}"
        else:
            self.month = f"{invoice_month}.{invoice_year} ({rep_kind})"


        # 6) Doppel-Abrechnungen verhindern (gleicher Monat)
        if not self.check_report_date():
            self.save(ignore_permissions=True)
            return

        # 7) Log
        log = [
            f"Quelle ZIP: {os.path.basename(zip_path)}",
            f"Verwendete XML: {chosen['name']}",
            f"Root: {root.tag} / Monat: {invoice_month} / Jahr: {invoice_year} (Quelle: {src})",
        ]

        # 8) Routing nach Root
        if root.tag == 'MesReport':
            summary = self._process_mes_report(root, invoice_month, invoice_year, log)
        elif root.tag == 'MxdrMspReport':
            summary = self._process_mxdr_report(root, invoice_month, invoice_year, log)
        else:
            # Detailed/Legacy
            summary = self._process_legacy_report(root, invoice_month, invoice_year, log)

        # 9) Abschluss
        log.append(summary)
        self.log = "\n".join(log)
        has_created = ("created=" in summary and int(summary.split("created=")[1].split(",")[0]) > 0)
        self.status = "Ausgangsrechnungen erstellt" if has_created else "Keine Rechnungen erstellt"
        m = re.search(r"Gesamtanzahl Clients gezählt: (\d+)", self.log)
        self.anzahl_clients = m.group(1) if m else "0"
        self.save(ignore_permissions=True)

    # ---------- MES (Short) ----------
    def _process_mes_report(self, root, invoice_month, invoice_year, log):
        total_clients = 0
        cnt_created = cnt_dup = cnt_not_found = cnt_error = 0

        for mng in root.findall('ManagementServer'):
            raw_id = (mng.attrib.get('Id') or mng.attrib.get('id') or '').strip()
            mac = _safe_int(mng.attrib.get('MaxActiveClients'), 0)

            if mac <= 0:
                msgtext = f"Übersprungen (0 Clients): {raw_id}"
                log.append(msgtext)
                frappe.msgprint(msgtext)
                continue

            try:
                result = self.create_mes_invoice(raw_id, mac, invoice_month, invoice_year)
                if result == "created":
                    log.append(f"OK: Rechnung erstellt → {raw_id} Qty={mac}")
                    total_clients += mac
                    cnt_created += 1
                elif result == "duplicate":
                    log.append(f"Übersprungen (duplicate): {raw_id}")
                    cnt_dup += 1
                elif result and str(result).startswith("not_found"):
                    log.append(f"Übersprungen (not_found): {raw_id}")
                    cnt_not_found += 1
                else:
                    log.append(f"Übersprungen (unbekannt): {raw_id} → {result}")
            except Exception as e:
                frappe.log_error(frappe.get_traceback(), f"MES Invoice fehlgeschlagen für {raw_id}")
                log.append(f"FEHLER: {raw_id} → {e}")
                cnt_error += 1

        log.append(f"Gesamtanzahl Clients gezählt: {total_clients}")
        return f"Summary: created={cnt_created}, duplicate={cnt_dup}, not_found={cnt_not_found}, error={cnt_error}"

    # ---------- MXDR/MSP (Short) ----------
    def _process_mxdr_report(self, root, invoice_month, invoice_year, log):
        total_clients = 0
        cnt_created = cnt_dup = cnt_not_found = cnt_error = 0

        for lic in root.findall('License'):
            license_key = (lic.attrib.get('LicenseKey') or '').strip()
            qty = _safe_int(lic.attrib.get('ActiveClients'), 0)

            if qty <= 0:
                msgtext = f"Übersprungen (0 Clients): {license_key or '(ohne Schlüssel)'}"
                log.append(msgtext)
                frappe.msgprint(msgtext)
                continue

            try:
                result = self.create_mxdr_invoice(license_key, qty, invoice_month, invoice_year)
                if result == "created":
                    log.append(f"OK: Rechnung erstellt → {license_key} Qty={qty}")
                    total_clients += qty
                    cnt_created += 1
                elif result == "duplicate":
                    log.append(f"Übersprungen (duplicate): {license_key}")
                    cnt_dup += 1
                elif result and str(result).startswith("not_found"):
                    log.append(f"Übersprungen (not_found): {license_key}")
                    cnt_not_found += 1
                else:
                    log.append(f"Übersprungen (unbekannt): {license_key} → {result}")
            except Exception as e:
                frappe.log_error(frappe.get_traceback(), f"MXDR Invoice fehlgeschlagen für {license_key}")
                log.append(f"FEHLER: {license_key} → {e}")
                cnt_error += 1

        log.append(f"Gesamtanzahl Clients gezählt: {total_clients}")
        return f"Summary: created={cnt_created}, duplicate={cnt_dup}, not_found={cnt_not_found}, error={cnt_error}"

    # ---------- Detailed/Legacy ----------
    def _process_legacy_report(self, root, invoice_month, invoice_year, log):
        reportentries = root.findall('ReportEntry')
        if not reportentries:
            if root.tag == 'MxdrMspReport':
                frappe.throw("Die Datei ist ein MXDR/MSP-Report (MxdrMspReport). Diese Variante ist unterstützt – bitte Short-Report mit <License>-Knoten liefern.")
            frappe.throw("In der XML fehlen ReportEntry-Knoten. Bitte eine Short- (MesReport) oder Detailed-Datei liefern.")

        if len(reportentries) != 1:
            frappe.msgprint('Keinen oder mehr als ein ReportEntry im XML-Code gefunden, breche ab.')
            return "Summary: created=0, duplicate=0, not_found=0, error=0"

        rep = reportentries[0]
        log += [
            f"Report für Firma: {rep.attrib.get('Company', '')}",
            f"G Data Kundennummer: {rep.attrib.get('GDCustomerNr', '')}",
            f"Login Name: {rep.attrib.get('Login', '')}",
            f"Produkt: {rep.attrib.get('Product', '')}",
            f"Gesamt (Report) MaxActiveClients: {rep.attrib.get('MaxActiveClients', '0')}",
        ]

        total_clients = 0
        cnt_created = cnt_dup = cnt_not_found = cnt_error = 0

        for mng in rep.findall('ManagementServer'):
            raw_id = (mng.attrib.get('id') or mng.attrib.get('Id') or '').strip()
            mac = _safe_int(mng.attrib.get('MaxActiveClients'), 0)

            if mac <= 0:
                msgtext = f"Übersprungen (0 Clients): {raw_id}"
                log.append(msgtext)
                frappe.msgprint(msgtext)
                continue

            try:
                result = self.create_mes_invoice(raw_id, mac, invoice_month, invoice_year)
                if result == "created":
                    log.append(f"OK: Rechnung erstellt → {raw_id} Qty={mac}")
                    total_clients += mac
                    cnt_created += 1
                elif result == "duplicate":
                    log.append(f"Übersprungen (duplicate): {raw_id}")
                    cnt_dup += 1
                elif result and str(result).startswith("not_found"):
                    log.append(f"Übersprungen (not_found): {raw_id}")
                    cnt_not_found += 1
                else:
                    log.append(f"Übersprungen (unbekannt): {raw_id} → {result}")
            except Exception as e:
                frappe.log_error(frappe.get_traceback(), f"MES Invoice fehlgeschlagen für {raw_id}")
                log.append(f"FEHLER: {raw_id} → {e}")
                cnt_error += 1

        log.append(f"Gesamtanzahl Clients gezählt: {total_clients}")
        return f"Summary: created={cnt_created}, duplicate={cnt_dup}, not_found={cnt_not_found}, error={cnt_error}"

    # --------------------------- Create Invoices ---------------------------

    def create_mes_invoice(self, mes_id, max_active_clients, invoice_month, invoice_year):
        """MES: Abrechnung je Management-Server-ID (inkl. '#'-Suffix)."""
        if _safe_int(max_active_clients) <= 0:
            return "zero_clients"

        settings = frappe.get_cached_doc("GDATAMES Settings")

        # EXAKTER Lookup
        matches = frappe.get_all('Management Server', filters={"management_server_id": mes_id}, fields=["name"])
        if not matches:
            trim_row = frappe.db.sql(
            """
            SELECT name FROM `tabManagement Server`
            WHERE UPPER(TRIM(management_server_id)) = UPPER(TRIM(%s))
            LIMIT 1
            """,
            (mes_id,),
            as_dict=True
            )
            if trim_row:
                matches = trim_row

        if not matches:
            frappe.msgprint(f"Management Server ID {mes_id} nicht gefunden.")
            return f"not_found: {mes_id}"

        ms_name = matches[0]["name"] if isinstance(matches[0], dict) else matches[0].name
        doc_mserver = frappe.get_doc("Management Server", ms_name)
        product = frappe.get_doc("Produkte", doc_mserver.product)
        item = frappe.get_doc("Item", product.item)
        customer_doc = frappe.get_doc("Customer", doc_mserver.customer)

        # Duplikate prüfen
        if self._find_existing_invoice(customer_doc, "MES", invoice_month, invoice_year, "Management Server ID", mes_id):
            frappe.msgprint("Rechnung existiert bereits. Überspringe.")
            return "duplicate"

        # Grundgerüst (Intro → introduction_text)
        sinv = _prepare_invoice_common(
        customer_doc, settings, invoice_month, invoice_year,
        prefix="MES", marker_label="Management Server ID", marker_value=mes_id,
        report_kind="MES", intro_field="introduction_text"
        )

        # Position + Preis
        row = sinv.append("items", {"doctype": "Sales Invoice Item", "item_code": item.name, "qty": int(max_active_clients)})
        log_msgs = []
        _apply_item_price(row, sinv, item, int(max_active_clients), log_msgs)

        # Steuern
        taxes_template = None
        try:
            taxes_template = party_st(sinv.customer, "Customer", sinv.posting_date, sinv.company)
        except TypeError:
            taxes_template = party_st(sinv.customer, "Customer")
        except Exception:
            taxes_template = None

        if taxes_template:
            sinv.taxes_and_charges = taxes_template
            try:
                tmpl = frappe.get_doc("Sales Taxes and Charges Template", taxes_template)
                for tx in tmpl.taxes:
                    sinv.append("taxes", {
                    "doctype": "Sales Taxes and Charges",
                    "charge_type": tx.charge_type,
                    "account_head": tx.account_head,
                    "rate": tx.rate,
                    "description": tx.description,
                   })
            except Exception:
                frappe.log_error(frappe.get_traceback(), f"Konnte Sales Taxes and Charges Template {taxes_template} nicht laden")

        _finalize_terms_and_totals(sinv)
        sinv.insert()
        for m in log_msgs:
            frappe.msgprint(m)
        return "created"


    def create_mxdr_invoice(self, license_key, active_clients, invoice_month, invoice_year):
        """MXDR/MSP: Abrechnung je Lizenzschlüssel – über Doctype 'Management Server' (ID = LicenseKey)."""
        if _safe_int(active_clients) <= 0:
            return "zero_clients"

        settings = frappe.get_cached_doc("GDATAMES Settings")

        # Lookup im Doctype "Management Server" (ID = LicenseKey)
        matches = frappe.get_all('Management Server', filters={"management_server_id": license_key}, fields=["name"])
        if not matches:
            trim_row = frappe.db.sql(
            """
            SELECT name FROM `tabManagement Server`
            WHERE UPPER(TRIM(management_server_id)) = UPPER(TRIM(%s))
            LIMIT 1
            """,
            (license_key,),
            as_dict=True
            )
            if trim_row:
                matches = trim_row

        if not matches:
            frappe.msgprint(f"MXDR Lizenzschlüssel {license_key} nicht gefunden (als Management Server ID).")
            return f"not_found: {license_key}"

        ms_name = matches[0]["name"] if isinstance(matches[0], dict) else matches[0].name
        doc_mserver = frappe.get_doc("Management Server", ms_name)
        product = frappe.get_doc("Produkte", doc_mserver.product)
        item = frappe.get_doc("Item", product.item)
        customer_doc = frappe.get_doc("Customer", doc_mserver.customer)

        # Duplikatprüfung (Titel + Marker)
        if self._find_existing_invoice(customer_doc, "MXDR", invoice_month, invoice_year, "MXDR Lizenzschlüssel", license_key):
            frappe.msgprint("Rechnung existiert bereits. Überspringe.")
            return "duplicate"

        # Intro-Feld bevorzugt 'introduction_text_mxdr' (mit Fallback)
        intro_field = "introduction_text_mxdr" if _sales_invoice_has_field("introduction_text_mxdr") else "introduction_text"

        # Grundgerüst
        sinv = _prepare_invoice_common(
            customer_doc, settings, invoice_month, invoice_year,
            prefix="MXDR", marker_label="MXDR Lizenzschlüssel", marker_value=license_key,
            report_kind="MXDR", intro_field=intro_field
            )

        # Position + Preis
        row = sinv.append("items", {"doctype": "Sales Invoice Item", "item_code": item.name, "qty": int(active_clients)})
        log_msgs = []
        _apply_item_price(row, sinv, item, int(active_clients), log_msgs)

        # Steuern
        taxes_template = None
        try:
            taxes_template = party_st(sinv.customer, "Customer", sinv.posting_date, sinv.company)
        except TypeError:
            taxes_template = party_st(sinv.customer, "Customer")
        except Exception:
            taxes_template = None

        if taxes_template:
            sinv.taxes_and_charges = taxes_template
            try:
                tmpl = frappe.get_doc("Sales Taxes and Charges Template", taxes_template)
                for tx in tmpl.taxes:
                    sinv.append("taxes", {
                    "doctype": "Sales Taxes and Charges",
                    "charge_type": tx.charge_type,
                    "account_head": tx.account_head,
                    "rate": tx.rate,
                    "description": tx.description,
                })
            except Exception:
                frappe.log_error(frappe.get_traceback(), f"Konnte Sales Taxes and Charges Template {taxes_template} nicht laden")

        _finalize_terms_and_totals(sinv)
        sinv.insert()
        for m in log_msgs:
            frappe.msgprint(m)
        return "created"


    # --------------------------- Utilities ---------------------------

    def _find_existing_invoice(self, customer_doc, prefix, month, year, marker_label, marker_value):
        """Sucht existierende Rechnung:
        - gleicher Kunde
        - Titel: '<prefix> M.Y Kunde'
        - Markertext (z. B. 'Management Server ID: <id>' oder 'MXDR Lizenzschlüssel: <key>')
        """
        title = f"{prefix} {month}.{year} {customer_doc.customer_name}"
        like_marker = f"%{marker_label}: {marker_value}%"
        rows = frappe.db.sql(
            """
            SELECT name FROM `tabSales Invoice`
            WHERE customer=%s AND title=%s AND docstatus < 2
              AND (introduction_text LIKE %s OR remarks LIKE %s)
            LIMIT 1
            """,
            (customer_doc.name, title, like_marker, like_marker),
        )
        return rows[0][0] if rows else None

    def extract_xml_from_zip(self, inputzip):
        """Liste von {name, content(bytes)} für alle *.xml in der ZIP."""
        results = []
        with zipfile.ZipFile(inputzip, 'r') as zf:
            for name in zf.namelist():
                if fnmatch.fnmatch(name.lower(), '*.xml'):
                    results.append({'name': name, 'content': zf.read(name)})
        return results

    def check_report_date(self):
        """True = Abrechnung für self.month existiert noch nicht."""
        existing = frappe.get_all("Abrechungen", filters={"month": self.month}, fields=["name"], limit=1)
        if existing:
            self.status = "fehlerhaft"
            self.log = (
                f"Zu dem Abrechnungsmonat {self.month} existiert bereits eine Abrechnung "
                f"({existing[0]['name']}), es wurde keine neue Abrechnung erstellt."
            )
            return False
        return True
