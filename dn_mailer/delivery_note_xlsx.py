"""
dn_mailer/delivery_note_xlsx.py
Точка входа из hooks.py.
Читает настройки из DocType «DN Mailer Settings».
"""

import frappe
from frappe.utils.file_manager import save_file

from dn_mailer.xlsx_builder import build_xlsx


def send_delivery_note_xlsx(doc, method=None):
    """Hook: Delivery Note → on_submit"""
    try:
        settings = frappe.get_single("DN Mailer Settings")

        recipients = _get_recipients(settings)
        if not recipients:
            frappe.logger("dn_mailer").warning(
                f"Нет получателей в DN Mailer Settings — письмо не отправлено "
                f"для {doc.name}"
            )
            return

        company_name = settings.company_name or frappe.defaults.get_global_default("company") or ""
        xlsx_bytes   = build_xlsx(doc, company_name)
        filename     = f"DN_{doc.name}_{doc.posting_date}.xlsx"

        if settings.attach_to_document:
            save_file(
                fname=filename,
                content=xlsx_bytes,
                dt=doc.doctype,
                dn=doc.name,
                is_private=1 if settings.file_is_private else 0,
            )

        _send_email(doc, filename, xlsx_bytes, recipients)

        frappe.logger("dn_mailer").info(
            f"Письмо отправлено → {recipients} | документ: {doc.name}"
        )

    except Exception:
        frappe.logger("dn_mailer").error(
            f"Ошибка при обработке {doc.name}:\n{frappe.get_traceback()}"
        )


def _get_recipients(settings) -> list[str]:
    """
    Для каждой строки таблицы берём email напрямую из поля User.
    fetch_from уже заполняет его в интерфейсе, но на случай
    пустого значения делаем fallback через frappe.db.
    """
    emails = []
    for row in settings.recipients:
        email = row.email or frappe.db.get_value("User", row.user, "email")
        if email:
            emails.append(email)
    return emails


def _send_email(doc, filename: str, xlsx_bytes: bytes, recipients: list[str]):
    frappe.sendmail(
        recipients=recipients,
        subject=f"Накладная {doc.name} от {doc.posting_date}",
        message=(
            f"<p>Добрый день!</p>"
            f"<p>Во вложении — накладная на отгрузку "
            f"<b>{doc.name}</b> для клиента <b>{doc.customer}</b> "
            f"от {doc.posting_date}.</p>"
            f"<p>Позиций в накладной: <b>{len(doc.items)}</b></p>"
        ),
        attachments=[{"fname": filename, "fcontent": xlsx_bytes}],
        reference_doctype=doc.doctype,
        reference_name=doc.name,
    )
