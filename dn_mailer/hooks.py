app_name        = "dn_mailer"
app_title       = "DN Mailer"
app_publisher   = "Your Company"
app_description = "Sends a formatted XLSX file by email when a Delivery Note is submitted"
app_version     = "0.2.0"
app_email       = "you@company.com"
app_license     = "MIT"

doc_events = {
    "Delivery Note": {
        "on_submit": "dn_mailer.delivery_note_xlsx.send_delivery_note_xlsx",
    }
}
