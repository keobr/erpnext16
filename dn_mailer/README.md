# dn_mailer

Custom App для ERPNext 16.  
При **Submit** документа **Delivery Note** автоматически:

1. Формирует `.xlsx` с шапкой, мета-данными и таблицей товаров
2. Прикрепляет файл к документу (вкладка Attachments)
3. Отправляет письмо с вложением всем получателям из настроек

---

## Структура

```
dn_mailer/
├── setup.py
├── pyproject.toml
└── dn_mailer/
    ├── __init__.py
    ├── hooks.py
    ├── modules.txt
    ├── delivery_note_xlsx.py       # точка входа из hooks
    ├── xlsx_builder.py             # построение XLSX (без frappe)
    └── doctype/
        ├── dn_mailer_settings/     # Single DocType с настройками
        │   ├── dn_mailer_settings.json
        │   └── dn_mailer_settings.py
        └── dn_mailer_recipient/    # дочерняя таблица получателей
            ├── dn_mailer_recipient.json
            └── dn_mailer_recipient.py
```

---

## Настройка после установки

Откройте **DN Mailer Settings** (поиск в строке навигации ERPNext):

| Поле                        | Описание                                          |
|-----------------------------|---------------------------------------------------|
| Company Name                | Название компании в шапке xlsx                    |
| Attach XLSX to Delivery Note| Прикреплять файл к документу (вкладка Attachments)|
| Save Attachment as Private  | Файл виден только залогиненным пользователям      |
| Recipients (таблица)        | Выбираете пользователей — email подтягивается сам |

В таблице Recipients можно добавить любое количество пользователей.  
Email берётся из профиля пользователя ERPNext автоматически.

---

## Установка

```bash
# 1. Скопируйте папку в apps/
cp -r dn_mailer /home/frappe/frappe-bench/apps/

# 2. Установите на сайт
cd /home/frappe/frappe-bench
bench --site your-site.local install-app dn_mailer

# 3. Примените миграцию (создаст DocType-ы в БД)
bench --site your-site.local migrate

# 4. Перезапустите
bench restart
```

---

## Проверка

```bash
# Убедитесь что hook зарегистрирован
bench --site your-site.local console
>>> frappe.get_hooks("doc_events")

# Логи с меткой dn_mailer
tail -f logs/frappe.log | grep dn_mailer
```

---

## Требования

- ERPNext 16 / Frappe 16
- `openpyxl` (входит в стандартный frappe-bench)
