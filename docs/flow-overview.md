# Power Automate Flow — Documentation

## Flow Name
`Acronis VA Device Report Filter`

## Trigger
**When a new email arrives (V3)**

| Setting | Value |
|---------|-------|
| From | `noreply-abc@cloud.acronis.com` |
| Subject filter | `DAILY REPORT ON` |
| Trigger Condition | `@startsWith(triggerOutputs()?['body/attachments'][0]['name'], 'Full Device List')` |

The trigger condition ensures the flow only activates for the specific Acronis full device list email, silently ignoring all other daily reports from the same sender.

---

## Step 1 — Create File (SharePoint)

Extracts the attachment from the email trigger and saves it to SharePoint.

| Setting | Value |
|---------|-------|
| Site Address | `SharePoint Site Address` |
| Folder Path | `Path to shared folder` |
| File Name | `concat('VA Devices Report - ', formatDateTime(utcNow(), 'MMMM yyyy'), '.xlsx')` |
| File Content | `triggerOutputs()?['body/attachments'][0]['contentBytes']` |

The dynamic filename produces clean monthly outputs like `VA Devices Report - March 2026.xlsx`, updating automatically each month without any configuration changes.

---

## Step 2 — Run Script (Excel Online Business)

Executes the Office Script against the newly created SharePoint file.

| Setting | Value |
|---------|-------|
| Location | SharePoint site |
| Document Library | Documents |
| File | *(dynamic ID from Step 1 output)* |
| Script | `Filter VA Devices` |

Using the dynamic file ID from the previous step ensures the script always operates on the correct file — no hardcoded paths required.

---

## Flow Diagram

```
[Email arrives from Acronis]
        │
        ▼
[Trigger Condition Check]
  Does attachment name
  start with "Full Device List"?
        │
     YES│
        ▼
[Save attachment to SharePoint]
  /Documents/01. Acronis/VA Devices Reports
  VA Devices Report - {Month Year}.xlsx
        │
        ▼
[Run Office Script]
  Filter VA Devices
  → Removes all non-VA rows
        │
        ▼
[Done — filtered file ready in SharePoint]
```

---

## Error Handling

If the "Device name" column is not found in the report, the Office Script throws a descriptive error which surfaces in the Power Automate run history, making it easy to diagnose format changes in future Acronis report versions.
