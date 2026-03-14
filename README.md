# 🔄 Acronis Virtual Appliance Device Report Automation

An end-to-end, zero-touch automation that filters the monthly Acronis full device list report and delivers a clean, Virtual Aplliances-only Excel file to SharePoint — entirely within Microsoft 365, at no additional cost.

---

## 🧩 Problem Statement

The organization receives a monthly automated email from Acronis containing a full list of managed devices. However, only devices with **"-VA-"** in their name are operationally relevant.

Previously, this required a team member to:
1. Open the Excel attachment manually
2. Filter by device name
3. Save the filtered file
4. Upload it to SharePoint

This was a **monthly, repetitive, error-prone task** with no business value beyond the filtering itself.

---

## ✅ Solution Overview

A fully automated, cloud-native workflow built entirely on **Microsoft 365** — using **Power Automate** for orchestration and **Office Scripts** for data transformation.

- ⚡ Triggered automatically by the incoming Acronis email
- 🔍 Filters only `-VA-` devices using a TypeScript Office Script
- 📁 Saves the result to SharePoint with a clean monthly filename
- 🧠 Zero manual intervention required

---

## 🛠️ Tech Stack

- **Power Automate** — Workflow orchestration
- **Office Scripts (TypeScript)** — Excel data filtering logic
- **SharePoint** — File storage and auditability
- **Microsoft 365** — No third-party tools or additional cost

---

## ⚙️ Technical Implementation

### Office Script — `filter-va-devices.ts`

- Opens the workbook and reads the used range of the first worksheet
- Dynamically locates the **"Device name"** column by header name
- Iterates rows **bottom-up** to avoid index-shifting bugs when deleting rows
- Removes any row where the device name does not contain `-VA-`
- Throws a descriptive error if the expected column is not found

### Power Automate Flow

The flow consists of three core actions:

| Step | Action | Detail |
|------|--------|--------|
| 1 | **Trigger** | "When a new email arrives (V3)" — watches for emails from `noreply-abc@cloud.acronis.com` with subject containing `"DAILY REPORT ON"`. A Trigger Condition filters only the `Full Device List` attachment using `@startsWith(...)` |
| 2 | **Create file (SharePoint)** | Saves the attachment to a company shared folder with a dynamic monthly filename, for example: `VA Devices Report - March 2026.xlsx` |
| 3 | **Run Script** | Executes `filter-va-devices.ts` against the newly created file using its dynamic ID — no hardcoded paths |

---

## 💼 Business Value

| Benefit | Detail |
|---------|--------|
| ⏱️ Zero manual effort | No human interaction required from email to filtered file |
| ✅ Consistency | Deterministic filter logic — identical results every time |
| 🗂️ Auditability | Files stored in SharePoint with clear monthly naming |
| 💰 No additional cost | Uses existing Microsoft 365 licences only |
| 📈 Scalable | Easily adapted for different identifiers, destinations, or notifications |

---

## 📁 Repository Structure

```
acronis-va-filter-automation/
├── scripts/
│   └── filter-va-devices.ts     # Office Script — filtering logic
├── docs/
│   └── flow-overview.md         # Power Automate flow documentation
└── README.md
```

---

## 🚀 How to Deploy

1. **Office Script** — Open Excel Online, go to Automate > New Script, paste the contents of `filter-va-devices.ts`, and save as `Filter VA Devices`
2. **Power Automate** — Create a new flow using the "When a new email arrives (V3)" trigger, configure the three steps as documented above, and reference the saved Office Script
3. **SharePoint** — Ensure the target folder `/Documents/01. Acronis/VA Devices Reports` exists and the flow has appropriate permissions

---

## 🔮 Future Improvements

- [ ] Email notification on successful execution
- [ ] Error alerting via Teams or Discord webhook if the flow fails
- [ ] Support for filtering by multiple identifiers
- [ ] Automated monthly summary report of VA device count trends

---

## 👤 Author

**Stylianos Vardakas**  
Cyber Security Support Engineer  
[LinkedIn](https://www.linkedin.com/in/stylianos-vardakas/) • [GitHub](https://github.com/steliosvardakas)

---

## 📄 License

MIT
