# COD Automation Project using Google Apps Script

This project automates daily **Cash on Delivery (COD) liability processing and vendor communication** using Google Apps Script. It was built to solve a real-world operations problem involving 5000+ rows of daily rider and COD data.

## 🔧 What It Does

- 📥 Processes raw rider delivery data (COD)
- 🔍 Maps rider IDs to vendor and brand info
- 🚦 Highlights unmapped or error rows
- 📊 Summarizes vendor-wise outstanding CODs
- 📤 Sends automated vendor emails with summary + raw data as Excel attachments
- 🧹 Deletes temporary files/sheets after execution

## 💡 Why I Built This

As an operations professional with 9+ years of experience, I wanted to **stop wasting hours on repetitive Excel work**. I collaborated with ChatGPT to convert my manual logic into a **fully automated script**, reducing errors and time spent.

> I didn’t know how to code — but I learned how to convert my thinking into automation.

## 📁 Files

| File Name                  | Purpose                                      |
|---------------------------|----------------------------------------------|
| `Step1_ProcessCODData.gs` | Reads and processes data into structured format |
| `Step2_EmailVendors.gs`   | Sends emails with attachments to each vendor  |

## ⚙️ Tech Used

- Google Sheets
- Google Apps Script (JavaScript-based)
- Gmail API
- Google Drive API

## 📈 Impact

- ⏱️ Reduced manual work by 90%
- ✅ Eliminated data errors due to copy-paste
- 💬 Enabled daily vendor reporting without human effort

## 🔗 Showcase Link

Feel free to clone or reuse for similar operational automations!

---

**Made with logic, no-code grit, and the help of ChatGPT.**
