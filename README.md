# DraftMate v3

DraftMate is a powerful desktop application designed to automate personal email outreach. It allows you to bulk-generate draft emails in Microsoft Outlook using data from CSV files or Google Sheets, customized with dynamic templates.

![DraftMate UI](https://via.placeholder.com/800x450.png?text=DraftMate+Interface+Preview)

## 🚀 Key Features

*   **Bulk Draft Generation**: Automatically create dozens or hundreds of email drafts in Microsoft Outlook with a single click.
*   **Smart Templating**: Use dynamic placeholders (e.g., `{{FirstName}}`, `{{Company}}`) that automatically populate from your data source.
*   **Multiple Data Sources**: Support for local CSV files and Google Sheets.
*   **Template Overrides**: Assign specific templates to specific recipients when a one-size-fits-all approach isn't enough.
*   **Template Export**: Export your curated templates to a ZIP file for backup or sharing.
*   **Preview Mode**: Safely preview subject lines and email bodies before generating them to ensure everything looks perfect.
*   **Resume/Attachment Support**: Automatically attach files (like your resume) to every generated draft.
*   **Safety First**: Includes a "Dry Run" mode to validate the process without cluttering your drafts folder.
*   **Modern UI**: Built with a sleek, dark-mode enabled interface for a premium user experience.

## 🛠 Tech Stack

*   **Frontend**: React, TypeScript, Vite, TailwindCSS (via custom styles)
*   **Backend**: Rust (Tauri v2)
*   **Engine**: Python (Data processing & AppleScript bridge)
*   **Integration**: AppleScript (for Microsoft Outlook control)

## 📋 Prerequisites

*   **Operating System**: macOS (Required for AppleScript integration).
*   **Email Client**: Microsoft Outlook for Mac.
    *   *Note: You must use "Classic Outlook". The "New Outlook" for Mac does not yet support the AppleScript features required for automation.*
*   **Runtime**: 
    *   Node.js (for building the UI)
    *   Rust (for compiling the application)
    *   Python 3 (for the processing engine)

## 📦 Installation & Setup

1.  **Clone the repository**
    ```bash
    git clone https://github.com/yourusername/draftmate-v3.git
    cd draftmate-v3
    ```

2.  **Install Frontend Dependencies**
    ```bash
    cd tauri-ui
    npm install
    ```

3.  **Setup Python Engine**
    Ensure you have Python installed. It's recommended to create a virtual environment:
    ```bash
    cd ../
    python3 -m venv .venv
    source .venv/bin/activate
    pip install -r requirements.txt
    ```

4.  **Run in Development Mode**
    ```bash
    cd tauri-ui
    npm run tauri dev
    ```

## 📖 Usage Guide

### 1. Prepare Your Data
Create a CSV file or Google Sheet with your recipient data. Ensure you have a header row with columns like `FirstName`, `Email`, `Company`, etc.
*   **Required Column**: You must have a column containing email addresses (e.g., `Email` or `Contact Email`).
*   **Optional**: Add a `Generate` column (values: `1`, `true`, `yes`) to control which rows get processed.

### 2. Load Data
*   Open DraftMate.
*   Click **"Import Data"** and select your CSV file or paste your Google Sheet URL.

### 3. Create Templates
*   Go to the **Templates** tab.
*   Create a new template. Use placeholders corresponding to your CSV headers inside double curly braces.
    *   Example: *"Hi {{FirstName}}, I saw your work at {{Company}}..."*

### 4. Preview & Validate
*   Use the **Preview** tab to see exactly how your emails will look.
*   Check that all placeholders are resolving correctly.

### 5. Generate Drafts
*   Enter your **Subject Line** (supports placeholders too!).
*   Optionally attach a file (e.g., PDF Resume).
*   Click **Generate**. DraftMate will open Outlook and create a draft for each recipient, leaving them open for you to review and hit "Send".

## ⚠️ Troubleshooting

**"Outlook AppleScript not available"**
*   Please ensure you are using **Classic Outlook**.
*   In Outlook, go to the `Help` menu or look for a toggle switch in the top right corner provided by Microsoft to "Revert to Legacy Outlook" or turn off "New Outlook".

**Placeholders not working**
*   Ensure the placeholder name matches your CSV header exactly (case-insensitive).
*   Example: `{{firstname}}` in template matches `FirstName` in CSV.

## 📄 License

DraftMate is proprietary software. A valid license key is required to activate the full functionality of the application.
