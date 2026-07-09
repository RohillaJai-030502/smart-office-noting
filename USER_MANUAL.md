# User Manual: Smart Noting Generator (AeroForm Automatrix)

Welcome to the **Smart Noting Generator**! This application is designed to automate the creation of official noting sheets, faxes, and emails for the TBRL Training and Internship divisions. 

By replacing manual word processor formatting with simple web forms and pre-formatted templates, it guarantees that all generated documents comply with official guidelines, margins, and page limits automatically.

---

## 📖 Part 1: Glossary of Dashboard Buttons

Use this section to look up what any button on the screen does and when to click it.

### Main Navigation Bar
* **`Dashboard`** (or **`← Dashboard`**): Returns you to the main home screen from any sub-page.
* **`Training Module`**: Opens the submenu containing noting sheets, faxes, and mail templates specifically for course trainees.
* **`Internship Module`**: Opens the submenu containing templates and forms for student interns (e.g., paid stipend approvals, appraisals, coordinator nominations).
* **`🔒 Manage Masters`**: Opens the lock-protected Master Document Manager (requires admin password).

### Document & Form Cards
* **`Draft Noting`** / **`Draft Document`**: Opens a web form where you enter the details (Reference numbers, subject, body, signatures) to compile a document.
* **`Download Blank`**: Downloads a blank, printable copy of the template so it can be filled out by hand.
* **`✏️ Edit Info`**: Opens a secure popup to rename the template name or rewrite the description shown on the dashboard (requires admin password).

### Master Manager Area (🔒 Lock Unlocked)
* **`📥 Download`**: Downloads the raw Microsoft Word (`.docx`) template file to your computer so you can edit its layout or wording.
* **`📤 Upload Edit`**: Displays an upload form next to the template, allowing you to upload a modified Word template to replace the old one.
* **`⭐ Set Default`**: Sets a template as the default choice (pre-selected when users open forms).
* **`❌ Remove Default`**: Removes default status from a template.
* **`🗑 Delete`**: Permanently removes a template from the database and deletes its file (requires secondary admin password authorization).
* **`✨ Extract Variables & Create Form`**: Analyzes an uploaded `.docx` template, extracts all variables wrapped in `{{ }}`, and automatically builds a new form in the selected module.
* **`➕ Create New Master`**: Copies an existing master template under a new name as a starting point for editing.
* **`🔒 Lock`**: Relocks the manager area, hiding all template configuration options.

---

## 🛠️ Part 2: Step-by-Step Task Guides

Follow these simple, step-by-step guides to complete everyday tasks in the application.

### Task 1: How to Generate and Download a Document
1. Navigate to either the **Training Module** or **Internship Module**.
2. Select the category tab on the left sidebar (e.g., *Noting*, *Fax & Mail*, or *General Internships*).
3. Find the card for the document you want to create and click **`Draft Noting`** (or **`Draft Document`**).
4. Fill out all required fields on the web form.
5. Review the A4 Page Layout on the right side of the screen. Watch out for the dashed red lines representing page breaks to make sure your text fits.
6. Click the gold **`Generate Document`** button at the bottom of the page.
7. Click **`Download Document`** to save the generated file to your computer.

### Task 2: How to Access the Master Document Manager
1. Click **`🔒 Manage Masters`** on the top navigation bar.
2. Enter the Admin Password: `Tbrl0000`.
3. Click **`Unlock Manager`**. You can now view and edit the database settings.

### Task 3: How to Rename a Template or Edit its Description
1. Locate the template card on the **Training** or **Internship** dashboard.
2. Click **`✏️ Edit Info`** on the card.
3. Enter the Admin Password (`Tbrl0000`).
4. Enter the new **Template Name** or edit the **Description** box.
5. Click **`Save Changes`**. The dashboard will update immediately.

### Task 4: How to Edit a Template's Layout or Wording
1. Unlock the **Master Document Manager** (Task 2).
2. Find the template you want to edit and click **`📥 Download`** to download its `.docx` file.
3. Open the file in **Microsoft Word**.
4. Make your edits. You can change static text, fonts, margins, or add variables wrapped in double curly brackets (e.g. `{{MY_NEW_VARIABLE}}`). 
5. Save the file on your computer.
6. Return to the web manager, click **`📤 Upload Edit`** next to that template, select your saved file, and click **`Upload Updated Master`**.

### Task 5: How to Upload a Brand New Template (Code-Free Form Creation)
1. Write a new document in Microsoft Word. Wrap any fields you want the user to fill out in double curly brackets (e.g., `{{trainee_name}}`, `{{joining_date}}`).
2. Unlock the **Master Document Manager** (Task 2).
3. Scroll down to the **✨ Upload Dynamic Master (Code-Free)** card.
4. Enter a **Template Name** (e.g., *Stipend Certificate*).
5. Enter a **Description** of its purpose.
6. Choose the **Target Module** (*Training* or *Internship*).
7. Choose the **Target Sub-Module** (*Noting*, *Fax & Mail*, or *General*).
8. Click **Choose File** and select your `.docx` document.
9. Click **`✨ Extract Variables & Create Form`**. The system will create the form and display it in the selected dashboard tab.

### Task 6: How to Create a New Static Master (From Existing)
1. Unlock the **Master Document Manager** (Task 2).
2. Scroll down to the **➕ Create New Static Master (From Existing)** card.
3. Select the template you want to copy from the **Base this on** dropdown.
4. Enter the **New Master Name** and a **Description**.
5. Click **➕ Create New Master**. The system will copy the base template and register the new template record.
6. Download the newly created template, customize it in Microsoft Word, and upload it back (Task 4).

### Task 7: How to Run the App Offline (Double-Click Launchers)
If you are running the program on the offline Windows PC:
* **First-Time Setup**: Double-click **`install_offline.bat`**. This automatically creates the isolated environment and installs all dependencies from the `windows_required_libraries/` folder.
* **To Start the App**: Double-click **`run_app.bat`**. This launches your web browser and starts the application server automatically.

---

## 📌 Part 3: Core Concepts

* **Master Template**: A blueprint document (`.docx`) containing static layout formats, borders, official headers, and blank placeholders.
* **Dynamic Variable**: Any word wrapped in `{{ }}` inside a master template. The generator finds these words and turns them into editable text boxes on your screen.
* **A4 Page Layout Simulator**: A visual helper on the draft page showing how the document will print. The dashed red lines represent page boundaries so you can edit the text to prevent it from spilling onto a second page.
