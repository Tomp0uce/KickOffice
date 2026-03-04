# Product Requirements Document (PRD): KickOffice

## 1. Product Vision & Target Audience
**KickOffice** is an internal, enterprise-grade AI assistant seamlessly integrated into the Microsoft Office suite (Word, Excel, PowerPoint, Outlook). 
Its goal is to boost productivity, automate repetitive tasks, and assist in complex document/data manipulation while keeping all data flows secure and internal.

**Target Audience:**
All employees within the organization, specifically tailored to handle the diverse needs of:
* High-level Engineers (Hardware, Software, Firmware, Mechanical)
* Project Managers
* Accounting & Finance
* Sales Representatives
* Administrative Services

## 2. Deployment, Quotas & Telemetry
* **Deployment:** Distributed internally. Users download/install the add-in via a manifest file hosted on the company's internal SharePoint or local server.
* **Monetization & Quotas:** No user-facing subscriptions or internal billing. Quotas, rate limiting, and access control are centrally managed by the internal LLM gateway (LiteLLM).
* **Telemetry & DLP (Data Loss Prevention):** The add-in itself does not log telemetry or block sensitive data. All AI telemetry, auditing, and DLP filtering are strictly delegated to the internal LiteLLM gateway.

## 3. Core Cross-App Capabilities
Regardless of the Office application being used, the assistant provides the following core features:
* **Chat Interface:** A conversational UI available in English and French, capable of responding in multiple languages.
* **Autonomous Document Agent:** The AI automatically understands the context of the open document (size, structure, current selection) without the user needing to explain it. It can plan and execute multi-step modifications autonomously.
* **Quick Actions:** One-click contextual buttons to perform immediate tasks (e.g., Translate, Summarize, Polish) without typing a prompt.
* **File Analysis:** Users can upload external files (PDF, DOCX, XLSX, CSV) into the chat to ask questions or extract data to be used in their current Office document.
* **Image Generation:** Ability to prompt the AI to generate images and insert them directly into the document (Word/PowerPoint).

## 4. Application-Specific Features (Deep Dive)

### 4.1. Microsoft Excel
**Capabilities:**
* **Data Discovery:** Automatically understands the workbook structure (number of sheets, rows, columns, active sheet).
* **Data Analysis & Formatting:** Can sort data, apply conditional formatting, and clean datasets.
* **Formula Generation:** Understands the user's goal and generates complex, localized Excel formulas.
* **Visualization:** Can autonomously read data ranges and create targeted charts or pivot tables anywhere in the workbook.
* **Structural Edits:** Can manipulate rows, columns, and sheet structures safely.
**Out of Scope / Constraints:**
* No interaction or modification of VBA Macros.
* No management or overriding of locked/password-protected sheets.

### 4.2. Microsoft Word
**Capabilities:**
* **Drafting & Editing:** Can write new sections, summarize long texts, translate paragraphs, and adjust the tone of the document (e.g., academic, formal).
* **Format Preservation:** When editing existing text, the AI surgically replaces words without destroying the user's existing formatting (bolding, italics, specific fonts).
* **Table Management:** Can generate, populate, and format tables based on prompt instructions or external uploaded data.
**Out of Scope / Constraints:**
* **Crucial Requirement:** The AI *MUST* utilize the native "Track Changes" feature whenever it modifies existing text, allowing the user to review and accept/reject the AI's edits.

### 4.3. Microsoft PowerPoint
**Capabilities:**
* **Presentation Structuring:** Automatically understands the slide count, current slide layout, and titles.
* **Content Generation:** Can generate bullet points, draft slide content, and adjust the text to be punchier or more concise.
* **Speaker Notes:** Can automatically generate conversational speaker notes based on the bullet points of a slide.
* **Visuals:** Can generate AI images and insert them directly into the active slide.
**Out of Scope / Constraints:**
* No modification or interaction with the Slide Master (Masque des diapositives).

### 4.4. Microsoft Outlook
**Capabilities:**
* **Email Assistance:** Can draft replies, summarize long emails, formalize tone, and extract action items/tasks from the current email.
* **Smart Context:** Automatically reads the subject, sender, and body of the *currently open* email to provide highly contextual replies.
* **Calendar Integration:** Can access the user's calendar strictly in "Read-Only" mode to check availability and propose meeting slots in email drafts.
**Out of Scope / Constraints:**
* The AI cannot read the user's entire inbox or emails other than the one currently open.
* To manage very long email threads, the system must truncate older historical messages while keeping the most recent ones.

## 5. Non-Functional Requirements & Edge Cases

### 5.1. Context Limits & Document Size
* The system utilizes the GPT-5.1 model, which has a large context window.
* **Requirement:** The application must dynamically calculate the estimated token size of the document/email history + the system prompt. If the calculated size exceeds the GPT-5.1 safe limit, the UI must block the action and display a specific, user-friendly error message indicating the document is too large to be processed entirely.

### 5.2. UX Feedback & Timeouts
* Because the AI performs complex reasoning that can take a long time, the UI must provide clear, dynamic visual feedback.
* **Requirement:** Whenever the AI is processing for an extended period, a status bar or dynamic visual indicator (e.g., a pulsing icon, a spinning wheel, or a "Reasoning in progress..." text step) must be displayed to assure the user the system has not frozen.

### 5.3. Offline & Connectivity Handling
* **Requirement:** There is no offline mode. If the user's internet connection drops or the internal server is unreachable, the UI must immediately intercept the failure and display a specific "Connection Lost / Offline" error message, halting any ongoing AI task safely.
* Concurrency (handling user edits while the AI is typing) is out of scope for now.

### 5.4. Data Retention & Privacy
* **File Retention:** Any external file uploaded to the chat is strictly tied to that specific conversation's lifecycle. 
* **Expiration:** Conversations (and their associated uploaded files) must be automatically hard-deleted after 1 month (30 days). If a user manually deletes a conversation earlier, the associated files must be deleted immediately.
* **Logging:** Application logs (prompts, errors, contexts) should only be stored/retained when the environment is set to Development/Debug mode.