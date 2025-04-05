# QA Evaluation System

A modular web application built on Google Apps Script and Google Sheets to manage quality assurance audits, evaluations, disputes, and administrative settings for support teams.

## Features

- **Dashboard**: At-a-glance metrics for audits, evaluations, and dispute performance.
- **Audit Queue**: Manage incoming tasks to be evaluated.
- **Evaluation Form**: Custom evaluation form based on task type and scoring.
- **Disputes**: Submit, review, and resolve evaluation disputes.
- **Admin Panel**: Manage users, questions, and system settings.
- **Character-limited feedback fields**, visual progress indicators, and dynamic UI interactions.

## Folder Structure

```bash
.
├── src/                  # Google Apps Script backend logic
│   ├── App.gs
│   ├── Users.gs
│   ├── Settings.gs
│   ├── Questions.gs
│   ├── Audits.gs
│   ├── Evaluations.gs
│   ├── Disputes.gs
│   ├── SpreadsheetSetup.gs
│   ├── CacheUtils.gs
│   └── Utils.gs

├── ui/                   # HTML frontend views and styles
│   ├── Index.html
│   ├── Styles.html
│   ├── JavaScript.html
│   ├── DashboardView.html
│   ├── AuditQueueView.html
│   ├── EvaluationFormView.html
│   ├── EvaluationsView.html
│   ├── ViewEvaluationView.html
│   ├── DisputesView.html
│   ├── DisputeFormView.html
│   ├── ResolveDisputeView.html
│   └── AdminView.html

├── test/                 # Optional scripts for testing
│   ├── seedData.gs
│   └── testHelpers.gs

├── docs/
│   └── changelog.md

└── README.md
```

## Setup Instructions

### 1. Copy Files into a New Google Apps Script Project

- Open [https://script.new](https://script.new) to create a new Apps Script project.
- Replace the `Code.gs` file with the modular `.gs` files from `/src`.
- Create each `.html` file in the **Editor's UI** and paste content from `/ui`.

### 2. Link to a Google Spreadsheet

- The app uses the active spreadsheet as its database (sheet tabs like `users`, `auditQueue`, etc.).
- Use the script menu or call `setupSpreadsheet()` to initialize all sheets and headers.

### 3. Deploy the App

- In the Script Editor: `Deploy > Test deployments > Web App`
- Select:
  - **Execute as**: *Me*
  - **Who has access**: *Anyone*
- Click **Deploy** and share the generated URL.

## Key Modules Explained

| Module         | Purpose                                                                 |
|----------------|-------------------------------------------------------------------------|
| `App.gs`       | Entry point: HTML UI loader (via `doGet()` and `include()`)             |
| `Users.gs`     | CRUD operations for users, role management                              |
| `Settings.gs`  | Get/save global settings like character limits                          |
| `Questions.gs` | CRUD operations for QA questions by task type                           |
| `Audits.gs`    | Fetch, update, and lock audit records                                   |
| `Evaluations.gs`| Evaluation creation, score calculation, and storage                    |
| `Disputes.gs`  | Dispute submission, resolution, and status tracking                     |
| `CacheUtils.gs`| Caching wrappers for performance optimization                           |
| `Utils.gs`     | Sheet-to-object mapping, formatting, and helpers                        |

## Sheet Structure

| Sheet Name      | Description                              |
|------------------|------------------------------------------|
| `users`          | User profiles and roles                  |
| `auditQueue`     | Tasks pending evaluation                 |
| `evalSummary`    | Evaluation-level summary data            |
| `evalQuest`      | Individual evaluation question responses |
| `questions`      | Configurable scoring questions           |
| `disputesQueue`  | Evaluation disputes & resolution logs    |
| `settings`       | System-wide configuration values         |

## Development Notes

- Views are dynamically shown/hidden by JavaScript and do **not** reload pages.
- Role-based UI access: Admin panel is shown only for `admin` and `qa_manager`.
- Character limits for feedback are customizable via **Admin > Settings**.

## Contributions

Feel free to extend this app with features like:
- Audit imports from external systems
- Role-based permissions enforcement
- Evaluation versioning or drafts
- Chart-based reporting (via Google Charts)

## License

MIT — Use and modify freely.
