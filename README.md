# CloudFuze Migrate – Incentive Calculator

Incentive Estimation Application for CloudFuze Migrate. Admins upload deal data via Excel; the system calculates Rep and Team Lead incentives using revenue-based slabs.

## Features

- **Role-based access**: Admin (full) | Sales Manager (team view) | Sales Rep (self view)
- **Excel upload**: Deal Name, Deal Owner, Amount, Payment Status, Team
- **Validation**: Mandatory fields, numeric amounts, existing users & teams
- **Revenue-based slabs** (fixed):
  - ≥ $100,000 → 4%
  - ≥ $75,000  → 2%
  - ≥ $30,000  → 1%
  - < $30,000  → 0%
- **Workflow**: Upload → DRAFT → Review → Finalize → Incentives generated
- **Reports**: Rep-wise, Team-wise, Excel/CSV export

## Prerequisites

- Python 3.9+
- PostgreSQL server (default port 5432)

## Installation

```bash
cd "c:\Users\SakshiPriya\OneDrive - CloudFuze, Inc\Desktop\Compensation_Tool"
python -m venv venv
venv\Scripts\activate
pip install -r requirements.txt
```

## Database Configuration

Create a PostgreSQL database (e.g. `cloudfuze_incentive`). Set environment variables:

**PowerShell:**
```powershell
$env:COMP_DB_HOST="localhost"
$env:COMP_DB_PORT="5432"
$env:COMP_DB_NAME="cloudfuze_incentive"
$env:COMP_DB_USER="your_user"
$env:COMP_DB_PASSWORD="your_password"
```

The app creates all tables on first run and seeds:
- Teams: SMB, Enterprise, Account Management
- Incentive slabs
- Default admin: `admin@cloudfuze.com` / `Admin@123` (change in production!)

## Running

```bash
streamlit run app.py
```

Open the URL (typically `http://localhost:8501`).

## Usage

1. **Login** with Admin credentials.
2. **User Management** (Admin): Add Sales Reps and Sales Managers with team assignment.
3. **Upload**: Download the sample template, fill Deal Name, Deal Owner (name/email), Amount, Payment Status (PAID/UNPAID/PARTIALLY_PAID), Team.
4. **Save as Draft** to store deals and review.
5. **Finalize** to aggregate revenue, apply slabs, and generate Rep & Team incentives.
6. **Export** to Excel or CSV.

## Excel Template Columns

| Column         | Description                    | Example          |
|----------------|--------------------------------|------------------|
| Deal Name      | Required                       | Deal ABC         |
| Deal Owner     | User full name or email        | John Doe         |
| Amount         | Revenue (numeric)              | 50000            |
| Payment Status | PAID, UNPAID, PARTIALLY_PAID   | UNPAID           |
| Team           | SMB, Enterprise, Account Mgmt  | Enterprise       |

## Teams

- SMB
- Enterprise
- Account Management

Teams are seeded on first run. Deal Owner and Team must exist in the system before upload validation passes.
