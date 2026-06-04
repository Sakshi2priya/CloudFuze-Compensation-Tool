# Incentive Calculator Application – CloudFuze Migrate
## Build Plan & Implementation Roadmap

### Technology Stack
- **Backend/UI**: Python + Streamlit
- **Database**: MySQL (existing setup)
- **Excel Handling**: pandas + openpyxl
- **Auth**: Session-based (Streamlit) with bcrypt for password hashing

### Phase Overview

| Phase | Component | Status |
|-------|-----------|--------|
| 1 | Database schema (Teams, Users, Deals, etc.) | 🚧 In Progress |
| 2 | Incentive calculation engine (revenue slabs) | ⏳ Pending |
| 3 | Excel upload, validation, DRAFT/Finalize flow | ⏳ Pending |
| 4 | Authentication & role-based access | ⏳ Pending |
| 5 | Admin Dashboard | ⏳ Pending |
| 6 | Member Dashboard | ⏳ Pending |
| 7 | Reports & Export (Excel/CSV) | ⏳ Pending |

### Implementation Order
1. Database schema & seed data (teams, slabs, admin user)
2. Incentive calculator (revenue-based slabs)
3. Excel upload & validation service
4. Auth & session management
5. Admin flow: Upload → Review → Finalize → Calculate
6. Member flow: Read-only incentives view
7. Reports & export
