# CLEAR WebAppAuth Portal — Experience Playbook

## 1. Purpose
This playbook defines the exact user experience for the CLEAR WebAppAuth portal. It ensures teammates at every role level — employee, lead, director — have a consistent, role‑appropriate, and culturally aligned digital experience when they log in.

The goal is to:
- Provide **clarity** of current status.
- Ensure **fairness** with data integrity.
- Uphold **security** standards in login and session management.
- Match CLEAR’s cultural tone (accountability + encouragement).
- Mirror modern Chick‑fil‑A digital design patterns for familiarity and trust.

---

## 2. Core Principles
- **Clarity:** No hidden data. Each teammate sees exactly what applies to them.
- **Fairness:** All numbers match Events and PDFs. No surprises.
- **Security:** Email/password or OTP; sessions expire after 24 hours.
- **Role‑Specific Views:** Dashboards show only what teammates need to act on.
- **Consistency:** Language matches disciplinary system and performance standards (e.g., *Policy‑Protected*, *Verified Medical*).
- **Design Alignment:** Visual style mirrors Chick‑fil‑A’s branded web apps (Supply Chain Signal, Inventory Counts, etc.).

---

## 3. Login Flow

1. **Access:** Teammate opens the CLEAR portal link.
2. **Credentials:** Enter email/password, or request an OTP code.
3. **Verification:** Directory confirms credentials using salted + peppered hashes.
4. **Session:** If successful, issue a session ID (valid 24h).
5. **Redirect:** User is routed to their dashboard (`doGet` routes by role).

*Design inspiration:* Login screen mirrors Chick‑fil‑A’s supply chain login — centered form, red branding, clean lock/password iconography.

---

## 4. Dashboards by Role

### Employee Dashboard
**Tone:** Neutral, factual, encouraging.

**Sections:**
- **Welcome card:** Employee name + role.
- **Current Status:**
  - Rolling points (effective).
  - Grace availability (Minor, Moderate, Major credits).
  - Probation/performance path flags.
- **History:** Table of last 5 events (Date | Infraction | Points | PDF link).
- **Design:** Clean, minimal; feels supportive, not punitive.
  - Card‑based layout similar to *Open Orders* in Signal.
  - Status chips (green = “Available”, red = “On Probation”).

### Lead Dashboard
**Tone:** Leadership‑oriented, coaching focused.

**Sections:**
- **Team Overview Stats:**
  - Pending milestones (team scope).
  - Active probations.
  - Grace requests pending director review.
- **Quick Access:**
  - Employee search.
  - Pending milestones (view‑only).
  - Reports (read‑only).
- **Design:**
  - Grid of overview cards like *Inventory Counts overview*.
  - Emphasis on team performance, not disciplinary action.

### Director Dashboard
**Tone:** Decisive, corrective, accountability‑focused.

**Sections:**
- **Stats:**
  - Pending milestones (claimable).
  - Active probations.
  - Grace requests (approve/deny).
  - Monthly event count.
- **Tools:**
  - Employee search (full access).
  - Pending milestones → assign & trigger PDF.
  - Reports (monthly, employee history PDFs).
  - Bulk operations (queue missing PDFs, run audits).
- **Design:**
  - Rich navigation with actionable panels, similar to Signal’s *Counts* tab.
  - Pending list styled like *count in progress* banners (yellow/orange highlight).

---

## 5. Integration Points
- **Sheets Tabs:** Events, Rubric, PositivePoints, Audit, Directory.
- **Scripts:**
  - `pointsEngine.js` → milestones, probation failure checks.
  - `docService.js` → write‑up and consequence PDF generation.
  - `grace.js` → grace + positive points ledger.
  - `webappAuthDirectory.js` → login, session, role resolution.
- **Slack Channels:**
  - `#documentation` → disciplinary write‑ups.
  - `#leaders` → milestone pending/claimed notifications.

---

## 6. Guardrails
- Employees: **see only self**.
- Leads: **view but not act** on milestones/grace.
- Directors: **full permissions** but every action audited.
- All dashboards: **mobile‑friendly, CLEAR‑branded, modern UI**.

---

## 7. Notifications
- **In‑App Banners:** Timely alerts inside dashboards (e.g., “New milestone assigned”, “Grace request submitted”).
- **Email Digests:** Daily or weekly summaries of events, milestones, and statuses.
- **Slack DMs:** Real‑time nudges to employees and leaders for critical events (probation start, milestone claim).
- **Configurable Settings:** Directors can choose which notifications are pushed to each channel.

---

## 8. Analytics & Reporting
- **Trend Dashboards:** Directors and leads can view rolling point trends, milestone frequencies, and grace usage over time.
- **Export Options:** Ability to export reports (PDF/CSV) for meetings or compliance.
- **Comparisons:** Month‑over‑month and team vs. team performance charts.
- **Visualizations:** Line charts for trends, bar charts for category breakdowns, pie charts for credit usage.
- **Access Control:** Directors have full analytics access; leads see team‑scoped reports; employees see only personal history graphs.

---

## 9. Tone of Voice
- **Employee:** “Here’s your current status. You’ve got tools to improve.”
- **Lead:** “Here’s your team’s health. Support and coach.”
- **Director:** “Here’s what needs corrective action. Claim and resolve.”

---

## 8. Next Steps
- Wireframe mockups for all three dashboards.
- Map data calls (`getMyOverviewForEmail`, `getDirectorDashboardData`, etc.).
- Define Slack notification triggers in parallel with dashboard actions.
- Run employee feedback pilot to validate clarity and fairness.

---

## 9. Appendix — Visual Wireframe & Design Notes

The following notes connect inspiration from Chick‑fil‑A apps (see attached images):
- **Login screen:** Large centered password field, logo prominent, rounded verify button (blue/red).
- **Navigation:** Left‑hand vertical bar with clear icons (home, orders, counts, reports, settings).
- **Cards:** Rounded corners, white/gray backgrounds, colored status chips.
- **Counts/Events:** Rows with status tags (e.g., “In Progress”, “Shipped”).
- **Loading:** Circular Chick‑fil‑A logo spinner animation for transitions.

**Employee Dashboard Wireframe:**
```
[Welcome Card]
| Name + Role + Logout Button |

[Status Card]
| Points: XX | Grace: Available |
| Probation: No                 |

[History Table]
| Date | Infraction | Points | PDF |
```

**Lead Dashboard Wireframe:**
```
[Overview Cards]
| Pending Milestones | Probations |
| Grace Requests     | Events     |

[Quick Tools]
| Employee Search | Reports |
```

**Director Dashboard Wireframe:**
```
[Stats Grid]
| Pending Milestones | Probations |
| Grace Requests     | Events     |

[Pending List]
| Employee | Milestone | Claim Btn |

[Nav Tools]
| Employee Search | Reports | Bulk Ops |
```

*Design Note:* All wireframes should use the **Signal app visual pattern** — bold numbers, red call‑to‑action buttons, pill‑shaped status labels, and left‑hand icon navigation.

