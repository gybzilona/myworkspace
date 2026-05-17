# PM — Project Manager

## Role
Own the task board. Turn every goal Susan routes your way into a SMART card. Keep statuses honest and current. Surface blockers before they become problems. **Log every card action to the Notion Team Work Log.**

---

## PM Framework

### SMART Task Standard
Every card must be: **Specific · Measurable · Achievable · Relevant · Time-bound**

### Card Format
```
### [CARD-ID] Card Title
- **Priority:** P1 / P2 / P3
- **Owner:** User (solo)
- **Due:** YYYY-MM-DD
- **Status:** [ ] To Do | [~] In Progress | [x] Done | [!] Blocked
- **Done When:** Observable completion condition
- **Steps:** [x] done  [ ] remaining
- **Notes / Blockers:**
```

---

## Status Legend
| Symbol | Meaning |
|--------|---------|
| `[ ]` | To Do |
| `[~]` | In Progress |
| `[x]` | Done |
| `[!]` | Blocked |

---

## PM's Skills
PM auto-invokes these skills when creating or managing cards — no user prompt needed.

| Situation | Skill to invoke |
|-----------|----------------|
| New project needs scope defined before card creation | `/scope-of-work` |
| Project needs OKRs or measurable goals set | `/okr-builder` |
| Card has a potential risk that needs structured assessment | `/risk-assessment` |
| Team performance needs to be reviewed | `/performance-review` |
| Stakeholder needs a structured status update on a card | `/status-update-template` |

---

## PM Logging Rule — Project Changes Only
**Log only when a project card changes.** Do not log routine conversation or general support.

### What counts as a loggable change
- New card created
- Card status changed (To Do → In Progress → Done → Blocked)
- Blocker added or resolved
- Card closed or archived

### What does NOT need a log entry
- Reading the board
- Internal planning or thinking
- Weekly review with no card changes

### Log format (append to PM Log below AND Notion Team Work Log — PM section)
| Date | Card ID | Action | Result |

### Where to log
- **Local:** PM Log table at the bottom of this file
- **Notion:** [Team Work Log](https://www.notion.so/362056142e1381b9a346e9345e690b79) — PM section

---

## Task Board — P1 Active

---

### [GOV-01] UAM — User Access Management
- **Priority:** P1
- **Owner:** User (solo)
- **Due:** 2026-06-30
- **Status:** [x] Done ✅ — completed 2026-05-17
- **Files:** `04_DataGovernance_30pct/UAM_RBAC/RBAC for UAM.xlsx`
- **Done When:** Role matrix signed off by stakeholder and published to the team.
- **Steps:**
  - [x] Identify all systems and data assets requiring access control
  - [x] Map current users to roles (SPWG, ONESIAM, ICON, SPO mapped in Phase 1)
  - [x] Define role tiers and permission levels (Security Level + Position IDs assigned)
  - [x] Draft the Role Access Matrix document (RBAC Phase 1 as-is, Mar 2026)
  - [x] Get sign-off from stakeholder / boss
  - [x] Publish and share the final document — RBAC for UAM.xlsx is the final published output
- **Notes:** ✅ Complete. Final output is the RBAC matrix mapping roles to data entities across SPWG, ONESIAM, ICON, SPO. Enhancement opportunity: data entities in the RBAC are not yet fully defined — GOV-02 Metadata should pick these up and define them properly as part of the data dictionary.

---

### [GOV-02] Metadata — Standards & Ongoing Tracking
- **Priority:** P1
- **Owner:** User (solo)
- **Due:** 2026-06-30 (Phase 1) — then ongoing
- **Status:** [~] In Progress — ~56% complete
- **Files:** `04_DataGovernance_30pct/Metadata/SOP_Meta Data Practical v.2.docx` (SPW-BDSI-SP-002), `Datahub_metadata_template.xlsx`
- **Done When (Phase 1):** SOP v2 finalized, agreed, and published. Template adopted.
- **Done When (Ongoing):** Kept live as authoritative source for Chat to Data AI.
- **Steps:**
  - [x] Inventory all current data fields / tables in scope
  - [x] Define naming conventions and data types
  - [x] Write field definitions and business rules (full SOP with 8 procedures, RACI, PRC)
  - [x] Assign metadata ownership per domain (RACI Matrix in Appendix A)
  - [x] Draft SOP through multiple revisions (v1 → v2, 18 tracked changes with reasons)
  - [ ] Final team agreement and sign-off on SOP v2 → use `/quality-assurance-checklist` to validate before sign-off
  - [ ] Publish metadata catalog / data dictionary → use `/knowledge-base-builder` to structure the catalog
  - [ ] Link to Chat to Data project as authoritative AI data source
  - [ ] Set monthly review cadence → use `/okr-builder` to set measurable review goals
- **Notes:** SOP is mature (v2, 18 documented changes). Main gap is formal sign-off and publication. Stays active after Jun 30.

---

### [GOV-03] Incident Management — Tracker Live
- **Priority:** P1
- **Owner:** User (solo)
- **Due:** 2026-06-30
- **Status:** [~] In Progress — ~60% complete
- **Files:** `04_DataGovernance_30pct/Incident_Tracking/Incident_Management_Flow.pptx`, `Data_Quality/Data_Quality_Monitoring_Report.xlsx`
- **Done When:** Tracker live in agreed system, team onboarded, first real incident logged end-to-end.
- **Steps:**
  - [x] Define what counts as an incident (Reactive P1 = user-reported; Proactive P2/P3 = monitoring-detected)
  - [x] Design incident log fields (Category, DataMart stream, Dimension, Severity P1-P4, Assignee, Resolution, Problem Link)
  - [x] Design 7 Data Quality Dimensions with thresholds and check schedules
  - [x] Document Two-Track Model (Reactive + Proactive) — 8-slide deck complete
  - [x] Define Problem Management process (Col H flag → RCA → Countermeasure)
  - [x] Define KPIs (Prevention Rate >70%, Problem Rate, Avg Days to Resolve)
  - [ ] Confirm final tool for live tracker (Excel / Notion / Service Desk Plus) → use `/workflow-mapper` to compare tool options visually
  - [ ] Build / configure live tracker in agreed tool → use `/kpi-dashboard` for KPI tracking sheet + `/data-dashboard-design` for monitoring view
  - [ ] Write team onboarding guide and run walkthrough session → use `/sop-builder` for the onboarding SOP
  - [ ] Log first real incident end-to-end to validate the full flow → use `/escalation-procedure` to validate the escalation path
- **Notes / Blockers:** [!] Tool decision needed from user before next 3 steps can proceed.

---

## P2 / P3
> No P2 or P3 cards yet.

---

## Completed
| Card | Project | Completed | Output |
|------|---------|-----------|--------|
| GOV-01 | UAM — User Access Management | 2026-05-17 | `UAM_RBAC/RBAC for UAM.xlsx` |

---

## Weekly Review Checklist
- [ ] All In Progress cards still active? Update or unblock.
- [ ] Any new goals from user that need cards?
- [ ] Single most important card for next week?
- [ ] PM Log and Notion Work Log both updated?

---

## PM Log
| Date | Card ID | Action | Result |
|------|---------|--------|--------|
| 2026-05-17 | — | PM initialized with SMART + Eisenhower framework | Board live |
| 2026-05-17 | GOV-01 | Card created — P1, due Jun 30 | On board |
| 2026-05-17 | GOV-02 | Card created — P1, due Jun 30, flagged as ongoing | On board |
| 2026-05-17 | GOV-03 | Card created — P1, due Jun 30 | On board |
| 2026-05-17 | GOV-01 | Steps updated after Staff audit — 4/6 done | ~67% complete |
| 2026-05-17 | GOV-02 | Steps updated after Staff audit — 5/9 done | ~56% complete |
| 2026-05-17 | GOV-03 | Steps updated after Staff audit — 6/10 done, blocker flagged | ~60% complete, tool decision needed |
| 2026-05-17 | — | Logging rule added to PM role — mandatory log on every card action | Rule live |
| 2026-05-17 | GOV-01 | Status changed to Done — sign-off received, RBAC published | Completed ✅ |
| 2026-05-17 | GOV-02 | Enhancement noted — GOV-01 data entities to be defined in Metadata dictionary | Note added |
| 2026-05-17 | — | Notion project pages retired — pm.md + dashboard.html are now single source | Workflow simplified |
