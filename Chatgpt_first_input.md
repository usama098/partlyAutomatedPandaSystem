# Supported Housing Session JSON Specification

## Overview

Generate supported housing sessions from the tenant's **Referral / Initial Assessment** document.

The generator should:

- Read the tenant's initial assessment.
- Determine the support needs selected on the assessment.
- Populate the support areas from the assessment.
- Generate realistic weekly support sessions.
- Produce JSON only.
- Do not generate evidence pictures, only descriptions of the evidence picture.

---

# Session Rules

## Dates

- Generate one Move-In session on the move-in date.
- Generate weekly sessions every 7 days.
- Stop generating sessions before the move-on date.
- **Never generate a move-out session.**
- The move-on date is only used as the end date for generation.

Example

Move In:
27/03/2026

Move On:
20/05/2026

Sessions:

27/03/2026
03/04/2026
10/04/2026
17/04/2026
24/04/2026
01/05/2026
08/05/2026
15/05/2026

No session should exist on 20/05/2026.

---

# Session Length

Every support session should realistically represent approximately **one hour** of support.

Support notes should therefore be detailed.

Do not create short case notes.

---

# Support Notes

Support notes must:

- be written in first person
- be written from the support worker's perspective
- describe everything completed during the session
- naturally flow
- explain discussions
- explain demonstrations
- explain assessments
- explain planning
- explain agreements made
- explain follow-up actions

The notes should only discuss ONE support topic during each session.

Do not mix support areas.

---

# Support Note Summary

Each session must contain:

```json
"support_note_summary"
```

This should contain approximately **100 words**.

The summary must accurately summarise the support notes.

It must never introduce information that wasn't discussed.

---

# Evidence Picture Description

Each session contains

```json
"evidence_table"
```

Instead of generating an image, generate approximately **100 words** describing a printable worksheet/checklist/tracker that matches the session.

This description will later be sent to ChatGPT to generate the evidence sheet.

The description should describe:

- title
- layout
- sections
- tables
- tick boxes
- notes area
- printable A4 format
- professional support worksheet

---

# Date With Time

Each session's `session` object may contain

```json
"date_with_time"
```

This is optional. When present, it is printed directly beneath the normal
date on the generated document, in the same field.

Format:

```
Start Date/Time -- End Date/Time

DD/MM/YYYY, H:MM AM/PM -- H:MM AM/PM
```

Example:

```
Start Date/Time -- End Date/Time

14/09/2026, 2:00 PM -- 3:00 PM
```

---

# Contact Type

Always

```json
"In Person"
```

unless instructed otherwise.

---

# Support Areas

Support areas and sub areas must match EXACTLY.

Never change wording.

Never change capitalisation.

Never change spelling.

Use these exact values.

## Achieve Economic Wellbeing

- Accessing Benefits
- Budgeting
- Reducing Debt
- Setting up a bank or savings account
- To learn how to shop wisely
- To recoup monies owed

## Be Healthy

- To better manage improve mental health
- To better manage improve physical health
- To follow a healthy diet
- To maintain good personal hygiene
- To reduce substance misuse
- To register with a Dentist
- To register with a GP

## Enjoy and Achieve

- Accessing education or training
- Accessing employment
- Accessing leisure cultural and faith activities
- Accessing volunteering
- Move on
- Support with equality and diversity

## Make a Positive Contribution

- Establishing positive support networks
- To address anti-social behaviour
- To address offending behaviour

## Stay Safe

- To develop independent living skills
- To maintain accommodation
- To minimise risk of harm

---

# list_of_support_areas_covered

The JSON contains

```json
"list_of_support_areas_covered"
```

This must be populated by reading the tenant's **Initial Assessment**.

Specifically, the support plan page.

Only include support areas that have been selected as required.

Example

```json
[
  {
    "area_covered": "Be Healthy",
    "sub_area_covered": "To register with a GP"
  },
  {
    "area_covered": "Be Healthy",
    "sub_area_covered": "To follow a healthy diet"
  },
  {
    "area_covered": "Stay Safe",
    "sub_area_covered": "To maintain accommodation"
  }
]
```

Do not include support areas that were not selected.

---

# Session Area Selection

Each session must use one of the support areas from

```json
list_of_support_areas_covered
```

Do not create sessions for support areas that were not selected in the Initial Assessment.

---

# JSON Structure

```json
{
  "tenant_name": "string",
  "gender": "male | female (optional, defaults to \"male\")",

  "list_of_support_areas_covered": [
    {
      "area_covered": "Main support area",
      "sub_area_covered": "Exact sub-area"
    }
  ],

  "move_in_date": "DD:MM:YYYY",

  "move_out_date": "DD:MM:YYYY",

  "sessions": [
    {
      "session_number": 1,

      "session": {

        "session_date": "DD:MM:YYYY",

        "date_with_time": "Start Date/Time -- End Date/Time\n\nDD/MM/YYYY, H:MM AM/PM -- H:MM AM/PM",

        "property_address": "string",

        "contact_type": "In Person",

        "area_covered": "Main support area",

        "sub_area_covered": "Exact sub-area discussed",

        "support_notes": "Detailed one-hour first-person support notes 400 to 600 words",

        "support_note_summary": "Approximately 100 words",

        "any_actions_to_be_completed_by_service_user": [
          "Action 1",
          "Action 2"
        ],

        "any_action_to_be_taken_by_support_worker": [
          "Action 1",
          "Action 2"
        ]
      },

      "evidence_table": "markdown table which is produced by chatgpt"
    }
  ]
}
```

# Overall Goal

The generated JSON should contain enough information that another application can:

1. Generate all support session documents.
2. Generate evidence pictures using the evidence picture descriptions.
3. Maintain continuity between sessions.
4. Produce realistic supported housing records that reflect approximately one hour of support per session.
5. Only generate sessions for support needs identified in the Initial Assessment.