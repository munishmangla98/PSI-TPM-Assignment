# PSI TPM Weekly Status Report Generator

A small command-line utility that reads the PSI assignment tracker and creates one Markdown status report per project.

## Run
Python 3.9+ is recommended.

```bash
python -m pip install -r requirements.txt
python generate_report.py sample_tracker.xlsx --out reports
```

Expected output:
- `reports/kavach_weekly_status.md`
- `reports/ledger_weekly_status.md`

## Workbook format
The input workbook must contain:
- `Summary`: Project, Client, Status, Next milestone, Milestone date, Main problem, TPM decision / next step.
- `Risks and Issues`: ID, Project, Type, Description, Probability, Impact, Owner, Action / mitigation, Due date, Status.

Use the supplied `sample_tracker.xlsx`.

## What the tool does and does not do
It groups risks by project and copies the tracker facts into Markdown. It does not infer progress, change RAG status, email clients, or make release decisions. A TPM must validate all dates, owners and status before sharing a report externally.

## AI usage disclosure
AI assistance was used to help draft the initial code structure, README and test ideas. I reviewed the assignment requirements, checked the workbook columns, ran the tool against the supplied tracker and corrected blank-cell handling. Project decisions and client-facing wording should be reviewed and explained by the candidate.

## One AI-generated error and correction
An initial draft rendered empty Excel cells as the literal text `None`. I corrected this by adding a `clean()` helper that displays blank values as `Not provided`. I tested the tool against the sample workbook and reviewed the generated reports.

## Prompts used with AI coding assistance
1. "Build a Python command-line tool using openpyxl to read Summary and Risks and Issues worksheets and generate a Markdown status report per project. Do not invent status or progress."
2. "Add validation for required worksheets, safe blank-cell handling and a configurable output directory."
3. "Review the script and write concise installation and usage instructions."

## Limitations / next steps
This is a lightweight local prototype. A production version could add a PDF template, report history, approval workflow, integration with Jira/CRM and access controls.
