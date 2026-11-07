#!/usr/bin/env python3

"""
Create one Markdown weekly status report per project
using data from the PSI TPM Excel project tracker.

Input:
    T5_Project_Tracker.xlsx

Required Excel sheets:
    1. Summary
    2. Risks and Issues

Output:
    One Markdown report (.md) for each project.
"""

# argparse: Handles command-line arguments provided by the user.
import argparse

# Path: Provides convenient and cross-platform file and folder handling.
from pathlib import Path

# date: Gets the current date for the weekly status report.
from datetime import date

# load_workbook: Reads data from Excel workbooks.
from openpyxl import load_workbook


def clean(value, fallback="Not provided"):
    """
    Clean Excel cell values before using them in reports.

    - Returns a fallback value if the cell is empty.
    - Removes unnecessary leading and trailing whitespace.
    - Converts non-string values into strings.

    This prevents blank or missing Excel values from
    creating incomplete report fields.
    """

    if value is None or str(value).strip() == "":
        return fallback

    return str(value).strip()


def load_data(path):
    """
    Load project information and risk details from Excel.

    Parameters:
        path: Location of the Excel tracker file.

    Returns:
        A tuple containing:
        1. Project summary records.
        2. Risk and issue records.

    The workbook must contain the required worksheets:
    'Summary' and 'Risks and Issues'.
    """

    # Load the Excel workbook.
    # data_only=True reads the last saved values of formula cells
    # instead of returning their Excel formulas.
    wb = load_workbook(path, data_only=True)

    # Define the worksheets required by this reporting tool.
    required = {"Summary", "Risks and Issues"}

    # Validate that all required worksheets exist.
    if not required.issubset(set(wb.sheetnames)):
        raise ValueError(
            "Workbook must contain Summary and Risks and Issues tabs."
        )

    def rows(sheet):
        """
        Convert worksheet data into a list of dictionaries.

        The first row contains column headers.
        Each subsequent non-empty row becomes a dictionary.

        Example:
            {"Project": "Ledger", "Status": "AMBER"}
        """

        # Read column names from the first row.
        headers = [c.value for c in sheet[1]]

        # Read all remaining rows and map each value to its header.
        # Completely empty rows are skipped.
        return [
            dict(zip(headers, row))
            for row in sheet.iter_rows(
                min_row=2,
                values_only=True
            )
            if any(v is not None for v in row)
        ]

    # Extract project information and associated risks separately.
    projects = rows(wb["Summary"])
    risks = rows(wb["Risks and Issues"])

    # Return both datasets for report generation.
    return projects, risks


def render(project, risks):
    """
    Generate the Markdown content for an individual project.

    Parameters:
        project: A dictionary containing project information.
        risks: A list containing risk and issue records.

    Returns:
        A formatted Markdown string for the project report.
    """

    # Get the project name, using a fallback if it is missing.
    name = clean(project.get("Project"))

    # Filter risks so that only risks belonging to the
    # current project appear in its report.
    related = [
        r for r in risks
        if clean(r.get("Project")) == name
    ]

    # Build the report section by section.
    # Each item represents a line in the Markdown report.
    lines = [

        # Report title.
        f"# Weekly Status Report – {name}",
        "",

        # Basic project information.
        f"**Client:** {clean(project.get('Client'))}  ",
        f"**Reporting date:** {date.today().strftime('%d %b %Y')}  ",
        f"**Overall status:** {clean(project.get('Status'))}",
        "",

        # Upcoming milestone and target date.
        "## Next milestone",
        f"- **Milestone:** {clean(project.get('Next milestone'))}",
        f"- **Target:** {clean(project.get('Milestone date'))}",
        "",

        # Current project situation and major problem.
        "## Current position",
        clean(project.get("Main problem")),
        "",

        # TPM's decision and recommended next action.
        "## TPM decision / next step",
        clean(project.get("TPM decision / next step")),
        "",

        # Project-specific risks and issues.
        "## Risks and issues"
    ]

    # Handle projects that do not have any recorded risks.
    if not related:
        lines.append(
            "- No risks or issues are listed for this project."
        )

    # Add each associated risk with its important details.
    for r in related:

        lines.append(
            f"- **{clean(r.get('ID'))}: "
            f"{clean(r.get('Description'))}** "

            # Include probability, impact, owner, due date and status.
            f"(Probability: {clean(r.get('Probability'))}; "
            f"Impact: {clean(r.get('Impact'))}; "
            f"Owner: {clean(r.get('Owner'))}; "
            f"Due: {clean(r.get('Due date'))}; "
            f"Status: {clean(r.get('Status'))}). "

            # Include the mitigation or corrective action.
            f"Action: {clean(r.get('Action / mitigation'))}"
        )

    # Add a final review note to remind the TPM to validate
    # the information before sharing the report externally.
    lines += [
        "",
        "## Review note",
        (
            "This report reflects tracker contents. "
            "The TPM must validate status, dates and owners "
            "before sharing externally."
        ),
        ""
    ]

    # Join all report lines using actual newline characters.
    # This ensures the generated Markdown file is properly formatted.
    return "\n".join(lines)


def main():
    """
    Main execution function.

    Responsibilities:
    1. Accept command-line arguments.
    2. Load Excel project and risk information.
    3. Validate that project records exist.
    4. Create the output directory.
    5. Generate and save one report per project.
    """

    # Create the command-line argument parser.
    parser = argparse.ArgumentParser(
        description="Generate weekly TPM reports from Excel."
    )

    # Required argument: Excel workbook path.
    parser.add_argument(
        "workbook",
        help="Path to T5_Project_Tracker.xlsx"
    )

    # Optional argument: Output directory.
    # If the user does not specify it, reports are saved
    # inside a folder named 'reports'.
    parser.add_argument(
        "--out",
        default="reports",
        help="Output folder"
    )

    # Read the arguments entered by the user.
    args = parser.parse_args()

    # Load project summaries and risk records from Excel.
    projects, risks = load_data(args.workbook)

    # Stop execution if the Summary worksheet has no project records.
    if not projects:
        raise SystemExit(
            "No project rows found in Summary."
        )

    # Convert the output directory into a Path object.
    out = Path(args.out)

    # Create the output directory if it does not already exist.
    # parents=True allows parent directories to be created.
    # exist_ok=True prevents an error if the folder already exists.
    out.mkdir(parents=True, exist_ok=True)

    # Process each project independently.
    for p in projects:

        # Create a filename-friendly version of the project name.
        # Example: "Sagar Publications" is not used here;
        # "Ledger" becomes "ledger".
        slug = (
            clean(p.get("Project"), "project")
            .lower()
            .replace(" ", "_")
        )

        # Create the output filename for this project.
        target = out / f"{slug}_weekly_status.md"

        # Generate the report and save it as a UTF-8 Markdown file.
        target.write_text(
            render(p, risks),
            encoding="utf-8"
        )

        # Display the generated file location in the terminal.
        print(f"Created: {target}")

    # Display a final summary after all reports are generated.
    print(
        f"Done. Generated {len(projects)} report(s)."
    )


# This condition ensures main() runs only when this script
# is executed directly, not when it is imported into another Python file.
if __name__ == "__main__":
    main()
