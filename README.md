# Configuration Manager Audit Report Generator

The Configuration Manager Audit Report Generator is a PowerShell script that automates the generation of an audit report for Microsoft Configuration Manager environments. The script gathers important audit data and generates a comprehensive report in Word format.

## Features

- Retrieves information about collections, applications, packages, deployments, task sequences, site details, site servers, and SQL server details.
- Formats the report with headers, bold text, indents, and color-coding for task sequence steps.
- Supports customization for different report formats and outputs.
- Easy to use and can be scheduled or run on-demand.

## Prerequisites

- Microsoft Configuration Manager PowerShell module.
- Configuration Manager console installed on the system.
- Microsoft Word application installed.

## Usage

1. Clone or download the repository to your local machine.
2. Open PowerShell or a PowerShell Integrated Scripting Environment (ISE).
3. Run the `Get-CMAuditReport.ps1` script.
4. The script will generate a Word document containing the audit report.

## Documentation

For more detailed information on how to use the script, including parameter options and additional notes, please refer to the [script documentation](./Documentation.md).

## License

This project is licensed under the [MIT License](./LICENSE).
