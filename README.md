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

## Contributions
Contributions are welcome. Please open an issue or submit a pull request.

## GNU General Public License
This script is free software: you can redistribute it and/or modify it under the terms of the GNU General Public License as published by the Free Software Foundation, either version 3 of the License, or (at your option) any later version.

This script is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU General Public License for more details.
