# NxtReport

NxtReport is a simple spreadsheet based report design and creation package.

The initial intent of the project was to supplement the reporting capabilities of the NextJob Job Shop Control Software package.

NxtReport has three software components:

NxtDesigner - Program that creates the report design file and the report design data file. This program is written in FreePascal with the Lazarus IDE.

NxtReport  - Program that reads the report data file and creates the report data xml file.  This program resides on the QM server, and is written in QMBasic.

NxtReportGen - Program the reads the report data xml file and the report design file and produces the final spreadsheet report. This program is written in FreePascal with the Lazarus IDE.

The created report is spreadsheet based, NxtReport utilizes the  Free Pascal FPSpreadsheet package to design and created the report.
Information on FPSpreadsheet can be found at: https://wiki.lazarus.freepascal.org/FPSpreadsheet

This package  makes use of Delphi4QM -A Delphi and Free Pascal Wrapper for the QMClient C API:  https://github.com/FOSS4MV/delphi4qm

Information on QM can be found at:  https://www.openqm.com/

Lazarus IDE: https://www.lazarus-ide.org/

