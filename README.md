# Report
This repository contains JS-scripts, executed by WScript in OS Windows.
Main JS-script is used for execution different types of reports, based on SQL-scripts, and contains the following steps:
- Reading configuration file, contained path to SQL-script and parameters, correspond to report to be executed.
- Executing SQL-script with parameters on configured Oracle DB using ADODB Connection (Oracle Client needs to be installed).

- Final result is transformed from ADO DB XML to Microsoft Excel (Table XML 2003) format via XSLT transform.
