# vb-wmi-server-health-report
A VB script that collects server information over WMI and assembles it in a Word document.
There is an example config file include: *config.xml*
This VB program requires Microsoft Office to funtion, as it fills out an Office document with a report of servers.

**How to use:**

All the servers listed in the config file will be checked according to their hardware and software health.
For every server checked, the data will be appended to the word document.

This program can be useful for internal reports or if you have clients where you have to do a health report of Windows Server systems.