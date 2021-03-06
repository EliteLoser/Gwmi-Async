# Gwmi-Async

Svendsen Tech's Gwmi-Async.ps1, available for download/clone here, is a "wrapper" around Get-WmiObject. It is designed to retrieve and collect data from a (potentially large) list of computers. You get a custom, very flexible XML file parser based on the schema used, that simplifies creating custom PowerShell objects or CSV data that you can process, from the XML that the aforementioned script creates.

Author: Joakim Borger Svendsen.

Online blog documentation here: https://www.powershelladmin.com/wiki/Get-wmiobject_wrapper - It's comprehensive to reproduce it here on GitHub.

Here's a screenshot of the version before the third major rewrite:

![image_gwmi_wrapper_async](/img/Gwmi-Async-Sample.png)

Example of the produced XML:

![image_gwmi_wrapper_async](img/Gwmi-Async-Xml-Sample.png)

