# How we receive and update information from PartnerCenter #

**PartnerCenter** **_(PC)_** is used by variety of _Reports_ sent us in Excel.xls files by Distributors and directly from the Portal as a weekly mail with the URL to updated report on **_PC_**. These files containes data:

  * Contract/Agreement Number, Subscription Start and End Dates, Status
  * End-Customer Number, Name, Address
  * Contract Manager **_(CM)_** or Software Coordinator: Name, Telephone, Mail
  * Subscription or Asset Number **_(SN)_**, Status, SKU/Description

From Distributor we also regularely (typically - Monthly) receive reports
  * **New Releases** - new **_SN_** for Autodesk goods uner Subscriptions genereted by Autodesk
  * **Renewal Reports** - reminders about necessity of end-customer Subsctiption prolongation

# Details #

All reports comes from _PC_ we put into file ADSK.xlsx.
Than we're doind **_Mapping_**, i.e. re-arrange the Column seequence and somtimes the content in according with the list of required fields listed above. This information is inserted as a line in **_Table-of-Content (TOC)_** list together with the Range, which pointed to the new handled Report.

Later on the particular report is selected and loaded into **match** application for further work on it.