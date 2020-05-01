Quick tutorial on migrating a Logic Apps for our Site provisioning.  

Video link is https://envisionit.sharepoint.com/sites/marketing/_layouts/15/guestaccess.aspx?docid=0dd8063780fa2450c9cdc3c56f7eaee61&authkey=AaRhla19F6X8c_siIizVwwQ&expiration=2020-03-26T14%3A41%3A53.000Z&e=XfnI4o.

Basic steps are:

•	Create a stub workflow in the target environment with each of the connections needed
•	Save the JSON
•	Copy the source JSON and replace the bottom parameters section with the connection parameters from the stub workflow
•	Paste the JSON into the target
•	Fix up the steps, either through search and replace in the JSON or through the designer, to connect to the proper resources (like the right SharePoint site or Azure Automation)
