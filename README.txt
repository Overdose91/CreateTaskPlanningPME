PlanningPME is a software developed by the company TargetSkills. 
This software is not free but the DLL is available free of charge on the website http://www.planningpme.com

With this DLL you can create insert controls (button) in the PlanningPME interface and associate actions or make automatic actions for example add a task 

Here is an example with Visual basic
http://www.planningpme.com/blog/post/Plug-in-creation.aspx

PlanningPME with C#

Example:

Connect an file excel (.XLS)
ppme.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=P:\Commercial\Clients\ALIANTIS\Script VBS\Test2\01052009 au 30062009.xls;Extended Properties=" & Chr(34) & "Excel 8.0;HDR=Yes" & Chr(34)
Connect database MS SQL Server
ppme.ConnectionString = "Provider=sqloledb;Data Source=servername;Initial Catalog=database;User Id=username;Password=password"

The names of interfaces with "s" to the end, are tables that you can read through with the boucle for
Example : Resources, Tasks, Departments...


Don't hesitate if you have questions...: Marc.absalon@free.fr

PS: Excuse me for my possible spelling errors...
