using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
//The DLL PlanningPME is available on http://www.planningpme.com
//using PlanningPME
using PlanningPME;

namespace ConsoleApplication19
{
    class Program
    {
        static void Main(string[] args)
        {
            /*
             Create one task in your software of plan PlanningPME
             More information in the file readme....
             */

            //Create Object PlanningPME
            IApplication ppme = new PlanningPME.Application();
            //Connection to database which will be modified
            //You can connect to the database with ....
            ppme.ConnectionString = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\ADMIN\Desktop\MairieParis.pp";
            ppme.Connect();
            //Create object Ressources
            //The object Ressources with "s" is a table
            Resources resources = new Resources();
            //Permit to get an table of all yours resources
            resources = ppme.GetResources();

            // Create an object DoTash which will allow you to create one task
            DoTask task = ppme.CreateItem(PpItemType.PpDoTask);
            //Add the resource who will do the task
            task.Resources.Add(resources.Item(2));
            
            task.BeginningDate = DateTime.ParseExact("06/03/2014", "dd/MM/yyyy", null);
            task.EndDate = DateTime.ParseExact("10/03/2014", "dd/MM/yyyy", null);

            task.BeginningTime = DateTime.ParseExact("08:00:00", "HH:mm:ss", null);
            task.EndTime = DateTime.ParseExact("23:59:59", "HH:mm:ss", null);

            task.Label = "Formation";

            //It exists two types of task, task and task unavailability
            //if unavailability = 1, it's a task unavailability
            task.Unavaibility = 1;
            //Don't forget TO save your task else it will not be inserted in your database
            task.Save();
            //Close your database's connection
            ppme.Close();
            Console.ReadKey();
        }
    }
}
