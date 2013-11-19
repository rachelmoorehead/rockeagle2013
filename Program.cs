using System;
using System.Threading;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Exchange.WebServices.Data;
using Microsoft.Exchange.WebServices.Autodiscover;
using System.IO;
using System.Collections.ObjectModel;
using System.Management.Automation;
using System.Management.Automation.Runspaces;


namespace HolidayInsertion
{
    class Program
    {
       
        /**
         * Function loadHolidays
         * Creates a List of String[name, date] of holidays from CSV
         **/
        static List<string[]> LoadHolidays(string filename, ExchangeService service)
        {
            // Initial local vairables
            string line;
            List<string[]> holidays = new List<string[]>();

            using (StreamReader file = new StreamReader(filename))
            {
                // While there are more lines
                while ((line = file.ReadLine()) != null)
                {
                    // And they contain data
                    if (line.Trim().Length > 0)
                    {
                        // Split into two cells on the comma
                        char[] delimiter = new char[] { ';' };
                        string[] cells = line.Split(delimiter);

                        // Add cells to array
                        holidays.Add(cells);
                    }
                }

                // Return the List of Holidays
                Console.WriteLine("Returning List of " + holidays.Count + " holidays");
                return holidays;

               
            }
        }

        /**
        * Function GetMailboxes
        * Creates a CSVs of mailboxes by beginning character
        **/
        static void GetMailboxes(string username, string password) {

            // Run PowerShell script to generate mailbox list
            RunspaceConfiguration runspaceConfiguration = RunspaceConfiguration.Create();
            Runspace runspace = RunspaceFactory.CreateRunspace(runspaceConfiguration);
            runspace.Open();

            RunspaceInvoke scriptInvoker = new RunspaceInvoke(runspace);

            runspace.SessionStateProxy.SetVariable("uname", username);
            runspace.SessionStateProxy.SetVariable("upass", password);

            Pipeline pipeline = runspace.CreatePipeline();

            string scriptfile = "../../mailboxes.ps1";
            Command command = new Command(scriptfile);
            pipeline.Commands.Add(command);
            pipeline.Invoke();
        
        }

        /**
        * Function LoadMailboxes
        * Creates a List of Strings of mailboxes from CSV
        **/
        static List<string> LoadMailboxes(string c)
        {
            
            // Load mailbox list into program
            string filename = "../../../mailboxes/" + c +"_mailboxes.csv";
            using (StreamReader file = new StreamReader(filename))
            {
                // Initial local vairables
                string line;
                List<string> mailboxes = new List<string>();

                // Ignore header information
                file.ReadLine();

                // While there are more lines
                while ((line = file.ReadLine()) != null)
                {
                    // And they contain data
                    if (line.Trim().Length > 0)
                    {

                        // Add items to list
                        mailboxes.Add(line);

                    }
                }

                // Return the List of Holidays
                return mailboxes;

            }
        }



        /**
         * Function GetBinding
         * Creates the connection to Office 365
         **/
        static ExchangeService GetBinding(string username, string password)
        {
            // Create a New Binding
            ExchangeService service = new ExchangeService(ExchangeVersion.Exchange2010_SP1);

            // Define Service Credentials with AppImpersonation and View Only Org Admin
            service.Credentials = new WebCredentials(username, password);

            // Use the AutodiscoverUrl method to locate the service endpoint for a User
            try
            {

                // Tries to find the AutoDiscoverURL and compares it to your ValidationCallback function (below)
                // If it does not validate, it throws an exception
                service.AutodiscoverUrl(username, RedirectionUrlValidationCallback);
            }
            catch (AutodiscoverRemoteException ex)
            {
                Console.WriteLine("Exception thrown: " + ex.Error.Message);
            }

            // Display the service URL
            return service;
        }

        /**
         * Function: RedirectionUrlValidationCallback
         * Validates that the returned AutoDiscover URL is indeeded a Microsoft Online URL
         **/
        static bool RedirectionUrlValidationCallback(string redirectionUrl)
        {
            // Perform validation.
            return (redirectionUrl == "https://autodiscover-s.outlook.com/autodiscover/autodiscover.xml");
        }

        /**
         * Function Add the Holidays
         **/
        static void AddHolidays(ExchangeService service, List<string[]> holidays, List<string> mailboxes) {

            // Log file
            string datetimeString = DateTime.Now.ToString("MMddyyyy");
            string logfile = "../../logs/" + datetimeString + "_add_holiday_log.txt";

            //Initiate Error List
            List<string> mbs = new List<string>();

            using (System.IO.StreamWriter log = new System.IO.StreamWriter(@logfile, true))
            {

                // Loop through each email address in the passed in mailboxes List
                foreach (string mailbox in mailboxes)
                {

                    // Impersonate that User
                    service.ImpersonatedUserId = new ImpersonatedUserId(ConnectingIdType.SmtpAddress, mailbox);
                    Console.WriteLine("Attempting to Add Holidays to: " + mailbox);

                    List<Appointment> uga_holidays = new List<Appointment>();

                    // Loop through all the holidays
                    foreach (string[] holiday in holidays)
                    {

                        //Create a new appointment
                        Appointment appointment = new Appointment(service);

                        // Set details
                        appointment.Subject = holiday[0];
                        appointment.Start = DateTime.Parse(holiday[1]);
                        appointment.End = appointment.Start.AddDays(1);
                        appointment.IsAllDayEvent = true;
                        StringList categories = new Microsoft.Exchange.WebServices.Data.StringList();
                        categories.Add("Holiday");
                        appointment.Categories = categories;
                        appointment.IsReminderSet = false;

                        uga_holidays.Add(appointment);

                    }
             
                    // Save and Send
                    try
                    {
                        service.CreateItems(uga_holidays, WellKnownFolderName.Calendar, MessageDisposition.SaveOnly, SendInvitationsMode.SendToNone);
                        Console.WriteLine("Added Holiday Successfully to: " + mailbox);

                        DateTime now = DateTime.Now;
                        log.WriteLine(now + " - Added holidays succesfully to Mailbox: " + mailbox);


                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine("Error During Initial Add - Mailbox: " + mailbox + "; Exception thrown: " + ex);
                        log.WriteLine("Error During Initial Add - Mailbox: " + mailbox + "; Exception thrown: " + ex);
                        mbs.Add(mailbox);
                       
                    }

                    // Clear impersonation.
                    service.ImpersonatedUserId = null;

                }

                //Process Rerun List
                if (mbs.Count > 0)
                {
                    Console.WriteLine("Looping through re-run mailboxes.");
                    
                    while (mbs.Count > 0)
                    {
                        // Current mailbox
                        string mb = mbs.ElementAt(0);
                        Console.WriteLine("On Mailbox: " + mb);

                        // Take the mailbox out of the first element slot
                        log.WriteLine("Removing mailbox " + mb + " from beginning of mbs.");
                        mbs.RemoveAt(0);
                        mbs.TrimExcess();

                        try
                        {
                            // Reruns: Removes
                            // Run search
                            // Impersonate that User
                            service.ImpersonatedUserId = new ImpersonatedUserId(ConnectingIdType.SmtpAddress, mb);

                            // Search String
                            String category = "Holiday";
                            // Search Filter
                            SearchFilter.IsEqualTo filter = new SearchFilter.IsEqualTo(AppointmentSchema.Categories, category);

                            // Result Return Size, number of items
                            ItemView holidayView = new ItemView(100);
                            // Limit data to only necesary components
                            holidayView.PropertySet = new PropertySet(BasePropertySet.IdOnly, ItemSchema.Categories);

                            FindItemsResults<Item> items = service.FindItems(WellKnownFolderName.Calendar, filter, holidayView);

                            if (items.TotalCount > 0)
                            {

                                Console.WriteLine("Removing " + items.TotalCount + " holidays from " + mb);
                                log.WriteLine("Found " + items.TotalCount + " holidays in the Calendar folder for " + mb + " to be removed.");

                                List<ItemId> ids = new List<ItemId>();
                                foreach (Item item in items)
                                {

                                    ids.Add(item.Id);

                                }

                                service.DeleteItems(ids, DeleteMode.MoveToDeletedItems, null, null);

                            }
                            else 
                            {
                                log.WriteLine("Found no holidays in the Calendar folder for " + mb + " to be removed.");
                            }
                            

                            // Rerun: Adds
                           
                            List<Appointment> holidays = new List<Appointment>();

                            // Loop through all the holidays
                            foreach (string[] holiday in holidays)
                            {

                                //Create a new appointment
                                Appointment appointment = new Appointment(service);

                                // Set details
                                appointment.Subject = holiday[0];
                                appointment.Start = DateTime.Parse(holiday[1]);
                                appointment.End = appointment.Start.AddDays(1);
                                appointment.IsAllDayEvent = true;
                                StringList categories = new Microsoft.Exchange.WebServices.Data.StringList();
                                categories.Add("Holiday");
                                appointment.Categories = categories;
                                appointment.IsReminderSet = false;

                                holidays.Add(appointment);

                            }

                            service.CreateItems(holidays, null, null, SendInvitationsMode.SendToNone);
                            Console.WriteLine("Added Holiday Successfully to" + mb);
                            DateTime now = DateTime.Now;
                            log.WriteLine(now + " - Added holidays succesfully to Mailbox: " + mb);

                        }
                        catch
                        {
                            log.WriteLine("Fatal Mailbox Errored on Re-Run Removes: " + mb + "; Will not retry.");
                            Console.WriteLine("Fatal Mailbox Errored on Re-Run Removes: " + mb + "; Will not retry.");
                        }

                        // Clear impersonation.
                        service.ImpersonatedUserId = null;

                        
                    }
                }
            }
        
        }

        /**
         * Function Undo Holiday Insertion
         */
        static void UndoInsertion(ExchangeService service, List<string> mailboxes)
        {

            // Log file
            string datetimeString = DateTime.Now.ToString("MMddyyyy");
            string logfile = "../../logs/" + datetimeString + "_undo_holiday_log.txt";

            using (System.IO.StreamWriter log = new System.IO.StreamWriter(@logfile, true))
            {

                // Mailboxes that need to be rerun that errored during this process
                List<string> rrmailboxes = new List<string>();

                foreach (string mailbox in mailboxes)
                {

                    // Find the holidays
                    try
                    {
                        // Run search
                        // Impersonate that User
                        service.ImpersonatedUserId = new ImpersonatedUserId(ConnectingIdType.SmtpAddress, mailbox);

                        // Search String
                        String category = "Holiday";
                        // Search Filter
                        SearchFilter.IsEqualTo filter = new SearchFilter.IsEqualTo(AppointmentSchema.Categories, category);

                        // Result Return Size, number of items
                        ItemView holidayView = new ItemView(500);
                        // Limit data to only necesary components
                        holidayView.PropertySet = new PropertySet(BasePropertySet.IdOnly, ItemSchema.Categories);                 
                       
                        FindItemsResults<Item> items = service.FindItems(WellKnownFolderName.Calendar, filter, holidayView);

                        if (items.TotalCount > 0)
                        {
                            List<ItemId> ids = new List<ItemId>();
                            foreach (Item item in items)
                            {

                                ids.Add(item.Id);

                            }

                            service.DeleteItems(ids, DeleteMode.MoveToDeletedItems, SendCancellationsMode.SendToNone, AffectedTaskOccurrence.AllOccurrences);
                            Console.WriteLine("Removed " + items.TotalCount + " holidays in the Calendar folder for " + mailbox);
                            DateTime now = DateTime.Now;
                            log.WriteLine(now + " - Removed " + items.TotalCount + " holidays in the Calendar folder for " + mailbox);

                        }
                        else
                        {
                            Console.WriteLine("Could not find any holidays for mailbox: " + mailbox);
                        }
                    }
                    catch
                    {
                        
                            log.WriteLine("Mailbox Errored: " + mailbox);
                            rrmailboxes.Add(mailbox);
                    
                    }

                    // Clear impersonation.
                    service.ImpersonatedUserId = null;
                }

                // Rerun errored accounts.

                if (rrmailboxes.Count > 0)
                {
                    Console.WriteLine("Looping through errored mailboxes.");
                }
                while (rrmailboxes.Count > 0)
                {

                    // Run search

                    // Current mailbox
                    string mb = rrmailboxes.ElementAt(0);
                    Console.WriteLine("On Mailbox: " + mb);
                    // Take the mailbox out of the first element slot
                    Console.WriteLine("Removing mailbox " + mb + " from beginning of rrmailboxes.");
                    rrmailboxes.RemoveAt(0);
                    rrmailboxes.TrimExcess();

                    // Find the holidays
                    try
                    {

                        // Impersonate that User
                        service.ImpersonatedUserId = new ImpersonatedUserId(ConnectingIdType.SmtpAddress, mb);

                        // Search String
                        String category = "Holiday";
                        // Search Filter
                        SearchFilter.IsEqualTo filter = new SearchFilter.IsEqualTo(AppointmentSchema.Categories, category);

                        // Result Return Size, number of items
                        ItemView holidayView = new ItemView(100);
                        // Limit data to only necesary components
                        holidayView.PropertySet = new PropertySet(BasePropertySet.IdOnly, ItemSchema.Categories);


                        FindItemsResults<Item> items = service.FindItems(WellKnownFolderName.Calendar, filter, holidayView);

                        if (items.TotalCount > 0)
                        {
                            //Log number of found items
                            Console.WriteLine("Found " + items.TotalCount + " holidays in the Calendar folder for " + mb);

                            List<ItemId> ids = new List<ItemId>();
                            foreach (Item item in items)
                            {

                                ids.Add(item.Id);
                                Console.WriteLine(mb + ": Added ItemID to be removed: " + item.Id);
                            }

                            service.DeleteItems(ids, DeleteMode.MoveToDeletedItems, null, null);
                            Console.WriteLine("Removed " + items.TotalCount + " holidays in the Calendar folder for " + mb);
                            DateTime now = DateTime.Now;
                            log.WriteLine(now + " - Removed " + items.TotalCount + " holidays in the Calendar folder for " + mb);
                            
                        }
                        else
                        {

                            Console.WriteLine("Could not find any holidays for mailbox: " + mb);
                            log.WriteLine("Could not find any holidays for mailbox: " + mb);
                        }
                    }
                    catch
                    {
                        DateTime now = DateTime.Now;
                        log.WriteLine(now + " - Fatal Mailbox Errored: " + mb + "; Will not retry");
                        Console.WriteLine("Fatal Mailbox Errored: " + mb + "; Will not retry");


                    }

                    // Clear impersonation.
                    service.ImpersonatedUserId = null;


                }
            }

        }
 

        static void Main(string[] args)
        {
            
           /**
            * Usage: HolidayInsertion <type> <csv>
            *   where:
            *       <type> can be full_insert, full_remove, list_insert, list_remove
            *       <csv> should be the location of the list to be added or removed
            *             must not have headers and be a single column of userprincipalnames (email@address.com)
            *             must be the full path and wrapped in quotes if there are spaces
            */

            // Test Inputs

            if (args.Length < 3) {

                Console.WriteLine("Please provide appropriate parameters.");
                Console.WriteLine("Expected Usage:  HolidayInsertion <type> <username> <password> <csv> (optional)");
                Console.WriteLine("   Where <type> can be 'full_insert', 'full_remove', 'list_insert', 'list_remove' ");
                Console.WriteLine("         <csv> is the location of a single column of userprincipalnames (email@address.com) (optional)");
                Console.WriteLine("Exit Code 0: No arguments");
                Console.WriteLine("Press ENTER to exit.");
                Console.ReadLine();
                Environment.Exit(0); // Exit Code 0 = No arguments
            
            }

            else if (args.Length == 3) {

                if (args[0].Equals("full_insert")) {

                    // Sanity Check
                    Console.WriteLine("Preparing to Insert Holidays into All Calendars.  Press ENTER to continue.");
                    Console.ReadLine();

                    // Make Connection
                    ExchangeService service = GetBinding(args[1], args[2]);
                    Console.WriteLine("Made Binding.");

                    // Load Holidays
                    string path = "../../holidays.txt";
                    List<string[]> holidays = LoadHolidays(path, service);
                    Console.WriteLine("Done Loading Holidays.");

                    // Get Mailboxes
                    GetMailboxes(args[1], args[2]);
                    Console.WriteLine("Done Getting Mailboxes.");

                    //Potential first characters
                    string[] set  = new string[36]{"0","1","2","3","4","5","6","7","8","9","a","b","c","d","e","f","g","h","i","j","k","l","m","n","o","p","q","r","s","t","u","v","w","x","y","z"}; 

                    // For each new character start a thread
                    foreach (string c in set)
                    {

                        // Load Mailboxes -- Returns mailboxes specific to character set
                        List<string> mailboxes = LoadMailboxes(c);

                        if (mailboxes.Count > 0)
                        {

                            Console.WriteLine("Loaded Mailboxes for Character: " + c);

                            // Add Holidays
                            AddHolidays(service, holidays, mailboxes);
                            Console.WriteLine("Added Holidays for Character: " + c);
                        }
                        else {
                            Console.WriteLine("No mailboxes found for character '"+c+"'");
                        }
                         
                    }
                
                }

                else if (args[0].Equals("full_remove"))
                {

                    // Sanity Check
                    Console.WriteLine("Preparing to Remove Holidays from All Calendars.  Press ENTER to continue.");
                    Console.ReadLine();

                    // Make Connection
                    ExchangeService service = GetBinding(args[1], args[2]);

                    // Undo Insertion
                    //Potential first characters
                    string[] set  = new string[36]{"0","1","2","3","4","5","6","7","8","9","a","b","c","d","e","f","g","h","i","j","k","l","m","n","o","p","q","r","s","t","u","v","w","x","y","z"}; 

                    // For each new character start a thread
                    foreach (string c in set)
                    {
                        List<string> mailboxes = LoadMailboxes(c);

                        if (mailboxes.Count > 0)
                        {
                            UndoInsertion(service, mailboxes);
                            Console.WriteLine("Removed Holidays for Character: " + c + "'");
                        }
                        else
                        {
                            Console.WriteLine("No mailboxes to remove for character '" + c + "'");
                        }
                    }

                }

                else {

                    Console.WriteLine("There was an error with your provided input.  Please provide the appropriate parameters.");
                    Console.WriteLine("Expected Usage:  HolidayInsertion <type> <username> <password> <csv> (optional)");
                    Console.WriteLine("   Where <type> can be 'full_insert', 'full_remove', 'list_insert', 'list_remove' ");
                    Console.WriteLine("         <csv> is the location of a single column of userprincipalnames (email@address.com) (optional)");
                    Console.WriteLine("Exit Code 1: Three arguments provided but not 'full_remove' or 'full_insert'.  Check your case and spelling.");
                    Console.WriteLine("Press ENTER to exit.");
                    Console.ReadLine();
                    Environment.Exit(1); // Exit Code 1 = Three arguments, but not one recognized (check case)
                }
            
            }
            else if (args.Length == 4) {

                //Initialize
                List<string> mailboxes = new List<string>();
                string curFile = @args[3];

                //Check file path 
                if (!(File.Exists(curFile)))
                {
                    Console.WriteLine("There was an error with your provided csv input.  Please provide the appropriate parameters.");
                    Console.WriteLine("Expected Usage:  HolidayInsertion <type> <username> <password> <csv> (optional)");
                    Console.WriteLine("   Where <type> can be 'full_insert', 'full_remove', 'list_insert', 'list_remove' ");
                    Console.WriteLine("         <csv> is the location of a single column of userprincipalnames (email@address.com) (optional)");
                    Console.WriteLine("Exit Code 2: File does not exist or this program does not have access to read/write to it.");
                    Console.WriteLine("Press ENTER to exit.");
                    Console.ReadLine();
                    Environment.Exit(2); // Exit Code 2 = File does not exist or this program does not have access to read/write to it.
                }

                // Load Mailboxes
                else
                {
                    try
                    {
                        String filename = args[3];

                        using (StreamReader file = new StreamReader(filename))
                        {
                            // Initial local vairables
                            string line;
                            

                            // While there are more lines
                            while ((line = file.ReadLine()) != null)
                            {
                                // And they contain data
                                if (line.Trim().Length > 0)
                                {
                                    // Add cells to array
                                    mailboxes.Add(line);

                                }
                            }

                            // Return the List of Holidays
                            Console.WriteLine("Returning List of " + mailboxes.Count + " mailboxes");

                        }
                    }
                    catch
                    {
                        
                        Console.WriteLine("There was an error with your provided csv input.  Please provide the appropriate parameters in the appropriate format.");
                        Console.WriteLine("Expected Usage:  HolidayInsertion <type> <username> <password> <csv> (optional)");
                        Console.WriteLine("   Where <type> can be 'full_insert', 'full_remove', 'list_insert', 'list_remove' ");
                        Console.WriteLine("         <csv> is the location of a single column of userprincipalnames (email@address.com) (optional)");
                        Console.WriteLine("Exit Code 3: There was an error reading your provided csv file.");
                        Console.WriteLine("Press ENTER to exit.");
                        Console.ReadLine();
                        Environment.Exit(3);
                    }
                }

                // Checking for Action argument
                if (args[0].Equals("list_insert"))
                {

                    // Sanity Check
                    Console.WriteLine("Preparing to INSERT Holidays into the Calendars from the List provided.");
                    
                    // Make Connection
                    ExchangeService service = GetBinding(args[1], args[2]);
                    
                    // Load Holidays
                    string path = "../../holidays.txt";
                    List<string[]> holidays = LoadHolidays(path, service);

                   // Add Holidays
                    AddHolidays(service, holidays, mailboxes);

                    // Close console.
                    Console.WriteLine("All Done! Please review the error logs.");
                }

                else if (args[0].Equals("list_remove"))
                {

                    // Sanity Check
                    Console.WriteLine("Preparing to REMOVE Holidays from the Calendars from the List provided.");

                    // Make Connection
                    ExchangeService service = GetBinding(args[1], args[2]);

                    // Undo Insertion 
                    UndoInsertion(service, mailboxes);

                    // Close console.
                    Console.WriteLine("All Done!  Please review the error logs.");

                }
                
                // If no action argument, exit
                else
                {

                    Console.WriteLine("There was an error with your provided input.  Please provide the appropriate parameters.");
                    Console.WriteLine("Expected Usage:  HolidayInsertion <type> <username> <password> <csv> (optional)");
                    Console.WriteLine("   Where <type> can be 'full_insert', 'full_remove', 'list_insert', 'list_remove' ");
                    Console.WriteLine("         <csv> is the location of a single column of userprincipalnames (email@address.com) (optional)");
                    Console.WriteLine("Exit Code 4: Action argument not 'list_insert' or 'list_remove' with two provided parameters.  Please provide the correct input parameters.");
                    Console.WriteLine("Press ENTER to exit.");
                    Console.ReadLine();
                    Environment.Exit(4);
                }  
            }
        }
    }
}
