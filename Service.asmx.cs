using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Services;
using Microsoft.Exchange.WebServices.Data;
using Microsoft.Exchange.WebServices.Autodiscover;
using System.Xml;


namespace CMCWebService
{
    /// <summary>
    /// Summary description for Calendar Manipulation Web Service:
    ///     New - Creates a new Entry on the Calendar
    ///     Update - Updates an existing Entry on the Calendar
    ///     Delete - Deletes an existing Entry on the Calendar
    /// </summary>
    /// 
    [WebService(Namespace = "http://yourdomain.com/namespace")]
    [WebServiceBinding(ConformsTo = WsiProfiles.BasicProfile1_1)]
    [System.ComponentModel.ToolboxItem(false)]
    // To allow this Web Service to be called from script, using ASP.NET AJAX, uncomment the following line. 
    // [System.Web.Script.Services.ScriptService]

        
    public class Service : System.Web.Services.WebService
    {

        /**
         * Function GetBinding
         * Creates the connection to Office 365
         **/
        static ExchangeService GetBinding()
        {
            Console.WriteLine("Creating Binding");

            // Create a New Binding
            ExchangeService service = new ExchangeService(ExchangeVersion.Exchange2010_SP1);

            // Define Service Credentials (Either AppImp or User's)
            service.Credentials = new WebCredentials("username", "password");

            // Turn on tracing for debugging
            service.TraceListener = new CMCWebService.TraceListener();
            service.TraceFlags = TraceFlags.All;
            service.TraceEnabled = true;
            
            // Use the AutodiscoverUrl method to locate the service endpoint for a User
            try
            {

                // Tries to find the AutoDiscoverURL and compares it to your ValidationCallback function (below)
                // If it does not validate, it throws an exception
                service.AutodiscoverUrl("username", RedirectionUrlValidationCallback);
            }
            catch (AutodiscoverRemoteException ex)
            {
                Console.WriteLine("Exception thrown: " + ex.Error.Message);
            }

            // Display the service URL
            Console.WriteLine("AutodiscoverURL: " + service.Url);
            return service;
        }

        /**
         * Function: RedirectionUrlValidationCallback
         * Validates that the returned AutoDiscover URL is indeeded a Microsoft Online URL
         **/
        static bool RedirectionUrlValidationCallback(string redirectionUrl)
        {
            // Perform validation.
            // Validation is developer dependent to ensure a safe redirect.
            return (redirectionUrl == "https://autodiscover-s.outlook.com/autodiscover/autodiscover.xml");
        }


        [WebMethod]
        public string NewApp(string subject, string body, string startDate, int duration, string status)
        {
            

            try
            {
                // Reformat date
                DateTime date = DateTime.Parse(startDate);

                // Create the Service Binding
                ExchangeService service = GetBinding();

                // Runs the method to create the new appointment
                string newAppId = CreateNewAppointment(service, subject, body, date, duration, status);

                // Returns the ItemId for future changes.
                return newAppId;
            }
            catch (HttpException h)
            {

                return "Caught HttpException: " + h;

            }
            catch (Exception e) {

                return "Generic Exception: " + e;

            }
            
        }

        /**
         * Function: CreateNewAppointment
         * Creates a new appointment.
         * Returns the AppID (appointment id) for future changes
         **/
        static string CreateNewAppointment(ExchangeService service, string subject, string body, DateTime date, int duration, string status)
        {
            Console.WriteLine("Creating a New Item: " + subject);

            // Create an appointmet and identify the Exchange service.
            Appointment appointment = new Appointment(service);

            // Set details
            appointment.Subject = subject;
            appointment.Body = body;
            appointment.Start = date;
            appointment.End = appointment.Start.AddHours(duration);
            StringList categories = new Microsoft.Exchange.WebServices.Data.StringList();
            categories.Add(status);
            appointment.Categories = categories;
            appointment.IsReminderSet = false;

            // Save and Send
            appointment.Save(SendInvitationsMode.SendToNone);

            // Get ItemId for future updates
            return ""+appointment.Id+"";

        }


        [WebMethod]
        public string UpdateApp(string appId, string subject, string body, string startDate, string duration, string status)
        {
            // Reformat date
            // Sets a default time
            DateTime date = DateTime.Parse("01/01/12 00:00:00");
            try
            {
                // Tries to set the time, succeeds if not null
                date = DateTime.Parse(startDate);
            }
            catch (FormatException)
            {
                Console.WriteLine("Date not passed in.");
            }
            finally
            {
                // Sets an invalid default duration
                int dur = -1;
                try
                {
                    // Reformat Duration, if exists
                    dur = int.Parse(duration);
                }
                catch (FormatException)
                {
                    Console.WriteLine("Duration not passed in.");
                }
                finally
                {
                    // Create the Service Binding
                    ExchangeService service = GetBinding();

                    // Runs the method to update an existing appointment
                    UpdateAppointment(service, appId, subject, body, date, dur, status);
                }
            }

            return "Success";
        }

        /**
         * Function: UpdateAppointment
         * Updates an existing appointment 
         **/
        static void UpdateAppointment(ExchangeService service, string appId, string subject, string body, DateTime date, int duration, string status)
        {
            Console.WriteLine("Updating Appointment: " + appId);

            // Bind existing appointment to an identifier
            Appointment appointment = Appointment.Bind(service, new ItemId(appId));

            // Make changes depending on the validity of the inputs
            if ((subject != null)&&(subject != ""))
            {
                appointment.Subject = subject;
            }
            if ((body != null)&&(body != ""))
            {
                appointment.Body = body;
            }
            if (date != DateTime.Parse("01/01/12 00:00:00"))
            {
                // Gets the initial duration to perserve it, if duration also not changed
                TimeSpan initialDuration = appointment.Duration;
                appointment.Start = date;
                appointment.End = appointment.Start.Add(initialDuration);
                
            }
            if (duration > 0)
            {
                appointment.End = appointment.Start.AddHours(duration);
            }
            if ((status != null)&&(status != ""))
            {
                StringList categories = new Microsoft.Exchange.WebServices.Data.StringList();
                categories.Add(status);
                appointment.Categories = categories;
            }

            // Save changes
            appointment.Update(ConflictResolutionMode.AlwaysOverwrite);
        }

        [WebMethod]
        public string DeleteApp(string appId)
        {
            // Create the Service Binding
            ExchangeService service = GetBinding();

            DeleteAppointment(service, appId);

            return "Success";
        }

        /**
         * Function: DeleteAppointment
         * Removes the indicated appointment
         **/
        static void DeleteAppointment(ExchangeService service, string appId)
        {
            // Bind existing appointment to an identifier
            Appointment appointment = Appointment.Bind(service, new ItemId(appId));

            // Delete the appointment
            appointment.Delete(DeleteMode.MoveToDeletedItems);
        }

    }
}
