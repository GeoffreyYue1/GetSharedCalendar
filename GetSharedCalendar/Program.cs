using Microsoft.Exchange.WebServices.Data;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;

namespace GetSharedCalendar
{
    class Program
    {
        static StreamWriter sw;
        static void Main(string[] args)
        {
            string mailbox, pwd, url, groupName,mailAddress;

            string[] configures = File.ReadAllLines("configs.txt");
            url = configures[0];
            mailbox = configures[1];
            pwd = configures[2];
            mailAddress = configures[3];
            groupName = configures[4];

            if (File.Exists("Calendar.txt"))
                File.Delete("Calendar.txt");
            sw = File.CreateText("Calendar.txt");

            ExchangeService service = new ExchangeService(ExchangeVersion.Exchange2010_SP1);

            service.Credentials = new WebCredentials(mailbox, pwd);

            service.TraceEnabled = true;
            service.TraceFlags = TraceFlags.All;

            service.Url = new Uri(url);

            ServicePointManager.ServerCertificateValidationCallback = (sender, certificate, chain, sslPolicyErrors) => true;


            WriteLog("Calendar of : " + mailbox);
                GetUserCalendar(service);
            WriteLog("");
            WriteLog("----------Shared Calendars----------");
            GetSharedCalendarFolders(service, mailAddress, groupName);


            sw.Close();

            Console.ReadKey();
            

        }


        private static void WriteLog( string log)
        {
            sw.WriteLine(log);

        }

        private static void GetUserCalendar(ExchangeService service)
        {
            /*
            // Initialize the calendar folder object with only the folder ID. 
            CalendarFolder calendar = CalendarFolder.Bind(service, WellKnownFolderName.Calendar, new PropertySet());


            ItemView itemview = new ItemView(1000);
            SearchFilter.SearchFilterCollection searchFilter = new SearchFilter.SearchFilterCollection();
            searchFilter.Add(new SearchFilter.IsGreaterThanOrEqualTo(AppointmentSchema.Start, DateTime.Today));
            searchFilter.Add(new SearchFilter.IsLessThanOrEqualTo(AppointmentSchema.Start, DateTime.Today.AddDays(1)));


            FindItemsResults<Item> items = calendar.FindItems(searchFilter, itemview);

            WriteLog("Total Item Number : " + items.Count());
            WriteLog("Start\t\tEnd\t\tSubject");
            foreach (Item it in items)
            {
                WriteLog((it as Appointment).Start.ToString() + "\t" + (it as Appointment).End + "\t" + (it as Appointment).Subject);
            }
            */
            PropertySet propSet = new PropertySet(AppointmentSchema.Subject,
                                     AppointmentSchema.Location,
                                     AppointmentSchema.Start,
                                     AppointmentSchema.End,
                                     AppointmentSchema.AppointmentType,
                                     AppointmentSchema.Sensitivity);

            CalendarView calView = new CalendarView(DateTime.Today, DateTime.Today.AddDays(1));
            calView.PropertySet = propSet;

            FindItemsResults<Appointment> results = service.FindAppointments(WellKnownFolderName.Calendar, calView);

            WriteLog("Total Item Number : " + results.Items.Count());
            WriteLog("Start\t\tEnd\t\tSubject\t\tSensitivity");

            foreach (Appointment appt in results.Items)
            {
                WriteLog(appt.Start.ToString() + "\t" + appt.End + "\t" + appt.Subject+"\t" + appt.Sensitivity.ToString());
            }

            
        }

        static Dictionary<string, Folder> GetSharedCalendarFolders(ExchangeService service, String mbMailboxname,string groupName)
        {
            Dictionary<String, Folder> rtList = new System.Collections.Generic.Dictionary<string, Folder>();

            FolderId rfRootFolderid = new FolderId(WellKnownFolderName.Root, mbMailboxname);
            FolderView fvFolderView = new FolderView(1000);
            SearchFilter sfSearchFilter = new SearchFilter.IsEqualTo(FolderSchema.DisplayName, "Common Views");

            FindFoldersResults ffoldres = service.FindFolders(rfRootFolderid, sfSearchFilter, fvFolderView);
            if (ffoldres.Folders.Count == 1)
            {

                PropertySet psPropset = new PropertySet(BasePropertySet.FirstClassProperties);
                ExtendedPropertyDefinition PidTagWlinkAddressBookEID = new ExtendedPropertyDefinition(0x6854, MapiPropertyType.Binary);
                ExtendedPropertyDefinition PidTagWlinkGroupName = new ExtendedPropertyDefinition(0x6851, MapiPropertyType.String);

                psPropset.Add(PidTagWlinkAddressBookEID);
                ItemView iv = new ItemView(1000);
                iv.PropertySet = psPropset;
                iv.Traversal = ItemTraversal.Associated;

                SearchFilter cntSearch = new SearchFilter.IsEqualTo(PidTagWlinkGroupName,groupName);
                // Can also find this using PidTagWlinkType = wblSharedFolder
                FindItemsResults<Item> fiResults = ffoldres.Folders[0].FindItems(cntSearch, iv);
                foreach (Item itItem in fiResults.Items)
                {
                    try
                    {
                        object GroupName = null;
                        object WlinkAddressBookEID = null;

                        // This property will only be there in Outlook 2010 and beyond
                        //https://msdn.microsoft.com/en-us/library/ee220131(v=exchg.80).aspx#Appendix_A_30
                        if (itItem.TryGetProperty(PidTagWlinkAddressBookEID, out WlinkAddressBookEID))
                        {

                            byte[] ssStoreID = (byte[])WlinkAddressBookEID;
                            int leLegDnStart = 0;
                            // Can also extract the DN by getting the 28th(or 30th?) byte to the second to last byte 
                            //https://msdn.microsoft.com/en-us/library/ee237564(v=exchg.80).aspx
                            //https://msdn.microsoft.com/en-us/library/hh354838(v=exchg.80).aspx
                            String lnLegDN = "";
                            for (int ssArraynum = (ssStoreID.Length - 2); ssArraynum != 0; ssArraynum--)
                            {
                                if (ssStoreID[ssArraynum] == 0)
                                {
                                    leLegDnStart = ssArraynum;
                                    lnLegDN = System.Text.ASCIIEncoding.ASCII.GetString(ssStoreID, leLegDnStart + 1, (ssStoreID.Length - (leLegDnStart + 2)));
                                    ssArraynum = 1;
                                }
                            }
                            NameResolutionCollection ncCol = service.ResolveName(lnLegDN, ResolveNameSearchLocation.DirectoryOnly, false);
                            if (ncCol.Count > 0)
                            {

                                FolderId SharedCalendarId = new FolderId(WellKnownFolderName.Calendar, ncCol[0].Mailbox.Address);
                                Folder SharedCalendaFolder = Folder.Bind(service, SharedCalendarId);

                                WriteLog("Calendar of : " + ncCol[0].Mailbox.Address);
                                WriteLog(ncCol[0].Mailbox.Address + " --- " + SharedCalendaFolder.DisplayName);

                                PropertySet propSet = new PropertySet(AppointmentSchema.Subject,
                                      AppointmentSchema.Location,
                                      AppointmentSchema.Start,
                                      AppointmentSchema.End,
                                      AppointmentSchema.AppointmentType,
                                      AppointmentSchema.Sensitivity);

                                CalendarView calView = new CalendarView(DateTime.Today, DateTime.Today.AddDays(1));
                                calView.PropertySet = propSet;

                                FindItemsResults<Appointment> results = service.FindAppointments(SharedCalendarId, calView);

                                WriteLog("Total Item Number : " + results.Items.Count());
                                WriteLog("Start\t\tEnd\t\tSubject\t\tSensitivity");

                                foreach (Appointment appt in results.Items)
                                {
                                    WriteLog(appt.Start.ToString() + "\t" + appt.End + "\t" + appt.Subject + "\t" + appt.Sensitivity.ToString());
                                }



                                /*
                               
                                ItemView itemview = new ItemView(1000);
                                SearchFilter.SearchFilterCollection searchFilter = new SearchFilter.SearchFilterCollection();
                                searchFilter.Add(new SearchFilter.IsGreaterThanOrEqualTo(AppointmentSchema.Start, DateTime.Today));
                                searchFilter.Add(new SearchFilter.IsLessThanOrEqualTo(AppointmentSchema.Start, DateTime.Today.AddDays(1)));


                                FindItemsResults<Item> items = SharedCalendaFolder.FindItems(searchFilter, itemview);

                                // FindItemsResults<Item> items = SharedCalendaFolder.FindItems(itemview);


                               

                              
                                
                                foreach(Item it in items)
                                {
                                    WriteLog((it as Appointment).Start.ToString() + "\t" + (it as Appointment).End +"\t" + (it as Appointment).Subject);
                                }
                                 */



                                rtList.Add(ncCol[0].Mailbox.Address, SharedCalendaFolder);


                            }

                        }
                    }
                    catch (Exception exception)
                    {
                        Console.WriteLine(exception.Message);
                       
                    }

                }
            }
            return rtList;
        }

    

    }




}

