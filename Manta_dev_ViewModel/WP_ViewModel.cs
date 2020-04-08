using System;
using System.Collections.Generic;
using System.Linq;
using System.Collections.Specialized;
using Manta_dev_Data;
using System.Security.Principal;
using System.DirectoryServices.AccountManagement;

namespace Manta_dev_ViewModel
{
    public class WP_ViewModel
    {
        public List<Aktive_WP> DataCollection_Aktive_WP { get; set; }
        public List<Calendar> DataCollection_Calendar { get; set; }
        public List<Setting_Group> DataCollection_Settings_Group { get; set; }
        public List<Settings_Name_Columns> DataCollection_Settings_Name_Column { get; set; }

        public DateTime pom;
        public String SysFrom = "null";
        public String SysTo = "null";
        public String IntFrom = "null";
        public String IntTo = "null";
        public String RegFrom = "null";
        public String RegTo = "null";
        public String DeployProd = "null";
        public String Sys = "Systémově integrační testy"; //name atributte in database
        public String Integr = "Integrační testy"; //name atributte in database
        public String Reg = "Regresní testy"; //name atributte in database
        public String Deploy = "Instalace na produkční prostředí"; //name atributte in database
        public String Unknow = "neznámo"; //Unknow date
        public readonly string Username = WindowsIdentity.GetCurrent().Name.ToString();
        public bool IsEditor = false;
        public bool IsAdmin = false;
        public int[] Col1Fil;

        public readonly string All = "All";
        public readonly string DateFormat = "d. M. yyyy";
        public string Message { get; set; }
        public WP_ViewModel()
             : base()
        {
            Init();

        }
        public void Init()
        {
            Manta_dev_DataSet db = null;

            try
            {
                db = new Manta_dev_DataSet();
                
                DataCollection_Calendar = db.Calendar.ToList();
                DataCollection_Settings_Name_Column = db.Settings_Name_Columns.ToList();
                DataCollection_Settings_Group = db.Setting_Group.ToList();
            }
            catch (Exception ex)
            {
                Publish(ex, "Error while loading products.");
            }

            PrincipalContext ctx = new PrincipalContext(ContextType.Domain, "AD");
            UserPrincipal user = UserPrincipal.FindByIdentity(ctx, System.Web.HttpContext.Current.User.Identity.Name.Split('\\').Last());

            if (user != null)
            {
                var groups = user.GetAuthorizationGroups().OfType<GroupPrincipal>().Select(s => s.Name);
                List<string> g = groups.ToList();
                IsEditor = g.Contains("app_manta_editor");
                IsAdmin = g.Contains("app_manta_manager");

            }
            Message = string.Empty;
            

            foreach (Calendar c in DataCollection_Calendar)
            {
                if (c.RELEASE == "RE" + DataCollection_Settings_Name_Column.FindLast(f => f.User == All).Current_Release)
                {
                    if (c.CAPTION == Sys)
                    {
                        pom = c.STARTTIME.Value;
                        SysFrom = pom.ToString(DateFormat);
                        pom = c.ENDTIME.Value;
                        SysTo = pom.ToString(DateFormat);
                    }
                    else if (c.CAPTION == Integr)
                    {
                        pom = c.STARTTIME.Value;
                        IntFrom = pom.ToString(DateFormat);
                        pom = c.ENDTIME.Value;
                        IntTo = pom.ToString(DateFormat);

                    }
                    else if (c.CAPTION == Reg)
                    {
                        pom = c.STARTTIME.Value;
                        RegFrom = pom.ToString(DateFormat);
                        pom = c.ENDTIME.Value;
                        RegTo = pom.ToString(DateFormat);
                    }
                    else if (c.CAPTION == Deploy)
                    {
                        pom = c.STARTTIME.Value;
                        DeployProd = pom.ToString(DateFormat);
                    }
                }
            }
        }
        public void Publish(Exception ex, string message)
        {
            Publish(ex, message, null);
        }
        public void Publish(Exception ex, string message,
                            NameValueCollection nvc)
        {
            Message = message;
        }

        protected void BuildCollection()
        {
            Manta_dev_DataSet db = null;
            try
            {
                db = new Manta_dev_DataSet();
                
                DataCollection_Aktive_WP = db.Aktive_WP.ToList();
                DataCollection_Calendar = db.Calendar.ToList();
                DataCollection_Settings_Name_Column = db.Settings_Name_Columns.ToList();
                DataCollection_Settings_Group = db.Setting_Group.ToList();
            }
            catch (Exception ex)
            {
                Publish(ex, "Error while loading WP or Calendar.");
            }
        }

        public void HandleRequest()
        {
            BuildCollection();
        }
    }
}
