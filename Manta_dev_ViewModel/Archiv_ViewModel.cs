using System;
using System.Collections.Generic;
using System.Linq;
using System.Collections.Specialized;
using Manta_dev_Data;
using System.Security.Principal;
using System.DirectoryServices.AccountManagement;
using PagedList;
using PagedList.Mvc;

namespace Manta_dev_ViewModel
{
    public class Archiv_ViewModel
    {
        public IPagedList<Archiv> DataCollection_Archiv { get; set; }
        public List<Settings_Name_Columns> DataCollection_Settings_Name_Column { get; set; }
        public List<Setting_Group> DataCollection_Settings_Group { get; set; }

        public readonly string Username = WindowsIdentity.GetCurrent().Name.ToString();
        public bool IsEditor = false;
        public bool IsAdmin = false;
        public readonly string All = "All";
        public string Message { get; set; }

        public Archiv_ViewModel()
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
                
                DataCollection_Settings_Name_Column = db.Settings_Name_Columns.ToList();
                DataCollection_Settings_Group = db.Setting_Group.ToList();
            }
            catch (Exception ex)
            {
                Publish(ex, "Error while loading products.");
            }
            Message = string.Empty;

            PrincipalContext ctx = new PrincipalContext(ContextType.Domain, "AD");
            UserPrincipal user = UserPrincipal.FindByIdentity(ctx, System.Web.HttpContext.Current.User.Identity.Name.Split('\\').Last());

            if (user != null)
            {
                var groups = user.GetAuthorizationGroups().OfType<GroupPrincipal>().Select(s => s.Name);
                List<string> g = groups.ToList();
                IsEditor = g.Contains("app_manta_editor");
                IsAdmin = g.Contains("app_manta_manager");
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
        protected void BuildCollection(int? page)
        {
            Manta_dev_DataSet db = null;
            try
            {
                db = new Manta_dev_DataSet();

                DataCollection_Archiv = db.Archiv.ToList().ToPagedList(page ?? 1, 10);
                DataCollection_Settings_Name_Column = db.Settings_Name_Columns.ToList();
                DataCollection_Settings_Group = db.Setting_Group.ToList();
            }
            catch (Exception ex)
            {
                Publish(ex, "Error while loading WP or Calendar.");
            }
        }

        public void HandleRequest(int? page)
        {
            BuildCollection(page);
        }
    }
}
