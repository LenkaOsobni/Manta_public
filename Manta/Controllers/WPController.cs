using System;
using System.Linq;
using System.Web.Mvc;
using Manta_dev_ViewModel;
using Manta_dev_Data;
using System.Data.Entity;
using System.Security.Principal;
using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Data.Entity.Infrastructure;
using System.Data;
using System.Collections.Generic;
using System.DirectoryServices.AccountManagement;

namespace Manta_dev.Controllers
{

    public class Option  {
        public string text;
        public int id;
    }
    public class WPController : Controller
    {
        private Manta_dev_DataSet db = new Manta_dev_DataSet();
       // public string Username = User.Identity.Name.ToString().Split('\\').Last();   
        public bool IsEditor = false;
        public bool IsAdmin = false;
        public WPController()
        {
            PrincipalContext ctx = new PrincipalContext(ContextType.Domain, "AD");
            UserPrincipal user = UserPrincipal.FindByIdentity(ctx, System.Web.HttpContext.Current.User.Identity.Name.Split('\\').Last());

            if (user != null)
            {
                var groups = user.GetAuthorizationGroups().OfType<GroupPrincipal>().Select(s => s.Name);
                List<string> g = groups.ToList();
                IsEditor = g.Contains("app_manta_editor");
                IsAdmin = g.Contains("app_manta_manager");
            }

            Settings_Name_Columns Sett = new Settings_Name_Columns();
            Sett = db.Settings_Name_Columns.FirstOrDefault();
            ViewData["Name_Column"] = Sett;


            Setting_Group Adm = new Setting_Group();
            Adm = db.Setting_Group.Where(w => w.User_Group == "Admin").FirstOrDefault();
            ViewData["Admin"] = Adm;

            Setting_Group Edr = new Setting_Group();
            Edr = db.Setting_Group.Where(w => w.User_Group == "Editor").FirstOrDefault();
            ViewData["Editor"] = Edr;

            Setting_Group Al = new Setting_Group();
            Al = db.Setting_Group.Where(w => w.User_Group == "All").FirstOrDefault();
            ViewData["All"] = Al;
        }
 
        public bool saveFilter(int[] Column_1_Filter, int[] Column_2_Filter, int[] Column_3_Filter, int[] Column_4_Filter, int[] Column_5_Filter, int[] Column_6_Filter, int[] Column_7_Filter, int[] Column_8_Filter, int[] Column_9_Filter, int[] Column_10_Filter, int[] Column_11_Filter, int[] Column_12_Filter, int[] Column_13_Filter, int[] Column_14_Filter, int[] Column_15_Filter, int[] Column_16_Filter, int[] Column_17_Filter, int[] Column_18_Filter, int[] Column_19_Filter, int[] Column_20_Filter, int[] Column_21_Filter, int[] Column_22_Filter, int[] Column_23_Filter, int[] Column_24_Filter, int[] Column_25_Filter, int[] Column_26_Filter, int[] Column_27_Filter, int[] Column_28_Filter, int[] Column_29_Filter, int[] Column_30_Filter)
        {
            string userString = User.Identity.Name.ToString().Split('\\').Last();    
            //smaz user
            List<Settings_Filter> delFilter = db.Settings_Filter.Where(w => w.User == userString).ToList();
            for (int i = 0; i < delFilter.Count; i++)
            {
                db.Settings_Filter.Remove(delFilter[i]);
                db.SaveChanges();
            }
            //uloz nove
            Settings_Filter newSettingFilter = new Settings_Filter();
            newSettingFilter.User = userString;
            if(Column_1_Filter != null) newSettingFilter.Filter_Column_1 = string.Join(",", Column_1_Filter);
            if(Column_2_Filter != null) newSettingFilter.Filter_Column_2 = string.Join(",", Column_2_Filter);
            if(Column_3_Filter != null) newSettingFilter.Filter_Column_3 = string.Join(",", Column_3_Filter);
            if(Column_4_Filter != null) newSettingFilter.Filter_Column_4 = string.Join(",", Column_4_Filter);
            if(Column_5_Filter != null) newSettingFilter.Filter_Column_5 = string.Join(",", Column_5_Filter);
            if(Column_6_Filter != null) newSettingFilter.Filter_Column_6 = string.Join(",", Column_6_Filter);
            if(Column_7_Filter != null) newSettingFilter.Filter_Column_7 = string.Join(",", Column_7_Filter);
            if(Column_8_Filter != null) newSettingFilter.Filter_Column_8 = string.Join(",", Column_8_Filter);
            if(Column_9_Filter != null) newSettingFilter.Filter_Column_9 = string.Join(",", Column_9_Filter);
            if(Column_10_Filter != null) newSettingFilter.Filter_Column_10 = string.Join(",", Column_10_Filter);
            if(Column_11_Filter != null) newSettingFilter.Filter_Column_11 = string.Join(",", Column_11_Filter);
            if(Column_12_Filter != null) newSettingFilter.Filter_Column_12 = string.Join(",", Column_12_Filter);
            if(Column_13_Filter != null) newSettingFilter.Filter_Column_13 = string.Join(",", Column_13_Filter);
            if(Column_14_Filter != null) newSettingFilter.Filter_Column_14 = string.Join(",", Column_14_Filter);
            if(Column_15_Filter != null) newSettingFilter.Filter_Column_15 = string.Join(",", Column_15_Filter);
            if(Column_16_Filter != null) newSettingFilter.Filter_Column_16 = string.Join(",", Column_16_Filter);
            if(Column_17_Filter != null) newSettingFilter.Filter_Column_17 = string.Join(",", Column_17_Filter);
            if(Column_18_Filter != null) newSettingFilter.Filter_Column_18 = string.Join(",", Column_18_Filter);
            if(Column_19_Filter != null) newSettingFilter.Filter_Column_19 = string.Join(",", Column_19_Filter);
            if(Column_20_Filter != null) newSettingFilter.Filter_Column_20 = string.Join(",", Column_20_Filter);
            if(Column_21_Filter != null) newSettingFilter.Filter_Column_21 = string.Join(",", Column_21_Filter);
            if(Column_22_Filter != null) newSettingFilter.Filter_Column_22 = string.Join(",", Column_22_Filter);
            if(Column_23_Filter != null) newSettingFilter.Filter_Column_23 = string.Join(",", Column_23_Filter);
            if(Column_24_Filter != null) newSettingFilter.Filter_Column_24 = string.Join(",", Column_24_Filter);
            if(Column_25_Filter != null) newSettingFilter.Filter_Column_25 = string.Join(",", Column_25_Filter);
            if(Column_26_Filter != null) newSettingFilter.Filter_Column_26 = string.Join(",", Column_26_Filter);
            if(Column_27_Filter != null) newSettingFilter.Filter_Column_27 = string.Join(",", Column_27_Filter);
            if(Column_28_Filter != null) newSettingFilter.Filter_Column_28 = string.Join(",", Column_28_Filter);
            if(Column_29_Filter != null) newSettingFilter.Filter_Column_29 = string.Join(",", Column_29_Filter);
            if(Column_30_Filter != null) newSettingFilter.Filter_Column_30 = string.Join(",", Column_30_Filter);

            db.Settings_Filter.Add(newSettingFilter);
            db.SaveChanges();
            return true;
        }

        [HttpGet]
        public ActionResult DeleteFilter()
        {
            string userString = User.Identity.Name.ToString().Split('\\').Last();    
            //smaz user
            List<Settings_Filter> delFilter = db.Settings_Filter.Where(w => w.User == userString).ToList();
            for (int i = 0; i < delFilter.Count; i++)
            {
                db.Settings_Filter.Remove(delFilter[i]);
                db.SaveChanges();
            }
            //uloz nove
            Settings_Filter newSettingFilter = new Settings_Filter();
            newSettingFilter.User = userString;
            db.Settings_Filter.Add(newSettingFilter);
            db.SaveChanges();

            return RedirectToAction("WP");
        }

        [HttpGet]
        public ActionResult WP(int[] Column_1_Filter, int[] Column_2_Filter, int[] Column_3_Filter, int[] Column_4_Filter, int[] Column_5_Filter, int[] Column_6_Filter, int[] Column_7_Filter, int[] Column_8_Filter, int[] Column_9_Filter, int[] Column_10_Filter, int[] Column_11_Filter, int[] Column_12_Filter, int[] Column_13_Filter, int[] Column_14_Filter, int[] Column_15_Filter, int[] Column_16_Filter, int[] Column_17_Filter, int[] Column_18_Filter, int[] Column_19_Filter, int[] Column_20_Filter, int[] Column_21_Filter, int[] Column_22_Filter, int[] Column_23_Filter, int[] Column_24_Filter, int[] Column_25_Filter, int[] Column_26_Filter, int[] Column_27_Filter, int[] Column_28_Filter, int[] Column_29_Filter, int[] Column_30_Filter)
        {
            //new filter
            if (Column_1_Filter != null || Column_2_Filter != null || Column_3_Filter != null || Column_4_Filter != null || Column_5_Filter != null || Column_6_Filter != null || Column_7_Filter != null || Column_8_Filter != null || Column_9_Filter != null || Column_10_Filter != null || Column_11_Filter != null || Column_12_Filter != null || Column_13_Filter != null || Column_14_Filter != null || Column_15_Filter != null || Column_16_Filter != null || Column_17_Filter != null || Column_18_Filter != null || Column_19_Filter != null || Column_20_Filter != null || Column_21_Filter != null || Column_22_Filter != null || Column_23_Filter != null || Column_24_Filter != null || Column_25_Filter != null || Column_26_Filter != null || Column_27_Filter != null || Column_28_Filter != null || Column_29_Filter != null || Column_30_Filter != null)
            {
                saveFilter(Column_1_Filter, Column_2_Filter, Column_3_Filter, Column_4_Filter, Column_5_Filter, Column_6_Filter, Column_7_Filter, Column_8_Filter, Column_9_Filter, Column_10_Filter, Column_11_Filter, Column_12_Filter, Column_13_Filter, Column_14_Filter, Column_15_Filter, Column_16_Filter, Column_17_Filter, Column_18_Filter, Column_19_Filter, Column_20_Filter, Column_21_Filter, Column_22_Filter, Column_23_Filter, Column_24_Filter, Column_25_Filter, Column_26_Filter, Column_27_Filter, Column_28_Filter, Column_29_Filter, Column_30_Filter);
            }
            //new user
            string userString = User.Identity.Name.ToString().Split('\\').Last();   
            List<Settings_Filter> myFilter = db.Settings_Filter.Where(w => w.User == userString).ToList();
            if (myFilter.Count == 0)
            {
                Settings_Filter newSettingFilter = new Settings_Filter();
                newSettingFilter.User = userString;
                db.Settings_Filter.Add(newSettingFilter);
                db.SaveChanges();
            }
            WP_ViewModel vm = new WP_ViewModel { };
            vm.HandleRequest();
            List<Aktive_WP> array = db.Aktive_WP.ToList();

            List<Settings_Filter> correctFilter = db.Settings_Filter.Where(w => w.User == userString).ToList();
            if (correctFilter.FirstOrDefault().Filter_Column_1 != null) Column_1_Filter = correctFilter.FirstOrDefault().Filter_Column_1.Split(',').Select(n => Convert.ToInt32(n)).ToArray();
            if (correctFilter.FirstOrDefault().Filter_Column_2 != null) Column_2_Filter = correctFilter.FirstOrDefault().Filter_Column_2.Split(',').Select(n => Convert.ToInt32(n)).ToArray();
            if (correctFilter.FirstOrDefault().Filter_Column_3 != null) Column_3_Filter = correctFilter.FirstOrDefault().Filter_Column_3.Split(',').Select(n => Convert.ToInt32(n)).ToArray();
            if (correctFilter.FirstOrDefault().Filter_Column_4 != null) Column_4_Filter = correctFilter.FirstOrDefault().Filter_Column_4.Split(',').Select(n => Convert.ToInt32(n)).ToArray();
            if (correctFilter.FirstOrDefault().Filter_Column_5 != null) Column_5_Filter = correctFilter.FirstOrDefault().Filter_Column_5.Split(',').Select(n => Convert.ToInt32(n)).ToArray();
            if (correctFilter.FirstOrDefault().Filter_Column_6 != null) Column_6_Filter = correctFilter.FirstOrDefault().Filter_Column_6.Split(',').Select(n => Convert.ToInt32(n)).ToArray();
            if (correctFilter.FirstOrDefault().Filter_Column_7 != null) Column_7_Filter = correctFilter.FirstOrDefault().Filter_Column_7.Split(',').Select(n => Convert.ToInt32(n)).ToArray();
            if (correctFilter.FirstOrDefault().Filter_Column_8 != null) Column_8_Filter = correctFilter.FirstOrDefault().Filter_Column_8.Split(',').Select(n => Convert.ToInt32(n)).ToArray();
            if (correctFilter.FirstOrDefault().Filter_Column_9 != null) Column_9_Filter = correctFilter.FirstOrDefault().Filter_Column_9.Split(',').Select(n => Convert.ToInt32(n)).ToArray();
            if (correctFilter.FirstOrDefault().Filter_Column_10 != null) Column_10_Filter = correctFilter.FirstOrDefault().Filter_Column_10.Split(',').Select(n => Convert.ToInt32(n)).ToArray();
            if (correctFilter.FirstOrDefault().Filter_Column_11 != null) Column_11_Filter = correctFilter.FirstOrDefault().Filter_Column_11.Split(',').Select(n => Convert.ToInt32(n)).ToArray();
            if (correctFilter.FirstOrDefault().Filter_Column_12 != null) Column_12_Filter = correctFilter.FirstOrDefault().Filter_Column_12.Split(',').Select(n => Convert.ToInt32(n)).ToArray();
            if (correctFilter.FirstOrDefault().Filter_Column_13 != null) Column_13_Filter = correctFilter.FirstOrDefault().Filter_Column_13.Split(',').Select(n => Convert.ToInt32(n)).ToArray();
            if (correctFilter.FirstOrDefault().Filter_Column_14 != null) Column_14_Filter = correctFilter.FirstOrDefault().Filter_Column_14.Split(',').Select(n => Convert.ToInt32(n)).ToArray();
            if (correctFilter.FirstOrDefault().Filter_Column_15 != null) Column_15_Filter = correctFilter.FirstOrDefault().Filter_Column_15.Split(',').Select(n => Convert.ToInt32(n)).ToArray();
            if (correctFilter.FirstOrDefault().Filter_Column_16 != null) Column_16_Filter = correctFilter.FirstOrDefault().Filter_Column_16.Split(',').Select(n => Convert.ToInt32(n)).ToArray();
            if (correctFilter.FirstOrDefault().Filter_Column_17 != null) Column_17_Filter = correctFilter.FirstOrDefault().Filter_Column_17.Split(',').Select(n => Convert.ToInt32(n)).ToArray();
            if (correctFilter.FirstOrDefault().Filter_Column_18 != null) Column_18_Filter = correctFilter.FirstOrDefault().Filter_Column_18.Split(',').Select(n => Convert.ToInt32(n)).ToArray();
            if (correctFilter.FirstOrDefault().Filter_Column_19 != null) Column_19_Filter = correctFilter.FirstOrDefault().Filter_Column_19.Split(',').Select(n => Convert.ToInt32(n)).ToArray();
            if (correctFilter.FirstOrDefault().Filter_Column_20 != null) Column_20_Filter = correctFilter.FirstOrDefault().Filter_Column_20.Split(',').Select(n => Convert.ToInt32(n)).ToArray();
            if (correctFilter.FirstOrDefault().Filter_Column_21 != null) Column_21_Filter = correctFilter.FirstOrDefault().Filter_Column_21.Split(',').Select(n => Convert.ToInt32(n)).ToArray();
            if (correctFilter.FirstOrDefault().Filter_Column_22 != null) Column_22_Filter = correctFilter.FirstOrDefault().Filter_Column_22.Split(',').Select(n => Convert.ToInt32(n)).ToArray();
            if (correctFilter.FirstOrDefault().Filter_Column_23 != null) Column_23_Filter = correctFilter.FirstOrDefault().Filter_Column_23.Split(',').Select(n => Convert.ToInt32(n)).ToArray();
            if (correctFilter.FirstOrDefault().Filter_Column_24 != null) Column_24_Filter = correctFilter.FirstOrDefault().Filter_Column_24.Split(',').Select(n => Convert.ToInt32(n)).ToArray();
            if (correctFilter.FirstOrDefault().Filter_Column_25 != null) Column_25_Filter = correctFilter.FirstOrDefault().Filter_Column_25.Split(',').Select(n => Convert.ToInt32(n)).ToArray();
            if (correctFilter.FirstOrDefault().Filter_Column_26 != null) Column_26_Filter = correctFilter.FirstOrDefault().Filter_Column_26.Split(',').Select(n => Convert.ToInt32(n)).ToArray();
            if (correctFilter.FirstOrDefault().Filter_Column_27 != null) Column_27_Filter = correctFilter.FirstOrDefault().Filter_Column_27.Split(',').Select(n => Convert.ToInt32(n)).ToArray();
            if (correctFilter.FirstOrDefault().Filter_Column_28 != null) Column_28_Filter = correctFilter.FirstOrDefault().Filter_Column_28.Split(',').Select(n => Convert.ToInt32(n)).ToArray();
            if (correctFilter.FirstOrDefault().Filter_Column_29 != null) Column_29_Filter = correctFilter.FirstOrDefault().Filter_Column_29.Split(',').Select(n => Convert.ToInt32(n)).ToArray();
            if (correctFilter.FirstOrDefault().Filter_Column_30 != null) Column_30_Filter = correctFilter.FirstOrDefault().Filter_Column_30.Split(',').Select(n => Convert.ToInt32(n)).ToArray();
            try
            {
                if ((Column_1_Filter == null ? 0 : Column_1_Filter.Count()) > 0)
                {
                    List<Aktive_WP> array_Filter = new List<Aktive_WP>();
                    for (int i = 0; i < Column_1_Filter.Count(); i++)
                    {
                        int a = Column_1_Filter[i];
                        string u = db.Aktive_WP.Where(x => x.Id == a).FirstOrDefault().Column_1;
                        if (u == null || u == "")
                        {
                            array_Filter.AddRange(db.Aktive_WP.Where(s => s.Column_1 == null).DefaultIfEmpty().ToList());
                            array_Filter.AddRange(db.Aktive_WP.Where(s => s.Column_1 == "").DefaultIfEmpty().ToList());
                        }
                        else
                        {
                            array_Filter.AddRange(db.Aktive_WP.Where(s => s.Column_1 == u).DefaultIfEmpty().ToList());
                        }
                    }
                    List<Aktive_WP> k = array.Intersect(array_Filter).ToList();
                    array = k;
                }
                if ((Column_2_Filter == null ? 0 : Column_2_Filter.Count()) > 0)
                {
                    List<Aktive_WP> array_Filter = new List<Aktive_WP>();
                    for (int i = 0; i < Column_2_Filter.Count(); i++)
                    {
                        int a = Column_2_Filter[i];
                        string u = db.Aktive_WP.Where(x => x.Id == a).FirstOrDefault().Column_2;
                        if (u == null || u == "")
                        {
                            array_Filter.AddRange(db.Aktive_WP.Where(s => s.Column_2 == null).DefaultIfEmpty().ToList());
                            array_Filter.AddRange(db.Aktive_WP.Where(s => s.Column_2 == "").DefaultIfEmpty().ToList());
                        }
                        else
                        {
                            array_Filter.AddRange(db.Aktive_WP.Where(s => s.Column_2 == u).DefaultIfEmpty().ToList());
                        }
                    }
                    List<Aktive_WP> k = array.Intersect(array_Filter).ToList();
                    array = k;
                }
                if ((Column_3_Filter == null ? 0 : Column_3_Filter.Count()) > 0)
                {
                    List<Aktive_WP> array_Filter = new List<Aktive_WP>();
                    for (int i = 0; i < Column_3_Filter.Count(); i++)
                    {
                        int a = Column_3_Filter[i];
                        string u = db.Aktive_WP.Where(x => x.Id == a).FirstOrDefault().Column_3;
                        if (u == null || u == "")
                        {
                            array_Filter.AddRange(db.Aktive_WP.Where(s => s.Column_3 == null).DefaultIfEmpty().ToList());
                            array_Filter.AddRange(db.Aktive_WP.Where(s => s.Column_3 == "").DefaultIfEmpty().ToList());
                        }
                        else
                        {
                            array_Filter.AddRange(db.Aktive_WP.Where(s => s.Column_3 == u).DefaultIfEmpty().ToList());
                        }
                    }
                    List<Aktive_WP> k = array.Intersect(array_Filter).ToList();
                    array = k;
                }
                if ((Column_4_Filter == null ? 0 : Column_4_Filter.Count()) > 0)
                {
                    List<Aktive_WP> array_Filter = new List<Aktive_WP>();
                    for (int i = 0; i < Column_4_Filter.Count(); i++)
                    {
                        int a = Column_4_Filter[i];
                        string u = db.Aktive_WP.Where(x => x.Id == a).FirstOrDefault().Column_4;
                        if (u == null || u == "")
                        {
                            array_Filter.AddRange(db.Aktive_WP.Where(s => s.Column_4 == null).DefaultIfEmpty().ToList());
                            array_Filter.AddRange(db.Aktive_WP.Where(s => s.Column_4 == "").DefaultIfEmpty().ToList());
                        }
                        else
                        {
                            array_Filter.AddRange(db.Aktive_WP.Where(s => s.Column_4 == u).DefaultIfEmpty().ToList());
                        }
                    }
                    List<Aktive_WP> k = array.Intersect(array_Filter).ToList();
                    array = k;
                }
                if ((Column_5_Filter == null ? 0 : Column_5_Filter.Count()) > 0)
                {
                    List<Aktive_WP> array_Filter = new List<Aktive_WP>();
                    for (int i = 0; i < Column_5_Filter.Count(); i++)
                    {
                        int a = Column_5_Filter[i];
                        string u = db.Aktive_WP.Where(x => x.Id == a).FirstOrDefault().Column_5;
                        if (u == null || u == "")
                        {
                            array_Filter.AddRange(db.Aktive_WP.Where(s => s.Column_5 == null).DefaultIfEmpty().ToList());
                            array_Filter.AddRange(db.Aktive_WP.Where(s => s.Column_5 == "").DefaultIfEmpty().ToList());
                        }
                        else
                        {
                            array_Filter.AddRange(db.Aktive_WP.Where(s => s.Column_5 == u).DefaultIfEmpty().ToList());
                        }
                    }
                    List<Aktive_WP> k = array.Intersect(array_Filter).ToList();
                    array = k;
                }
                if ((Column_6_Filter == null ? 0 : Column_6_Filter.Count()) > 0)
                {
                    List<Aktive_WP> array_Filter = new List<Aktive_WP>();
                    for (int i = 0; i < Column_6_Filter.Count(); i++)
                    {
                        int a = Column_6_Filter[i];
                        string u = db.Aktive_WP.Where(x => x.Id == a).FirstOrDefault().Column_6;
                        if (u == null || u == "")
                        {
                            array_Filter.AddRange(db.Aktive_WP.Where(s => s.Column_6 == null).DefaultIfEmpty().ToList());
                            array_Filter.AddRange(db.Aktive_WP.Where(s => s.Column_6 == "").DefaultIfEmpty().ToList());
                        }
                        else
                        {
                            array_Filter.AddRange(db.Aktive_WP.Where(s => s.Column_6 == u).DefaultIfEmpty().ToList());
                        }
                    }
                    List<Aktive_WP> k = array.Intersect(array_Filter).ToList();
                    array = k;
                }
                if ((Column_7_Filter == null ? 0 : Column_7_Filter.Count()) > 0)
                {
                    List<Aktive_WP> array_Filter = new List<Aktive_WP>();
                    for (int i = 0; i < Column_7_Filter.Count(); i++)
                    {
                        int a = Column_7_Filter[i];
                        string u = db.Aktive_WP.Where(x => x.Id == a).FirstOrDefault().Column_7;
                        if (u == null || u == "")
                        {
                            array_Filter.AddRange(db.Aktive_WP.Where(s => s.Column_7 == null).DefaultIfEmpty().ToList());
                            array_Filter.AddRange(db.Aktive_WP.Where(s => s.Column_7 == "").DefaultIfEmpty().ToList());
                        }
                        else
                        {
                            array_Filter.AddRange(db.Aktive_WP.Where(s => s.Column_7 == u).DefaultIfEmpty().ToList());
                        }
                    }
                    List<Aktive_WP> k = array.Intersect(array_Filter).ToList();
                    array = k;
                }
                if ((Column_8_Filter == null ? 0 : Column_8_Filter.Count()) > 0)
                {
                    List<Aktive_WP> array_Filter = new List<Aktive_WP>();
                    for (int i = 0; i < Column_8_Filter.Count(); i++)
                    {
                        int a = Column_8_Filter[i];
                        string u = db.Aktive_WP.Where(x => x.Id == a).FirstOrDefault().Column_8;
                        if (u == null || u == "")
                        {
                            array_Filter.AddRange(db.Aktive_WP.Where(s => s.Column_8 == null).DefaultIfEmpty().ToList());
                            array_Filter.AddRange(db.Aktive_WP.Where(s => s.Column_8 == "").DefaultIfEmpty().ToList());
                        }
                        else
                        {
                            array_Filter.AddRange(db.Aktive_WP.Where(s => s.Column_8 == u).DefaultIfEmpty().ToList());
                        }
                    }
                    List<Aktive_WP> k = array.Intersect(array_Filter).ToList();
                    array = k;
                }
                if ((Column_9_Filter == null ? 0 : Column_9_Filter.Count()) > 0)
                {
                    List<Aktive_WP> array_Filter = new List<Aktive_WP>();
                    for (int i = 0; i < Column_9_Filter.Count(); i++)
                    {
                        int a = Column_9_Filter[i];
                        string u = db.Aktive_WP.Where(x => x.Id == a).FirstOrDefault().Column_9;
                        if (u == null || u == "")
                        {
                            array_Filter.AddRange(db.Aktive_WP.Where(s => s.Column_9 == null).DefaultIfEmpty().ToList());
                            array_Filter.AddRange(db.Aktive_WP.Where(s => s.Column_9 == "").DefaultIfEmpty().ToList());
                        }
                        else
                        {
                            array_Filter.AddRange(db.Aktive_WP.Where(s => s.Column_9 == u).DefaultIfEmpty().ToList());
                        }
                    }
                    List<Aktive_WP> k = array.Intersect(array_Filter).ToList();
                    array = k;
                }
                if ((Column_10_Filter == null ? 0 : Column_10_Filter.Count()) > 0)
                {
                    List<Aktive_WP> array_Filter = new List<Aktive_WP>();
                    for (int i = 0; i < Column_10_Filter.Count(); i++)
                    {
                        int a = Column_10_Filter[i];
                        string u = db.Aktive_WP.Where(x => x.Id == a).FirstOrDefault().Column_10;
                        if (u == null || u == "")
                        {
                            array_Filter.AddRange(db.Aktive_WP.Where(s => s.Column_10 == null).DefaultIfEmpty().ToList());
                            array_Filter.AddRange(db.Aktive_WP.Where(s => s.Column_10 == "").DefaultIfEmpty().ToList());
                        }
                        else
                        {
                            array_Filter.AddRange(db.Aktive_WP.Where(s => s.Column_10 == u).DefaultIfEmpty().ToList());
                        }
                    }
                    List<Aktive_WP> k = array.Intersect(array_Filter).ToList();
                    array = k;
                }
                if ((Column_11_Filter == null ? 0 : Column_11_Filter.Count()) > 0)
                {
                    List<Aktive_WP> array_Filter = new List<Aktive_WP>();
                    for (int i = 0; i < Column_11_Filter.Count(); i++)
                    {
                        int a = Column_11_Filter[i];
                        string u = db.Aktive_WP.Where(x => x.Id == a).FirstOrDefault().Column_11;
                        if (u == null || u == "")
                        {
                            array_Filter.AddRange(db.Aktive_WP.Where(s => s.Column_11 == null).DefaultIfEmpty().ToList());
                            array_Filter.AddRange(db.Aktive_WP.Where(s => s.Column_11 == "").DefaultIfEmpty().ToList());
                        }
                        else
                        {
                            array_Filter.AddRange(db.Aktive_WP.Where(s => s.Column_11 == u).DefaultIfEmpty().ToList());
                        }
                    }
                    List<Aktive_WP> k = array.Intersect(array_Filter).ToList();
                    array = k;
                }
                if ((Column_12_Filter == null ? 0 : Column_12_Filter.Count()) > 0)
                {
                    List<Aktive_WP> array_Filter = new List<Aktive_WP>();
                    for (int i = 0; i < Column_12_Filter.Count(); i++)
                    {
                        int a = Column_12_Filter[i];
                        string u = db.Aktive_WP.Where(x => x.Id == a).FirstOrDefault().Column_12;
                        if (u == null || u == "")
                        {
                            array_Filter.AddRange(db.Aktive_WP.Where(s => s.Column_12 == null).DefaultIfEmpty().ToList());
                            array_Filter.AddRange(db.Aktive_WP.Where(s => s.Column_12 == "").DefaultIfEmpty().ToList());
                        }
                        else
                        {
                            array_Filter.AddRange(db.Aktive_WP.Where(s => s.Column_12 == u).DefaultIfEmpty().ToList());
                        }
                    }
                    List<Aktive_WP> k = array.Intersect(array_Filter).ToList();
                    array = k;
                }
                if ((Column_13_Filter == null ? 0 : Column_13_Filter.Count()) > 0)
                {
                    List<Aktive_WP> array_Filter = new List<Aktive_WP>();
                    for (int i = 0; i < Column_13_Filter.Count(); i++)
                    {
                        int a = Column_13_Filter[i];
                        string u = db.Aktive_WP.Where(x => x.Id == a).FirstOrDefault().Column_13;
                        if (u == null || u == "")
                        {
                            array_Filter.AddRange(db.Aktive_WP.Where(s => s.Column_13 == null).DefaultIfEmpty().ToList());
                            array_Filter.AddRange(db.Aktive_WP.Where(s => s.Column_13 == "").DefaultIfEmpty().ToList());
                        }
                        else
                        {
                            array_Filter.AddRange(db.Aktive_WP.Where(s => s.Column_13 == u).DefaultIfEmpty().ToList());
                        }
                    }
                    List<Aktive_WP> k = array.Intersect(array_Filter).ToList();
                    array = k;
                }
                if ((Column_14_Filter == null ? 0 : Column_14_Filter.Count()) > 0)
                {
                    List<Aktive_WP> array_Filter = new List<Aktive_WP>();
                    for (int i = 0; i < Column_14_Filter.Count(); i++)
                    {
                        int a = Column_14_Filter[i];
                        string u = db.Aktive_WP.Where(x => x.Id == a).FirstOrDefault().Column_14;
                        if (u == null || u == "")
                        {
                            array_Filter.AddRange(db.Aktive_WP.Where(s => s.Column_14 == null).DefaultIfEmpty().ToList());
                            array_Filter.AddRange(db.Aktive_WP.Where(s => s.Column_14 == "").DefaultIfEmpty().ToList());
                        }
                        else
                        {
                            array_Filter.AddRange(db.Aktive_WP.Where(s => s.Column_14 == u).DefaultIfEmpty().ToList());
                        }
                    }
                    List<Aktive_WP> k = array.Intersect(array_Filter).ToList();
                    array = k;
                }
                if ((Column_15_Filter == null ? 0 : Column_15_Filter.Count()) > 0)
                {
                    List<Aktive_WP> array_Filter = new List<Aktive_WP>();
                    for (int i = 0; i < Column_15_Filter.Count(); i++)
                    {
                        int a = Column_15_Filter[i];
                        string u = db.Aktive_WP.Where(x => x.Id == a).FirstOrDefault().Column_15;
                        if (u == null || u == "")
                        {
                            array_Filter.AddRange(db.Aktive_WP.Where(s => s.Column_15 == null).DefaultIfEmpty().ToList());
                            array_Filter.AddRange(db.Aktive_WP.Where(s => s.Column_15 == "").DefaultIfEmpty().ToList());
                        }
                        else
                        {
                            array_Filter.AddRange(db.Aktive_WP.Where(s => s.Column_15 == u).DefaultIfEmpty().ToList());
                        }
                    }
                    List<Aktive_WP> k = array.Intersect(array_Filter).ToList();
                    array = k;
                }
                if ((Column_16_Filter == null ? 0 : Column_16_Filter.Count()) > 0)
                {
                    List<Aktive_WP> array_Filter = new List<Aktive_WP>();
                    for (int i = 0; i < Column_16_Filter.Count(); i++)
                    {
                        int a = Column_16_Filter[i];
                        string u = db.Aktive_WP.Where(x => x.Id == a).FirstOrDefault().Column_16;
                        if (u == null || u == "")
                        {
                            array_Filter.AddRange(db.Aktive_WP.Where(s => s.Column_16 == null).DefaultIfEmpty().ToList());
                            array_Filter.AddRange(db.Aktive_WP.Where(s => s.Column_16 == "").DefaultIfEmpty().ToList());
                        }
                        else
                        {
                            array_Filter.AddRange(db.Aktive_WP.Where(s => s.Column_16 == u).DefaultIfEmpty().ToList());
                        }
                    }
                    List<Aktive_WP> k = array.Intersect(array_Filter).ToList();
                    array = k;
                }
                if ((Column_17_Filter == null ? 0 : Column_17_Filter.Count()) > 0)
                {
                    List<Aktive_WP> array_Filter = new List<Aktive_WP>();
                    for (int i = 0; i < Column_17_Filter.Count(); i++)
                    {
                        int a = Column_17_Filter[i];
                        string u = db.Aktive_WP.Where(x => x.Id == a).FirstOrDefault().Column_17;
                        if (u == null || u == "")
                        {
                            array_Filter.AddRange(db.Aktive_WP.Where(s => s.Column_17 == null).DefaultIfEmpty().ToList());
                            array_Filter.AddRange(db.Aktive_WP.Where(s => s.Column_17 == "").DefaultIfEmpty().ToList());
                        }
                        else
                        {
                            array_Filter.AddRange(db.Aktive_WP.Where(s => s.Column_17 == u).DefaultIfEmpty().ToList());
                        }
                    }
                    List<Aktive_WP> k = array.Intersect(array_Filter).ToList();
                    array = k;
                }
                if ((Column_18_Filter == null ? 0 : Column_18_Filter.Count()) > 0)
                {
                    List<Aktive_WP> array_Filter = new List<Aktive_WP>();
                    for (int i = 0; i < Column_18_Filter.Count(); i++)
                    {
                        int a = Column_18_Filter[i];
                        string u = db.Aktive_WP.Where(x => x.Id == a).FirstOrDefault().Column_18;
                        if (u == null || u == "")
                        {
                            array_Filter.AddRange(db.Aktive_WP.Where(s => s.Column_18 == null).DefaultIfEmpty().ToList());
                            array_Filter.AddRange(db.Aktive_WP.Where(s => s.Column_18 == "").DefaultIfEmpty().ToList());
                        }
                        else
                        {
                            array_Filter.AddRange(db.Aktive_WP.Where(s => s.Column_18 == u).DefaultIfEmpty().ToList());
                        }
                    }
                    List<Aktive_WP> k = array.Intersect(array_Filter).ToList();
                    array = k;
                }
                if ((Column_19_Filter == null ? 0 : Column_19_Filter.Count()) > 0)
                {
                    List<Aktive_WP> array_Filter = new List<Aktive_WP>();
                    for (int i = 0; i < Column_19_Filter.Count(); i++)
                    {
                        int a = Column_19_Filter[i];
                        string u = db.Aktive_WP.Where(x => x.Id == a).FirstOrDefault().Column_19;
                        if (u == null || u == "")
                        {
                            array_Filter.AddRange(db.Aktive_WP.Where(s => s.Column_19 == null).DefaultIfEmpty().ToList());
                            array_Filter.AddRange(db.Aktive_WP.Where(s => s.Column_19 == "").DefaultIfEmpty().ToList());
                        }
                        else
                        {
                            array_Filter.AddRange(db.Aktive_WP.Where(s => s.Column_19 == u).DefaultIfEmpty().ToList());
                        }
                    }
                    List<Aktive_WP> k = array.Intersect(array_Filter).ToList();
                    array = k;
                }
                if ((Column_20_Filter == null ? 0 : Column_20_Filter.Count()) > 0)
                {
                    List<Aktive_WP> array_Filter = new List<Aktive_WP>();
                    for (int i = 0; i < Column_20_Filter.Count(); i++)
                    {
                        int a = Column_20_Filter[i];
                        string u = db.Aktive_WP.Where(x => x.Id == a).FirstOrDefault().Column_20;
                        if (u == null || u == "")
                        {
                            array_Filter.AddRange(db.Aktive_WP.Where(s => s.Column_20 == null).DefaultIfEmpty().ToList());
                            array_Filter.AddRange(db.Aktive_WP.Where(s => s.Column_20 == "").DefaultIfEmpty().ToList());
                        }
                        else
                        {
                            array_Filter.AddRange(db.Aktive_WP.Where(s => s.Column_20 == u).DefaultIfEmpty().ToList());
                        }
                    }
                    List<Aktive_WP> k = array.Intersect(array_Filter).ToList();
                    array = k;
                }
                if ((Column_21_Filter == null ? 0 : Column_21_Filter.Count()) > 0)
                {
                    List<Aktive_WP> array_Filter = new List<Aktive_WP>();
                    for (int i = 0; i < Column_21_Filter.Count(); i++)
                    {
                        int a = Column_21_Filter[i];
                        string u = db.Aktive_WP.Where(x => x.Id == a).FirstOrDefault().Column_21;
                        if (u == null || u == "")
                        {
                            array_Filter.AddRange(db.Aktive_WP.Where(s => s.Column_21 == null).DefaultIfEmpty().ToList());
                            array_Filter.AddRange(db.Aktive_WP.Where(s => s.Column_21 == "").DefaultIfEmpty().ToList());
                        }
                        else
                        {
                            array_Filter.AddRange(db.Aktive_WP.Where(s => s.Column_21 == u).DefaultIfEmpty().ToList());
                        }
                    }
                    List<Aktive_WP> k = array.Intersect(array_Filter).ToList();
                    array = k;
                }
                if ((Column_22_Filter == null ? 0 : Column_22_Filter.Count()) > 0)
                {
                    List<Aktive_WP> array_Filter = new List<Aktive_WP>();
                    for (int i = 0; i < Column_22_Filter.Count(); i++)
                    {
                        int a = Column_22_Filter[i];
                        string u = db.Aktive_WP.Where(x => x.Id == a).FirstOrDefault().Column_22;
                        if (u == null || u == "")
                        {
                            array_Filter.AddRange(db.Aktive_WP.Where(s => s.Column_22 == null).DefaultIfEmpty().ToList());
                            array_Filter.AddRange(db.Aktive_WP.Where(s => s.Column_22 == "").DefaultIfEmpty().ToList());
                        }
                        else
                        {
                            array_Filter.AddRange(db.Aktive_WP.Where(s => s.Column_22 == u).DefaultIfEmpty().ToList());
                        }
                    }
                    List<Aktive_WP> k = array.Intersect(array_Filter).ToList();
                    array = k;
                }
                if ((Column_23_Filter == null ? 0 : Column_23_Filter.Count()) > 0)
                {
                    List<Aktive_WP> array_Filter = new List<Aktive_WP>();
                    for (int i = 0; i < Column_23_Filter.Count(); i++)
                    {
                        int a = Column_23_Filter[i];
                        string u = db.Aktive_WP.Where(x => x.Id == a).FirstOrDefault().Column_23;
                        if (u == null || u == "")
                        {
                            array_Filter.AddRange(db.Aktive_WP.Where(s => s.Column_23 == null).DefaultIfEmpty().ToList());
                            array_Filter.AddRange(db.Aktive_WP.Where(s => s.Column_23 == "").DefaultIfEmpty().ToList());
                        }
                        else
                        {
                            array_Filter.AddRange(db.Aktive_WP.Where(s => s.Column_23 == u).DefaultIfEmpty().ToList());
                        }
                    }
                    List<Aktive_WP> k = array.Intersect(array_Filter).ToList();
                    array = k;
                }
                if ((Column_24_Filter == null ? 0 : Column_24_Filter.Count()) > 0)
                {
                    List<Aktive_WP> array_Filter = new List<Aktive_WP>();
                    for (int i = 0; i < Column_24_Filter.Count(); i++)
                    {
                        int a = Column_24_Filter[i];
                        string u = db.Aktive_WP.Where(x => x.Id == a).FirstOrDefault().Column_24;
                        if (u == null || u == "")
                        {
                            array_Filter.AddRange(db.Aktive_WP.Where(s => s.Column_24 == null).DefaultIfEmpty().ToList());
                            array_Filter.AddRange(db.Aktive_WP.Where(s => s.Column_24 == "").DefaultIfEmpty().ToList());
                        }
                        else
                        {
                            array_Filter.AddRange(db.Aktive_WP.Where(s => s.Column_24 == u).DefaultIfEmpty().ToList());
                        }
                    }
                    List<Aktive_WP> k = array.Intersect(array_Filter).ToList();
                    array = k;
                }
                if ((Column_25_Filter == null ? 0 : Column_25_Filter.Count()) > 0)
                {
                    List<Aktive_WP> array_Filter = new List<Aktive_WP>();
                    for (int i = 0; i < Column_25_Filter.Count(); i++)
                    {
                        int a = Column_25_Filter[i];
                        string u = db.Aktive_WP.Where(x => x.Id == a).FirstOrDefault().Column_25;
                        if (u == null || u == "")
                        {
                            array_Filter.AddRange(db.Aktive_WP.Where(s => s.Column_25 == null).DefaultIfEmpty().ToList());
                            array_Filter.AddRange(db.Aktive_WP.Where(s => s.Column_25 == "").DefaultIfEmpty().ToList());
                        }
                        else
                        {
                            array_Filter.AddRange(db.Aktive_WP.Where(s => s.Column_25 == u).DefaultIfEmpty().ToList());
                        }
                    }
                    List<Aktive_WP> k = array.Intersect(array_Filter).ToList();
                    array = k;
                }
                if ((Column_26_Filter == null ? 0 : Column_26_Filter.Count()) > 0)
                {
                    List<Aktive_WP> array_Filter = new List<Aktive_WP>();
                    for (int i = 0; i < Column_26_Filter.Count(); i++)
                    {
                        int a = Column_26_Filter[i];
                        string u = db.Aktive_WP.Where(x => x.Id == a).FirstOrDefault().Column_26;
                        if (u == null || u == "")
                        {
                            array_Filter.AddRange(db.Aktive_WP.Where(s => s.Column_26 == null).DefaultIfEmpty().ToList());
                            array_Filter.AddRange(db.Aktive_WP.Where(s => s.Column_26 == "").DefaultIfEmpty().ToList());
                        }
                        else
                        {
                            array_Filter.AddRange(db.Aktive_WP.Where(s => s.Column_26 == u).DefaultIfEmpty().ToList());
                        }
                    }
                    List<Aktive_WP> k = array.Intersect(array_Filter).ToList();
                    array = k;
                }
                if ((Column_27_Filter == null ? 0 : Column_27_Filter.Count()) > 0)
                {
                    List<Aktive_WP> array_Filter = new List<Aktive_WP>();
                    for (int i = 0; i < Column_27_Filter.Count(); i++)
                    {
                        int a = Column_27_Filter[i];
                        string u = db.Aktive_WP.Where(x => x.Id == a).FirstOrDefault().Column_27;
                        if (u == null || u == "")
                        {
                            array_Filter.AddRange(db.Aktive_WP.Where(s => s.Column_27 == null).DefaultIfEmpty().ToList());
                            array_Filter.AddRange(db.Aktive_WP.Where(s => s.Column_27 == "").DefaultIfEmpty().ToList());
                        }
                        else
                        {
                            array_Filter.AddRange(db.Aktive_WP.Where(s => s.Column_27 == u).DefaultIfEmpty().ToList());
                        }
                    }
                    List<Aktive_WP> k = array.Intersect(array_Filter).ToList();
                    array = k;
                }
                if ((Column_28_Filter == null ? 0 : Column_28_Filter.Count()) > 0)
                {
                    List<Aktive_WP> array_Filter = new List<Aktive_WP>();
                    for (int i = 0; i < Column_28_Filter.Count(); i++)
                    {
                        int a = Column_28_Filter[i];
                        string u = db.Aktive_WP.Where(x => x.Id == a).FirstOrDefault().Column_29;
                        if (u == null || u == "")
                        {
                            array_Filter.AddRange(db.Aktive_WP.Where(s => s.Column_28 == null).DefaultIfEmpty().ToList());
                            array_Filter.AddRange(db.Aktive_WP.Where(s => s.Column_28 == "").DefaultIfEmpty().ToList());
                        }
                        else
                        {
                            array_Filter.AddRange(db.Aktive_WP.Where(s => s.Column_28 == u).DefaultIfEmpty().ToList());
                        }
                    }
                    List<Aktive_WP> k = array.Intersect(array_Filter).ToList();
                    array = k;
                }
                if ((Column_29_Filter == null ? 0 : Column_29_Filter.Count()) > 0)
                {
                    List<Aktive_WP> array_Filter = new List<Aktive_WP>();
                    for (int i = 0; i < Column_29_Filter.Count(); i++)
                    {
                        int a = Column_29_Filter[i];
                        string u = db.Aktive_WP.Where(x => x.Id == a).FirstOrDefault().Column_29;
                        if (u == null || u == "")
                        {
                            array_Filter.AddRange(db.Aktive_WP.Where(s => s.Column_29 == null).DefaultIfEmpty().ToList());
                            array_Filter.AddRange(db.Aktive_WP.Where(s => s.Column_29 == "").DefaultIfEmpty().ToList());
                        }
                        else
                        {
                            array_Filter.AddRange(db.Aktive_WP.Where(s => s.Column_29 == u).DefaultIfEmpty().ToList());
                        }
                    }
                    List<Aktive_WP> k = array.Intersect(array_Filter).ToList();
                    array = k;
                }
                if ((Column_30_Filter == null ? 0 : Column_30_Filter.Count()) > 0)
                {
                    List<Aktive_WP> array_Filter = new List<Aktive_WP>();
                    for (int i = 0; i < Column_30_Filter.Count(); i++)
                    {
                        int a = Column_30_Filter[i];
                        string u = db.Aktive_WP.Where(x => x.Id == a).FirstOrDefault().Column_30;
                        if (u == null || u == "")
                        {
                            array_Filter.AddRange(db.Aktive_WP.Where(s => s.Column_30 == null).DefaultIfEmpty().ToList());
                            array_Filter.AddRange(db.Aktive_WP.Where(s => s.Column_30 == "").DefaultIfEmpty().ToList());
                        }
                        else
                        {
                            array_Filter.AddRange(db.Aktive_WP.Where(s => s.Column_30 == u).DefaultIfEmpty().ToList());
                        }
                    }
                    List<Aktive_WP> k = array.Intersect(array_Filter).ToList();
                    array = k;
                }
                vm.DataCollection_Aktive_WP = array.Distinct().OrderBy(x => x.Column_1).ToList();

                return View(vm);
            }
            catch (Exception ex)
            {
                DeleteFilter();
                vm.DataCollection_Aktive_WP = db.Aktive_WP.OrderBy(x => x.Column_1).ToList();
                return View(vm);

            }
            }

        [HttpGet]
        public ActionResult WP_Add()
        {
            return View();
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult WP_Add([Bind(Include = "Column_1, Column_2, Column_3, Column_4, Column_5, Column_6, Column_7, Column_8, Column_9, Column_10, Column_11, Column_12, Column_13, Column_14, Column_15, Column_16, Column_17, Column_18, Column_19, Column_20, Column_21, Column_22, Column_23, Column_24, Column_25, Column_26, Column_27, Column_28, Column_29, Column_30")]Aktive_WP awp)
        {
            Settings_Name_Columns Sett = new Settings_Name_Columns();
            Sett = db.Settings_Name_Columns.First();
            ViewData["Name_Column"] = Sett;
            if(db.Aktive_WP.FirstOrDefault(f => f.Column_1 == awp.Column_1) != null) return RedirectToAction("WP");
            if (ModelState.IsValid)
            {
                db.Aktive_WP.Add(awp);
                awp.Color_1 = "#ffffff";
                awp.Color_2 = "#ffffff";
                awp.Color_3 = "#ffffff";
                awp.Color_4 = "#ffffff";
                awp.Color_5 = "#ffffff";
                awp.Color_6 = "#ffffff";
                awp.Column_2 = "";
                awp.Column_3 = "";
                awp.Column_4 = "";
                awp.Column_5 = "";
                awp.Column_6 = "";
                awp.Column_7 = "";
                awp.Column_8 = "";
                awp.Column_9 = "";
                awp.Column_10 = "";
                awp.Column_11 = "";
                awp.Column_12 = "";
                awp.Column_13 = "";
                awp.Column_14 = "";
                awp.Column_15 = "";
                awp.Column_16 = "";
                awp.Column_17 = "";
                awp.Column_18 = "";
                awp.Column_19 = "";
                awp.Column_20 = "";
                awp.Column_21 = "";
                awp.Column_22 = "";
                awp.Column_23 = "";
                awp.Column_24 = "";
                awp.Column_25 = "";
                awp.Column_26 = "";
                awp.Column_27 = "";
                awp.Column_28 = "";
                awp.Column_29 = "";
                awp.Column_30 = "";
                db.SaveChanges();
            }

            return RedirectToAction("WP");
        }

        [HttpGet]
        public ActionResult WP_Edit(int id)
        {
            var editWP = db.Aktive_WP.Find(id);
            Settings_Name_Columns Sett = new Settings_Name_Columns();
            Sett = db.Settings_Name_Columns.First();
            ViewData["Name_Column"] = Sett;


            Setting_Group Adm = new Setting_Group();
            Adm = db.Setting_Group.Where(w => w.User_Group == "Admin").FirstOrDefault();
            ViewData["Admin"] = Adm;

            Setting_Group Edr = new Setting_Group();
            Edr = db.Setting_Group.Where(w => w.User_Group == "Editor").FirstOrDefault();
            ViewData["Editor"] = Edr;

            Setting_Group Al = new Setting_Group();
            Al = db.Setting_Group.Where(w => w.User_Group == "All").FirstOrDefault();
            ViewData["All"] = Al;

            ViewData["IsEditor"] = IsEditor;
            ViewData["IsAdmin"] = IsAdmin;
            
            if (editWP == null)
            {
                return HttpNotFound();
            }
            db.Aktive_WP.Where(w => w.Id == id).FirstOrDefault().RowVersion = null;
            db.SaveChanges();
            return View(editWP);
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult WP_Edit([Bind(Include = "Id, RowVersion, Color_1, Color_2, Color_3, Color_4, Color_5, Color_6, Column_1, Column_2, Column_3, Column_4, Column_5, Column_6, Column_7, Column_8, Column_9, Column_10, Column_11, Column_12, Column_13, Column_14, Column_15, Column_16, Column_17, Column_18, Column_19, Column_20, Column_21, Column_22, Column_23, Column_24, Column_25, Column_26, Column_27, Column_28, Column_29, Column_30  ")]Aktive_WP awp)
        {
            try
            {
                if (ModelState.IsValid)
                {
                    db.Entry(awp).State = EntityState.Modified;
                    db.SaveChanges();

                    string userString = User.Identity.Name.ToString().Split('\\').Last();    
                    Settings_Filter myFilter = db.Settings_Filter.Where(w => w.User == userString).FirstOrDefault();

                    return RedirectToAction("WP");
                   /* return RedirectToAction("WP", new { Column_1_Filter = myFilter.Filter_Column_1,
                                                        Column_2_Filter = myFilter.Filter_Column_2,
                                                        Column_3_Filter = myFilter.Filter_Column_3,
                                                        Column_4_Filter = myFilter.Filter_Column_4,
                                                        Column_5_Filter = myFilter.Filter_Column_5,
                                                        Column_6_Filter = myFilter.Filter_Column_6,
                                                        Column_7_Filter = myFilter.Filter_Column_7,
                                                        Column_8_Filter = myFilter.Filter_Column_8,
                                                        Column_9_Filter = myFilter.Filter_Column_9,
                                                        Column_10_Filter = myFilter.Filter_Column_10,
                                                        Column_11_Filter = myFilter.Filter_Column_11,
                                                        Column_12_Filter = myFilter.Filter_Column_12,
                                                        Column_13_Filter = myFilter.Filter_Column_13,
                                                        Column_14_Filter = myFilter.Filter_Column_14,
                                                        Column_15_Filter = myFilter.Filter_Column_15,
                                                        Column_16_Filter = myFilter.Filter_Column_16,
                                                        Column_17_Filter = myFilter.Filter_Column_17,
                                                        Column_18_Filter = myFilter.Filter_Column_18,
                                                        Column_19_Filter = myFilter.Filter_Column_19,
                                                        Column_20_Filter = myFilter.Filter_Column_20,
                                                        Column_21_Filter = myFilter.Filter_Column_21,
                                                        Column_22_Filter = myFilter.Filter_Column_22,
                                                        Column_23_Filter = myFilter.Filter_Column_23,
                                                        Column_24_Filter = myFilter.Filter_Column_24,
                                                        Column_25_Filter = myFilter.Filter_Column_25,
                                                        Column_26_Filter = myFilter.Filter_Column_26,
                                                        Column_27_Filter = myFilter.Filter_Column_27,
                                                        Column_28_Filter = myFilter.Filter_Column_28,
                                                        Column_29_Filter = myFilter.Filter_Column_29,
                                                        Column_30_Filter = myFilter.Filter_Column_30 });*/

                }
            }
            catch (DbUpdateConcurrencyException ex)
            {
                var entry = ex.Entries.Single();
                Aktive_WP clientValues = (Aktive_WP)entry.Entity;
                Aktive_WP databaseValues = (Aktive_WP)entry.GetDatabaseValues().ToObject();

                if (databaseValues.Column_1 != clientValues.Column_1)
                    ModelState.AddModelError("Column_1", "Současná hodnota: "
                        + databaseValues.Column_1);
                if (databaseValues.Column_2 != clientValues.Column_2)
                    ModelState.AddModelError("Column_2", "Současná hodnota: "
                        + databaseValues.Column_2);
                if (databaseValues.Column_3 != clientValues.Column_3)
                    ModelState.AddModelError("Column_3", "Současná hodnota: "
                        + databaseValues.Column_3);
                if (databaseValues.Column_4 != clientValues.Column_4)
                    ModelState.AddModelError("Column_4", "Současná hodnota: "
                        + databaseValues.Column_4);
                if (databaseValues.Column_5 != clientValues.Column_5)
                    ModelState.AddModelError("Column_5", "Současná hodnota: "
                        + databaseValues.Column_5);
                if (databaseValues.Column_6 != clientValues.Column_6)
                    ModelState.AddModelError("Column_6", "Současná hodnota: "
                        + databaseValues.Column_6);
                if (databaseValues.Column_7 != clientValues.Column_7)
                    ModelState.AddModelError("Column_7", "Současná hodnota: "
                        + databaseValues.Column_7);
                if (databaseValues.Column_8 != clientValues.Column_8)
                    ModelState.AddModelError("Column_8", "Současná hodnota: "
                        + databaseValues.Column_8);
                if (databaseValues.Column_9 != clientValues.Column_9)
                    ModelState.AddModelError("Column_9", "Současná hodnota: "
                        + databaseValues.Column_9);
                if (databaseValues.Column_10 != clientValues.Column_10)
                    ModelState.AddModelError("Column_10", "Současná hodnota: "
                        + databaseValues.Column_10);
                if (databaseValues.Column_11 != clientValues.Column_11)
                    ModelState.AddModelError("Column_11", "Současná hodnota: "
                        + databaseValues.Column_11);
                if (databaseValues.Column_12 != clientValues.Column_12)
                    ModelState.AddModelError("Column_12", "Současná hodnota: "
                        + databaseValues.Column_12);
                if (databaseValues.Column_13 != clientValues.Column_13)
                    ModelState.AddModelError("Column_13", "Současná hodnota: "
                        + databaseValues.Column_13);
                if (databaseValues.Column_14 != clientValues.Column_14)
                    ModelState.AddModelError("Column_14", "Současná hodnota: "
                        + databaseValues.Column_14);
                if (databaseValues.Column_15 != clientValues.Column_15)
                    ModelState.AddModelError("Column_15", "Současná hodnota: "
                        + databaseValues.Column_15);
                if (databaseValues.Column_16 != clientValues.Column_16)
                    ModelState.AddModelError("Column_16", "Současná hodnota: "
                        + databaseValues.Column_16);
                if (databaseValues.Column_17 != clientValues.Column_17)
                    ModelState.AddModelError("Column_17", "Současná hodnota: "
                        + databaseValues.Column_17);
                if (databaseValues.Column_18 != clientValues.Column_18)
                    ModelState.AddModelError("Column_18", "Současná hodnota: "
                        + databaseValues.Column_18);
                if (databaseValues.Column_19 != clientValues.Column_19)
                    ModelState.AddModelError("Column_19", "Současná hodnota: "
                        + databaseValues.Column_19);
                if (databaseValues.Column_20 != clientValues.Column_20)
                    ModelState.AddModelError("Column_20", "Současná hodnota: "
                        + databaseValues.Column_20);
                if (databaseValues.Column_21 != clientValues.Column_21)
                    ModelState.AddModelError("Column_21", "Současná hodnota: "
                        + databaseValues.Column_21);
                if (databaseValues.Column_22 != clientValues.Column_22)
                    ModelState.AddModelError("Column_22", "Současná hodnota: "
                        + databaseValues.Column_22);
                if (databaseValues.Column_23 != clientValues.Column_23)
                    ModelState.AddModelError("Column_23", "Současná hodnota: "
                        + databaseValues.Column_23);
                if (databaseValues.Column_24 != clientValues.Column_24)
                    ModelState.AddModelError("Column_24", "Současná hodnota: "
                        + databaseValues.Column_24);
                if (databaseValues.Column_25 != clientValues.Column_25)
                    ModelState.AddModelError("Column_25", "Současná hodnota: "
                        + databaseValues.Column_25);
                if (databaseValues.Column_26 != clientValues.Column_26)
                    ModelState.AddModelError("Column_26", "Současná hodnota: "
                        + databaseValues.Column_26);
                if (databaseValues.Column_27 != clientValues.Column_27)
                    ModelState.AddModelError("Column_27", "Současná hodnota: "
                        + databaseValues.Column_27);
                if (databaseValues.Column_28 != clientValues.Column_28)
                    ModelState.AddModelError("Column_28", "Současná hodnota: "
                        + databaseValues.Column_28);
                if (databaseValues.Column_29 != clientValues.Column_29)
                    ModelState.AddModelError("Column_29", "Současná hodnota: "
                        + databaseValues.Column_29);
                if (databaseValues.Column_30 != clientValues.Column_30)
                    ModelState.AddModelError("Column_30", "Současná hodnota: "
                        + databaseValues.Column_30);

                ModelState.AddModelError(string.Empty, "Tento záznam byl v průběhu editace změněn. Pokud kliknete znovu na uložit, data se přepíší");
                awp.RowVersion = databaseValues.RowVersion;
            }
            catch (DataException /* dex */)
            {
                //Log the error (uncomment dex variable name after DataException and add a line here to write a log.
                ModelState.AddModelError(string.Empty, "Unable to save changes. Try again, and if the problem persists contact your system Administrator.");
            }
             return View(awp);
        }


        [HttpGet]
        public ActionResult Delete(int id)
        {
            var deleteWP = db.Aktive_WP.Find(id);
            db.Aktive_WP.Remove(deleteWP);
            db.SaveChanges();
            return RedirectToAction("WP");
        }


        [HttpGet]
        public ActionResult Export()
        {
            MemoryStream ms = new MemoryStream();
            SpreadsheetDocument xl = SpreadsheetDocument.Create(ms, SpreadsheetDocumentType.Workbook);
            WorkbookPart wbp = xl.AddWorkbookPart();
            WorksheetPart wsp = wbp.AddNewPart<WorksheetPart>();
            Workbook wb = new Workbook();
            FileVersion fv = new FileVersion();
            fv.ApplicationName = "Microsoft Office Excel";
            Worksheet ws = new Worksheet();

            UInt32Value i = 1;
            SheetData sd = new SheetData();

            Row r_head = new Row { RowIndex = i };
            i++;

            Cell cellH1 = new Cell();
            cellH1.DataType = CellValues.String;
            cellH1.CellValue = new CellValue(db.Settings_Name_Columns.Where(w => w.User == "All").FirstOrDefault().Name_Column_1);
            r_head.Append(cellH1);

            Cell cellH2 = new Cell();
            cellH2.DataType = CellValues.String;
            cellH2.CellValue = new CellValue(db.Settings_Name_Columns.Where(w => w.User == "All").FirstOrDefault().Name_Column_2);
            r_head.Append(cellH2);

            if (((db.Setting_Group.Where(r => r.User_Group == "All").FirstOrDefault().Column_3_Show) == true) ||
                        (IsAdmin && (db.Setting_Group.Where(r => r.User_Group == "Admin").FirstOrDefault().Column_3_Show) == true) ||
                        (IsEditor && (db.Setting_Group.Where(r => r.User_Group == "Editor").FirstOrDefault().Column_3_Show) == true)
                        )
            {
                Cell cellH3 = new Cell();
                cellH3.DataType = CellValues.String;
                cellH3.CellValue = new CellValue(db.Settings_Name_Columns.Where(w => w.User == "All").FirstOrDefault().Name_Column_3);
                r_head.Append(cellH3);
            }

            if (((db.Setting_Group.Where(r => r.User_Group == "All").FirstOrDefault().Column_4_Show) == true) ||
                        (IsAdmin && (db.Setting_Group.Where(r => r.User_Group == "Admin").FirstOrDefault().Column_4_Show) == true) ||
                        (IsEditor && (db.Setting_Group.Where(r => r.User_Group == "Editor").FirstOrDefault().Column_4_Show) == true)
                        )
            {
                Cell cellH4 = new Cell();
                cellH4.DataType = CellValues.String;
                cellH4.CellValue = new CellValue(db.Settings_Name_Columns.Where(w => w.User == "All").FirstOrDefault().Name_Column_4);
                r_head.Append(cellH4);
            }

            if (((db.Setting_Group.Where(r => r.User_Group == "All").FirstOrDefault().Column_5_Show) == true) ||
                        (IsAdmin && (db.Setting_Group.Where(r => r.User_Group == "Admin").FirstOrDefault().Column_5_Show) == true) ||
                        (IsEditor && (db.Setting_Group.Where(r => r.User_Group == "Editor").FirstOrDefault().Column_5_Show) == true)
                        )
            {
                Cell cellH5 = new Cell();
                cellH5.DataType = CellValues.String;
                cellH5.CellValue = new CellValue(db.Settings_Name_Columns.Where(w => w.User == "All").FirstOrDefault().Name_Column_5);
                r_head.Append(cellH5);
            }

            if (((db.Setting_Group.Where(r => r.User_Group == "All").FirstOrDefault().Column_6_Show) == true) ||
                        (IsAdmin && (db.Setting_Group.Where(r => r.User_Group == "Admin").FirstOrDefault().Column_6_Show) == true) ||
                        (IsEditor && (db.Setting_Group.Where(r => r.User_Group == "Editor").FirstOrDefault().Column_6_Show) == true)
                        )
            {
                Cell cellH6 = new Cell();
                cellH6.DataType = CellValues.String;
                cellH6.CellValue = new CellValue(db.Settings_Name_Columns.Where(w => w.User == "All").FirstOrDefault().Name_Column_6);
                r_head.Append(cellH6);
            }

            if (((db.Setting_Group.Where(r => r.User_Group == "All").FirstOrDefault().Column_7_Show) == true) ||
                        (IsAdmin && (db.Setting_Group.Where(r => r.User_Group == "Admin").FirstOrDefault().Column_7_Show) == true) ||
                        (IsEditor && (db.Setting_Group.Where(r => r.User_Group == "Editor").FirstOrDefault().Column_7_Show) == true)
                        )
            {
                Cell cellH7 = new Cell();
                cellH7.DataType = CellValues.String;
                cellH7.CellValue = new CellValue(db.Settings_Name_Columns.Where(w => w.User == "All").FirstOrDefault().Name_Column_7);
                r_head.Append(cellH7);
            }

            if (((db.Setting_Group.Where(r => r.User_Group == "All").FirstOrDefault().Column_8_Show) == true) ||
                        (IsAdmin && (db.Setting_Group.Where(r => r.User_Group == "Admin").FirstOrDefault().Column_8_Show) == true) ||
                        (IsEditor && (db.Setting_Group.Where(r => r.User_Group == "Editor").FirstOrDefault().Column_8_Show) == true)
                        )
            {
                Cell cellH8 = new Cell();
                cellH8.DataType = CellValues.String;
                cellH8.CellValue = new CellValue(db.Settings_Name_Columns.Where(w => w.User == "All").FirstOrDefault().Name_Column_8);
                r_head.Append(cellH8);
            }

            if (((db.Setting_Group.Where(r => r.User_Group == "All").FirstOrDefault().Column_9_Show) == true) ||
                        (IsAdmin && (db.Setting_Group.Where(r => r.User_Group == "Admin").FirstOrDefault().Column_9_Show) == true) ||
                        (IsEditor && (db.Setting_Group.Where(r => r.User_Group == "Editor").FirstOrDefault().Column_9_Show) == true)
                        )
            {
                Cell cellH9 = new Cell();
                cellH9.DataType = CellValues.String;
                cellH9.CellValue = new CellValue(db.Settings_Name_Columns.Where(w => w.User == "All").FirstOrDefault().Name_Column_9);
                r_head.Append(cellH9);
            }

            if (((db.Setting_Group.Where(r => r.User_Group == "All").FirstOrDefault().Column_10_Show) == true) ||
                        (IsAdmin && (db.Setting_Group.Where(r => r.User_Group == "Admin").FirstOrDefault().Column_10_Show) == true) ||
                        (IsEditor && (db.Setting_Group.Where(r => r.User_Group == "Editor").FirstOrDefault().Column_10_Show) == true)
                        )
            {
                Cell cellH10 = new Cell();
                cellH10.DataType = CellValues.String;
                cellH10.CellValue = new CellValue(db.Settings_Name_Columns.Where(w => w.User == "All").FirstOrDefault().Name_Column_10);
                r_head.Append(cellH10);
            }

            if (((db.Setting_Group.Where(r => r.User_Group == "All").FirstOrDefault().Column_11_Show) == true) ||
                        (IsAdmin && (db.Setting_Group.Where(r => r.User_Group == "Admin").FirstOrDefault().Column_11_Show) == true) ||
                        (IsEditor && (db.Setting_Group.Where(r => r.User_Group == "Editor").FirstOrDefault().Column_11_Show) == true)
                        )
            {
                Cell cellH11 = new Cell();
                cellH11.DataType = CellValues.String;
                cellH11.CellValue = new CellValue(db.Settings_Name_Columns.Where(w => w.User == "All").FirstOrDefault().Name_Column_11);
                r_head.Append(cellH11);
            }

            if (((db.Setting_Group.Where(r => r.User_Group == "All").FirstOrDefault().Column_12_Show) == true) ||
                        (IsAdmin && (db.Setting_Group.Where(r => r.User_Group == "Admin").FirstOrDefault().Column_12_Show) == true) ||
                        (IsEditor && (db.Setting_Group.Where(r => r.User_Group == "Editor").FirstOrDefault().Column_12_Show) == true)
                        )
            {
                Cell cellH12 = new Cell();
                cellH12.DataType = CellValues.String;
                cellH12.CellValue = new CellValue(db.Settings_Name_Columns.Where(w => w.User == "All").FirstOrDefault().Name_Column_12);
                r_head.Append(cellH12);
            }

            if (((db.Setting_Group.Where(r => r.User_Group == "All").FirstOrDefault().Column_13_Show) == true) ||
                        (IsAdmin && (db.Setting_Group.Where(r => r.User_Group == "Admin").FirstOrDefault().Column_13_Show) == true) ||
                        (IsEditor && (db.Setting_Group.Where(r => r.User_Group == "Editor").FirstOrDefault().Column_13_Show) == true)
                        )
            {
                Cell cellH13 = new Cell();
                cellH13.DataType = CellValues.String;
                cellH13.CellValue = new CellValue(db.Settings_Name_Columns.Where(w => w.User == "All").FirstOrDefault().Name_Column_13);
                r_head.Append(cellH13);
            }

            if (((db.Setting_Group.Where(r => r.User_Group == "All").FirstOrDefault().Column_14_Show) == true) ||
                        (IsAdmin && (db.Setting_Group.Where(r => r.User_Group == "Admin").FirstOrDefault().Column_14_Show) == true) ||
                        (IsEditor && (db.Setting_Group.Where(r => r.User_Group == "Editor").FirstOrDefault().Column_14_Show) == true)
                        )
            {
                Cell cellH14 = new Cell();
                cellH14.DataType = CellValues.String;
                cellH14.CellValue = new CellValue(db.Settings_Name_Columns.Where(w => w.User == "All").FirstOrDefault().Name_Column_14);
                r_head.Append(cellH14);
            }

            if (((db.Setting_Group.Where(r => r.User_Group == "All").FirstOrDefault().Column_15_Show) == true) ||
                        (IsAdmin && (db.Setting_Group.Where(r => r.User_Group == "Admin").FirstOrDefault().Column_15_Show) == true) ||
                        (IsEditor && (db.Setting_Group.Where(r => r.User_Group == "Editor").FirstOrDefault().Column_15_Show) == true)
                        )
            {
                Cell cellH15 = new Cell();
                cellH15.DataType = CellValues.String;
                cellH15.CellValue = new CellValue(db.Settings_Name_Columns.Where(w => w.User == "All").FirstOrDefault().Name_Column_15);
                r_head.Append(cellH15);
            }

            if (((db.Setting_Group.Where(r => r.User_Group == "All").FirstOrDefault().Column_16_Show) == true) ||
                        (IsAdmin && (db.Setting_Group.Where(r => r.User_Group == "Admin").FirstOrDefault().Column_16_Show) == true) ||
                        (IsEditor && (db.Setting_Group.Where(r => r.User_Group == "Editor").FirstOrDefault().Column_16_Show) == true)
                        )
            {
                Cell cellH16 = new Cell();
                cellH16.DataType = CellValues.String;
                cellH16.CellValue = new CellValue(db.Settings_Name_Columns.Where(w => w.User == "All").FirstOrDefault().Name_Column_16);
                r_head.Append(cellH16);
            }

            if (((db.Setting_Group.Where(r => r.User_Group == "All").FirstOrDefault().Column_17_Show) == true) ||
                        (IsAdmin && (db.Setting_Group.Where(r => r.User_Group == "Admin").FirstOrDefault().Column_17_Show) == true) ||
                        (IsEditor && (db.Setting_Group.Where(r => r.User_Group == "Editor").FirstOrDefault().Column_17_Show) == true)
                        )
            {
                Cell cellH17 = new Cell();
                cellH17.DataType = CellValues.String;
                cellH17.CellValue = new CellValue(db.Settings_Name_Columns.Where(w => w.User == "All").FirstOrDefault().Name_Column_17);
                r_head.Append(cellH17);
            }

            if (((db.Setting_Group.Where(r => r.User_Group == "All").FirstOrDefault().Column_18_Show) == true) ||
                        (IsAdmin && (db.Setting_Group.Where(r => r.User_Group == "Admin").FirstOrDefault().Column_18_Show) == true) ||
                        (IsEditor && (db.Setting_Group.Where(r => r.User_Group == "Editor").FirstOrDefault().Column_18_Show) == true)
                        )
            {
                Cell cellH18 = new Cell();
                cellH18.DataType = CellValues.String;
                cellH18.CellValue = new CellValue(db.Settings_Name_Columns.Where(w => w.User == "All").FirstOrDefault().Name_Column_18);
                r_head.Append(cellH18);
            }

            if (((db.Setting_Group.Where(r => r.User_Group == "All").FirstOrDefault().Column_19_Show) == true) ||
                        (IsAdmin && (db.Setting_Group.Where(r => r.User_Group == "Admin").FirstOrDefault().Column_19_Show) == true) ||
                        (IsEditor && (db.Setting_Group.Where(r => r.User_Group == "Editor").FirstOrDefault().Column_19_Show) == true)
                        )
            {
                Cell cellH19 = new Cell();
                cellH19.DataType = CellValues.String;
                cellH19.CellValue = new CellValue(db.Settings_Name_Columns.Where(w => w.User == "All").FirstOrDefault().Name_Column_19);
                r_head.Append(cellH19);
            }

            if (((db.Setting_Group.Where(r => r.User_Group == "All").FirstOrDefault().Column_20_Show) == true) ||
                        (IsAdmin && (db.Setting_Group.Where(r => r.User_Group == "Admin").FirstOrDefault().Column_20_Show) == true) ||
                        (IsEditor && (db.Setting_Group.Where(r => r.User_Group == "Editor").FirstOrDefault().Column_20_Show) == true)
                        )
            {
                Cell cellH20 = new Cell();
                cellH20.DataType = CellValues.String;
                cellH20.CellValue = new CellValue(db.Settings_Name_Columns.Where(w => w.User == "All").FirstOrDefault().Name_Column_20);
                r_head.Append(cellH20);
            }

            if (((db.Setting_Group.Where(r => r.User_Group == "All").FirstOrDefault().Column_21_Show) == true) ||
                        (IsAdmin && (db.Setting_Group.Where(r => r.User_Group == "Admin").FirstOrDefault().Column_21_Show) == true) ||
                        (IsEditor && (db.Setting_Group.Where(r => r.User_Group == "Editor").FirstOrDefault().Column_21_Show) == true)
                        )
            {
                Cell cellH21 = new Cell();
                cellH21.DataType = CellValues.String;
                cellH21.CellValue = new CellValue(db.Settings_Name_Columns.Where(w => w.User == "All").FirstOrDefault().Name_Column_21);
                r_head.Append(cellH21);
            }

            if (((db.Setting_Group.Where(r => r.User_Group == "All").FirstOrDefault().Column_22_Show) == true) ||
                        (IsAdmin && (db.Setting_Group.Where(r => r.User_Group == "Admin").FirstOrDefault().Column_22_Show) == true) ||
                        (IsEditor && (db.Setting_Group.Where(r => r.User_Group == "Editor").FirstOrDefault().Column_22_Show) == true)
                        )
            {
                Cell cellH22 = new Cell();
                cellH22.DataType = CellValues.String;
                cellH22.CellValue = new CellValue(db.Settings_Name_Columns.Where(w => w.User == "All").FirstOrDefault().Name_Column_22);
                r_head.Append(cellH22);
            }

            if (((db.Setting_Group.Where(r => r.User_Group == "All").FirstOrDefault().Column_23_Show) == true) ||
                        (IsAdmin && (db.Setting_Group.Where(r => r.User_Group == "Admin").FirstOrDefault().Column_23_Show) == true) ||
                        (IsEditor && (db.Setting_Group.Where(r => r.User_Group == "Editor").FirstOrDefault().Column_23_Show) == true)
                        )
            {
                Cell cellH23 = new Cell();
                cellH23.DataType = CellValues.String;
                cellH23.CellValue = new CellValue(db.Settings_Name_Columns.Where(w => w.User == "All").FirstOrDefault().Name_Column_23);
                r_head.Append(cellH23);
            }

            if (((db.Setting_Group.Where(r => r.User_Group == "All").FirstOrDefault().Column_24_Show) == true) ||
                        (IsAdmin && (db.Setting_Group.Where(r => r.User_Group == "Admin").FirstOrDefault().Column_24_Show) == true) ||
                        (IsEditor && (db.Setting_Group.Where(r => r.User_Group == "Editor").FirstOrDefault().Column_24_Show) == true)
                        )
            {
                Cell cellH24 = new Cell();
                cellH24.DataType = CellValues.String;
                cellH24.CellValue = new CellValue(db.Settings_Name_Columns.Where(w => w.User == "All").FirstOrDefault().Name_Column_24);
                r_head.Append(cellH24);
            }

            if (((db.Setting_Group.Where(r => r.User_Group == "All").FirstOrDefault().Column_25_Show) == true) ||
                        (IsAdmin && (db.Setting_Group.Where(r => r.User_Group == "Admin").FirstOrDefault().Column_25_Show) == true) ||
                        (IsEditor && (db.Setting_Group.Where(r => r.User_Group == "Editor").FirstOrDefault().Column_25_Show) == true)
                        )
            {
                Cell cellH25 = new Cell();
                cellH25.DataType = CellValues.String;
                cellH25.CellValue = new CellValue(db.Settings_Name_Columns.Where(w => w.User == "All").FirstOrDefault().Name_Column_25);
                r_head.Append(cellH25);
            }

            if (((db.Setting_Group.Where(r => r.User_Group == "All").FirstOrDefault().Column_26_Show) == true) ||
                        (IsAdmin && (db.Setting_Group.Where(r => r.User_Group == "Admin").FirstOrDefault().Column_26_Show) == true) ||
                        (IsEditor && (db.Setting_Group.Where(r => r.User_Group == "Editor").FirstOrDefault().Column_26_Show) == true)
                        )
            {
                Cell cellH26 = new Cell();
                cellH26.DataType = CellValues.String;
                cellH26.CellValue = new CellValue(db.Settings_Name_Columns.Where(w => w.User == "All").FirstOrDefault().Name_Column_26);
                r_head.Append(cellH26);
            }

            if (((db.Setting_Group.Where(r => r.User_Group == "All").FirstOrDefault().Column_27_Show) == true) ||
                        (IsAdmin && (db.Setting_Group.Where(r => r.User_Group == "Admin").FirstOrDefault().Column_27_Show) == true) ||
                        (IsEditor && (db.Setting_Group.Where(r => r.User_Group == "Editor").FirstOrDefault().Column_27_Show) == true)
                        )
            {
                Cell cellH27 = new Cell();
                cellH27.DataType = CellValues.String;
                cellH27.CellValue = new CellValue(db.Settings_Name_Columns.Where(w => w.User == "All").FirstOrDefault().Name_Column_27);
                r_head.Append(cellH27);
            }

            if (((db.Setting_Group.Where(r => r.User_Group == "All").FirstOrDefault().Column_28_Show) == true) ||
                        (IsAdmin && (db.Setting_Group.Where(r => r.User_Group == "Admin").FirstOrDefault().Column_28_Show) == true) ||
                        (IsEditor && (db.Setting_Group.Where(r => r.User_Group == "Editor").FirstOrDefault().Column_28_Show) == true)
                        )
            {
                Cell cellH28 = new Cell();
                cellH28.DataType = CellValues.String;
                cellH28.CellValue = new CellValue(db.Settings_Name_Columns.Where(w => w.User == "All").FirstOrDefault().Name_Column_28);
                r_head.Append(cellH28);
            }

            if (((db.Setting_Group.Where(r => r.User_Group == "All").FirstOrDefault().Column_29_Show) == true) ||
                        (IsAdmin && (db.Setting_Group.Where(r => r.User_Group == "Admin").FirstOrDefault().Column_29_Show) == true) ||
                        (IsEditor && (db.Setting_Group.Where(r => r.User_Group == "Editor").FirstOrDefault().Column_29_Show) == true)
                        )
            {
                Cell cellH29 = new Cell();
                cellH29.DataType = CellValues.String;
                cellH29.CellValue = new CellValue(db.Settings_Name_Columns.Where(w => w.User == "All").FirstOrDefault().Name_Column_29);
                r_head.Append(cellH29);
            }

            if (((db.Setting_Group.Where(r => r.User_Group == "All").FirstOrDefault().Column_30_Show) == true) ||
                        (IsAdmin && (db.Setting_Group.Where(r => r.User_Group == "Admin").FirstOrDefault().Column_30_Show) == true) ||
                        (IsEditor && (db.Setting_Group.Where(r => r.User_Group == "Editor").FirstOrDefault().Column_30_Show) == true)
                        )
            {
                Cell cellH30 = new Cell();
                cellH30.DataType = CellValues.String;
                cellH30.CellValue = new CellValue(db.Settings_Name_Columns.Where(w => w.User == "All").FirstOrDefault().Name_Column_30);
                r_head.Append(cellH30);
            }

            sd.Append(r_head);

            //    WP_ViewModel vm = new WP_ViewModel { };
            //  vm.HandleRequest();
            string userString = User.Identity.Name.ToString().Split('\\').Last();    
            Settings_Filter myFilter = db.Settings_Filter.Where(w => w.User == userString).FirstOrDefault();
            List<Aktive_WP> array = db.Aktive_WP.ToList();
            string ColumnValueString = "";
            if ((myFilter.Filter_Column_1 == null ? 0 : myFilter.Filter_Column_1.Count()) > 0)
            {
                List<Aktive_WP> array_Filter = new List<Aktive_WP>();
                ColumnValueString = myFilter.Filter_Column_1;
                string[] tokes = ColumnValueString.Split(',');
                int[] IdArray = Array.ConvertAll<string, int>(tokes, int.Parse);
                for (int n = 0; n < IdArray.Count(); n++)
                {
                    int a = IdArray[n];
                    string u = db.Aktive_WP.Where(x => x.Id == a).FirstOrDefault().Column_1;
                    if (u == null || u == "")
                    {
                        array_Filter.AddRange(db.Aktive_WP.Where(s => s.Column_1 == null).DefaultIfEmpty().ToList());
                        array_Filter.AddRange(db.Aktive_WP.Where(s => s.Column_1 == "").DefaultIfEmpty().ToList());
                    }
                    else
                    {
                        array_Filter.AddRange(db.Aktive_WP.Where(s => s.Column_1 == u).DefaultIfEmpty().ToList());
                    }
                }
                List<Aktive_WP> k = array.Intersect(array_Filter).ToList();
                array = k;
            }
            if ((myFilter.Filter_Column_2 == null ? 0 : myFilter.Filter_Column_2.Count()) > 0)
            {
                List<Aktive_WP> array_Filter = new List<Aktive_WP>();
                ColumnValueString = myFilter.Filter_Column_2;
                string[] tokes = ColumnValueString.Split(',');
                int[] IdArray = Array.ConvertAll<string, int>(tokes, int.Parse);
                for (int n = 0; n < IdArray.Count(); n++)
                {
                    int a = IdArray[n];
                    string u = db.Aktive_WP.Where(x => x.Id == a).FirstOrDefault().Column_2;
                    if (u == null || u == "")
                    {
                        array_Filter.AddRange(db.Aktive_WP.Where(s => s.Column_2 == null).DefaultIfEmpty().ToList());
                        array_Filter.AddRange(db.Aktive_WP.Where(s => s.Column_2 == "").DefaultIfEmpty().ToList());
                    }
                    else
                    {
                        array_Filter.AddRange(db.Aktive_WP.Where(s => s.Column_2 == u).DefaultIfEmpty().ToList());
                    }
                }
                List<Aktive_WP> k = array.Intersect(array_Filter).ToList();
                array = k;
            }
            if ((myFilter.Filter_Column_3 == null ? 0 : myFilter.Filter_Column_3.Count()) > 0)
            {
                List<Aktive_WP> array_Filter = new List<Aktive_WP>();
                ColumnValueString = myFilter.Filter_Column_3;
                string[] tokes = ColumnValueString.Split(',');
                int[] IdArray = Array.ConvertAll<string, int>(tokes, int.Parse);
                for (int n = 0; n < IdArray.Count(); n++)
                {
                    int a = IdArray[n];
                    string u = db.Aktive_WP.Where(x => x.Id == a).FirstOrDefault().Column_3;
                    if (u == null || u == "")
                    {
                        array_Filter.AddRange(db.Aktive_WP.Where(s => s.Column_3 == null).DefaultIfEmpty().ToList());
                        array_Filter.AddRange(db.Aktive_WP.Where(s => s.Column_3 == "").DefaultIfEmpty().ToList());
                    }
                    else
                    {
                        array_Filter.AddRange(db.Aktive_WP.Where(s => s.Column_3 == u).DefaultIfEmpty().ToList());
                    }
                }
                List<Aktive_WP> k = array.Intersect(array_Filter).ToList();
                array = k;
            }
            if ((myFilter.Filter_Column_4 == null ? 0 : myFilter.Filter_Column_4.Count()) > 0)
            {
                List<Aktive_WP> array_Filter = new List<Aktive_WP>();
                ColumnValueString = myFilter.Filter_Column_4;
                string[] tokes = ColumnValueString.Split(',');
                int[] IdArray = Array.ConvertAll<string, int>(tokes, int.Parse);
                for (int n = 0; n < IdArray.Count(); n++)
                {
                    int a = IdArray[n];
                    string u = db.Aktive_WP.Where(x => x.Id == a).FirstOrDefault().Column_4;
                    if (u == null || u == "")
                    {
                        array_Filter.AddRange(db.Aktive_WP.Where(s => s.Column_4 == null).DefaultIfEmpty().ToList());
                        array_Filter.AddRange(db.Aktive_WP.Where(s => s.Column_4 == "").DefaultIfEmpty().ToList());
                    }
                    else
                    {
                        array_Filter.AddRange(db.Aktive_WP.Where(s => s.Column_4 == u).DefaultIfEmpty().ToList());
                    }
                }
                List<Aktive_WP> k = array.Intersect(array_Filter).ToList();
                array = k;
            }
            if ((myFilter.Filter_Column_5 == null ? 0 : myFilter.Filter_Column_5.Count()) > 0)
            {
                List<Aktive_WP> array_Filter = new List<Aktive_WP>();
                ColumnValueString = myFilter.Filter_Column_5;
                string[] tokes = ColumnValueString.Split(',');
                int[] IdArray = Array.ConvertAll<string, int>(tokes, int.Parse);
                for (int n = 0; n < IdArray.Count(); n++)
                {
                    int a = IdArray[n];
                    string u = db.Aktive_WP.Where(x => x.Id == a).FirstOrDefault().Column_5;
                    if (u == null || u == "")
                    {
                        array_Filter.AddRange(db.Aktive_WP.Where(s => s.Column_5 == null).DefaultIfEmpty().ToList());
                        array_Filter.AddRange(db.Aktive_WP.Where(s => s.Column_5 == "").DefaultIfEmpty().ToList());
                    }
                    else
                    {
                        array_Filter.AddRange(db.Aktive_WP.Where(s => s.Column_5 == u).DefaultIfEmpty().ToList());
                    }
                }
                List<Aktive_WP> k = array.Intersect(array_Filter).ToList();
                array = k;
            }
            if ((myFilter.Filter_Column_6 == null ? 0 : myFilter.Filter_Column_6.Count()) > 0)
            {
                List<Aktive_WP> array_Filter = new List<Aktive_WP>();
                ColumnValueString = myFilter.Filter_Column_6;
                string[] tokes = ColumnValueString.Split(',');
                int[] IdArray = Array.ConvertAll<string, int>(tokes, int.Parse);
                for (int n = 0; n < IdArray.Count(); n++)
                {
                    int a = IdArray[n];
                    string u = db.Aktive_WP.Where(x => x.Id == a).FirstOrDefault().Column_6;
                    if (u == null || u == "")
                    {
                        array_Filter.AddRange(db.Aktive_WP.Where(s => s.Column_6 == null).DefaultIfEmpty().ToList());
                        array_Filter.AddRange(db.Aktive_WP.Where(s => s.Column_6 == "").DefaultIfEmpty().ToList());
                    }
                    else
                    {
                        array_Filter.AddRange(db.Aktive_WP.Where(s => s.Column_6 == u).DefaultIfEmpty().ToList());
                    }
                }
                List<Aktive_WP> k = array.Intersect(array_Filter).ToList();
                array = k;
            }
            if ((myFilter.Filter_Column_7 == null ? 0 : myFilter.Filter_Column_7.Count()) > 0)
            {
                List<Aktive_WP> array_Filter = new List<Aktive_WP>();
                ColumnValueString = myFilter.Filter_Column_7;
                string[] tokes = ColumnValueString.Split(',');
                int[] IdArray = Array.ConvertAll<string, int>(tokes, int.Parse);
                for (int n = 0; n < IdArray.Count(); n++)
                {
                    int a = IdArray[n];
                    string u = db.Aktive_WP.Where(x => x.Id == a).FirstOrDefault().Column_7;
                    if (u == null || u == "")
                    {
                        array_Filter.AddRange(db.Aktive_WP.Where(s => s.Column_7 == null).DefaultIfEmpty().ToList());
                        array_Filter.AddRange(db.Aktive_WP.Where(s => s.Column_7 == "").DefaultIfEmpty().ToList());
                    }
                    else
                    {
                        array_Filter.AddRange(db.Aktive_WP.Where(s => s.Column_7 == u).DefaultIfEmpty().ToList());
                    }
                }
                List<Aktive_WP> k = array.Intersect(array_Filter).ToList();
                array = k;
            }
            if ((myFilter.Filter_Column_8 == null ? 0 : myFilter.Filter_Column_8.Count()) > 0)
            {
                List<Aktive_WP> array_Filter = new List<Aktive_WP>();
                ColumnValueString = myFilter.Filter_Column_8;
                string[] tokes = ColumnValueString.Split(',');
                int[] IdArray = Array.ConvertAll<string, int>(tokes, int.Parse);
                for (int n = 0; n < IdArray.Count(); n++)
                {
                    int a = IdArray[n];
                    string u = db.Aktive_WP.Where(x => x.Id == a).FirstOrDefault().Column_8;
                    if (u == null || u == "")
                    {
                        array_Filter.AddRange(db.Aktive_WP.Where(s => s.Column_8 == null).DefaultIfEmpty().ToList());
                        array_Filter.AddRange(db.Aktive_WP.Where(s => s.Column_8 == "").DefaultIfEmpty().ToList());
                    }
                    else
                    {
                        array_Filter.AddRange(db.Aktive_WP.Where(s => s.Column_8 == u).DefaultIfEmpty().ToList());
                    }
                }
                List<Aktive_WP> k = array.Intersect(array_Filter).ToList();
                array = k;
            }
            if ((myFilter.Filter_Column_9 == null ? 0 : myFilter.Filter_Column_9.Count()) > 0)
            {
                List<Aktive_WP> array_Filter = new List<Aktive_WP>();
                ColumnValueString = myFilter.Filter_Column_9;
                string[] tokes = ColumnValueString.Split(',');
                int[] IdArray = Array.ConvertAll<string, int>(tokes, int.Parse);
                for (int n = 0; n < IdArray.Count(); n++)
                {
                    int a = IdArray[n];
                    string u = db.Aktive_WP.Where(x => x.Id == a).FirstOrDefault().Column_9;
                    if (u == null || u == "")
                    {
                        array_Filter.AddRange(db.Aktive_WP.Where(s => s.Column_9 == null).DefaultIfEmpty().ToList());
                        array_Filter.AddRange(db.Aktive_WP.Where(s => s.Column_9 == "").DefaultIfEmpty().ToList());
                    }
                    else
                    {
                        array_Filter.AddRange(db.Aktive_WP.Where(s => s.Column_9 == u).DefaultIfEmpty().ToList());
                    }
                }
                List<Aktive_WP> k = array.Intersect(array_Filter).ToList();
                array = k;
            }
            if ((myFilter.Filter_Column_10 == null ? 0 : myFilter.Filter_Column_10.Count()) > 0)
            {
                List<Aktive_WP> array_Filter = new List<Aktive_WP>();
                ColumnValueString = myFilter.Filter_Column_10;
                string[] tokes = ColumnValueString.Split(',');
                int[] IdArray = Array.ConvertAll<string, int>(tokes, int.Parse);
                for (int n = 0; n < IdArray.Count(); n++)
                {
                    int a = IdArray[n];
                    string u = db.Aktive_WP.Where(x => x.Id == a).FirstOrDefault().Column_10;
                    if (u == null || u == "")
                    {
                        array_Filter.AddRange(db.Aktive_WP.Where(s => s.Column_10 == null).DefaultIfEmpty().ToList());
                        array_Filter.AddRange(db.Aktive_WP.Where(s => s.Column_10 == "").DefaultIfEmpty().ToList());
                    }
                    else
                    {
                        array_Filter.AddRange(db.Aktive_WP.Where(s => s.Column_10 == u).DefaultIfEmpty().ToList());
                    }
                }
                List<Aktive_WP> k = array.Intersect(array_Filter).ToList();
                array = k;
            }
            if ((myFilter.Filter_Column_11 == null ? 0 : myFilter.Filter_Column_11.Count()) > 0)
            {
                List<Aktive_WP> array_Filter = new List<Aktive_WP>();
                ColumnValueString = myFilter.Filter_Column_11;
                string[] tokes = ColumnValueString.Split(',');
                int[] IdArray = Array.ConvertAll<string, int>(tokes, int.Parse);
                for (int n = 0; n < IdArray.Count(); n++)
                {
                    int a = IdArray[n];
                    string u = db.Aktive_WP.Where(x => x.Id == a).FirstOrDefault().Column_11;
                    if (u == null || u == "")
                    {
                        array_Filter.AddRange(db.Aktive_WP.Where(s => s.Column_11 == null).DefaultIfEmpty().ToList());
                        array_Filter.AddRange(db.Aktive_WP.Where(s => s.Column_11 == "").DefaultIfEmpty().ToList());
                    }
                    else
                    {
                        array_Filter.AddRange(db.Aktive_WP.Where(s => s.Column_11 == u).DefaultIfEmpty().ToList());
                    }
                }
                List<Aktive_WP> k = array.Intersect(array_Filter).ToList();
                array = k;
            }
            if ((myFilter.Filter_Column_12 == null ? 0 : myFilter.Filter_Column_12.Count()) > 0)
            {
                List<Aktive_WP> array_Filter = new List<Aktive_WP>();
                ColumnValueString = myFilter.Filter_Column_12;
                string[] tokes = ColumnValueString.Split(',');
                int[] IdArray = Array.ConvertAll<string, int>(tokes, int.Parse);
                for (int n = 0; n < IdArray.Count(); n++)
                {
                    int a = IdArray[n];
                    string u = db.Aktive_WP.Where(x => x.Id == a).FirstOrDefault().Column_12;
                    if (u == null || u == "")
                    {
                        array_Filter.AddRange(db.Aktive_WP.Where(s => s.Column_12 == null).DefaultIfEmpty().ToList());
                        array_Filter.AddRange(db.Aktive_WP.Where(s => s.Column_12 == "").DefaultIfEmpty().ToList());
                    }
                    else
                    {
                        array_Filter.AddRange(db.Aktive_WP.Where(s => s.Column_12 == u).DefaultIfEmpty().ToList());
                    }
                }
                List<Aktive_WP> k = array.Intersect(array_Filter).ToList();
                array = k;
            }
            if ((myFilter.Filter_Column_13 == null ? 0 : myFilter.Filter_Column_13.Count()) > 0)
            {
                List<Aktive_WP> array_Filter = new List<Aktive_WP>();
                ColumnValueString = myFilter.Filter_Column_13;
                string[] tokes = ColumnValueString.Split(',');
                int[] IdArray = Array.ConvertAll<string, int>(tokes, int.Parse);
                for (int n = 0; n < IdArray.Count(); n++)
                {
                    int a = IdArray[n];
                    string u = db.Aktive_WP.Where(x => x.Id == a).FirstOrDefault().Column_13;
                    if (u == null || u == "")
                    {
                        array_Filter.AddRange(db.Aktive_WP.Where(s => s.Column_13 == null).DefaultIfEmpty().ToList());
                        array_Filter.AddRange(db.Aktive_WP.Where(s => s.Column_13 == "").DefaultIfEmpty().ToList());
                    }
                    else
                    {
                        array_Filter.AddRange(db.Aktive_WP.Where(s => s.Column_13 == u).DefaultIfEmpty().ToList());
                    }
                }
                List<Aktive_WP> k = array.Intersect(array_Filter).ToList();
                array = k;
            }
            if ((myFilter.Filter_Column_14 == null ? 0 : myFilter.Filter_Column_14.Count()) > 0)
            {
                List<Aktive_WP> array_Filter = new List<Aktive_WP>();
                ColumnValueString = myFilter.Filter_Column_14;
                string[] tokes = ColumnValueString.Split(',');
                int[] IdArray = Array.ConvertAll<string, int>(tokes, int.Parse);
                for (int n = 0; n < IdArray.Count(); n++)
                {
                    int a = IdArray[n];
                    string u = db.Aktive_WP.Where(x => x.Id == a).FirstOrDefault().Column_14;
                    if (u == null || u == "")
                    {
                        array_Filter.AddRange(db.Aktive_WP.Where(s => s.Column_14 == null).DefaultIfEmpty().ToList());
                        array_Filter.AddRange(db.Aktive_WP.Where(s => s.Column_14 == "").DefaultIfEmpty().ToList());
                    }
                    else
                    {
                        array_Filter.AddRange(db.Aktive_WP.Where(s => s.Column_14 == u).DefaultIfEmpty().ToList());
                    }
                }
                List<Aktive_WP> k = array.Intersect(array_Filter).ToList();
                array = k;
            }
            if ((myFilter.Filter_Column_15 == null ? 0 : myFilter.Filter_Column_15.Count()) > 0)
            {
                List<Aktive_WP> array_Filter = new List<Aktive_WP>();
                ColumnValueString = myFilter.Filter_Column_15;
                string[] tokes = ColumnValueString.Split(',');
                int[] IdArray = Array.ConvertAll<string, int>(tokes, int.Parse);
                for (int n = 0; n < IdArray.Count(); n++)
                {
                    int a = IdArray[n];
                    string u = db.Aktive_WP.Where(x => x.Id == a).FirstOrDefault().Column_15;
                    if (u == null || u == "")
                    {
                        array_Filter.AddRange(db.Aktive_WP.Where(s => s.Column_15 == null).DefaultIfEmpty().ToList());
                        array_Filter.AddRange(db.Aktive_WP.Where(s => s.Column_15 == "").DefaultIfEmpty().ToList());
                    }
                    else
                    {
                        array_Filter.AddRange(db.Aktive_WP.Where(s => s.Column_15 == u).DefaultIfEmpty().ToList());
                    }
                }
                List<Aktive_WP> k = array.Intersect(array_Filter).ToList();
                array = k;
            }
            if ((myFilter.Filter_Column_16 == null ? 0 : myFilter.Filter_Column_16.Count()) > 0)
            {
                List<Aktive_WP> array_Filter = new List<Aktive_WP>();
                ColumnValueString = myFilter.Filter_Column_16;
                string[] tokes = ColumnValueString.Split(',');
                int[] IdArray = Array.ConvertAll<string, int>(tokes, int.Parse);
                for (int n = 0; n < IdArray.Count(); n++)
                {
                    int a = IdArray[n];
                    string u = db.Aktive_WP.Where(x => x.Id == a).FirstOrDefault().Column_16;
                    if (u == null || u == "")
                    {
                        array_Filter.AddRange(db.Aktive_WP.Where(s => s.Column_16 == null).DefaultIfEmpty().ToList());
                        array_Filter.AddRange(db.Aktive_WP.Where(s => s.Column_16 == "").DefaultIfEmpty().ToList());
                    }
                    else
                    {
                        array_Filter.AddRange(db.Aktive_WP.Where(s => s.Column_16 == u).DefaultIfEmpty().ToList());
                    }
                }
                List<Aktive_WP> k = array.Intersect(array_Filter).ToList();
                array = k;
            }
            if ((myFilter.Filter_Column_17 == null ? 0 : myFilter.Filter_Column_17.Count()) > 0)
            {
                List<Aktive_WP> array_Filter = new List<Aktive_WP>();
                ColumnValueString = myFilter.Filter_Column_17;
                string[] tokes = ColumnValueString.Split(',');
                int[] IdArray = Array.ConvertAll<string, int>(tokes, int.Parse);
                for (int n = 0; n < IdArray.Count(); n++)
                {
                    int a = IdArray[n];
                    string u = db.Aktive_WP.Where(x => x.Id == a).FirstOrDefault().Column_17;
                    if (u == null || u == "")
                    {
                        array_Filter.AddRange(db.Aktive_WP.Where(s => s.Column_17 == null).DefaultIfEmpty().ToList());
                        array_Filter.AddRange(db.Aktive_WP.Where(s => s.Column_17 == "").DefaultIfEmpty().ToList());
                    }
                    else
                    {
                        array_Filter.AddRange(db.Aktive_WP.Where(s => s.Column_17 == u).DefaultIfEmpty().ToList());
                    }
                }
                List<Aktive_WP> k = array.Intersect(array_Filter).ToList();
                array = k;
            }
            if ((myFilter.Filter_Column_18 == null ? 0 : myFilter.Filter_Column_18.Count()) > 0)
            {
                List<Aktive_WP> array_Filter = new List<Aktive_WP>();
                ColumnValueString = myFilter.Filter_Column_18;
                string[] tokes = ColumnValueString.Split(',');
                int[] IdArray = Array.ConvertAll<string, int>(tokes, int.Parse);
                for (int n = 0; n < IdArray.Count(); n++)
                {
                    int a = IdArray[n];
                    string u = db.Aktive_WP.Where(x => x.Id == a).FirstOrDefault().Column_18;
                    if (u == null || u == "")
                    {
                        array_Filter.AddRange(db.Aktive_WP.Where(s => s.Column_18 == null).DefaultIfEmpty().ToList());
                        array_Filter.AddRange(db.Aktive_WP.Where(s => s.Column_18 == "").DefaultIfEmpty().ToList());
                    }
                    else
                    {
                        array_Filter.AddRange(db.Aktive_WP.Where(s => s.Column_18 == u).DefaultIfEmpty().ToList());
                    }
                }
                List<Aktive_WP> k = array.Intersect(array_Filter).ToList();
                array = k;
            }
            if ((myFilter.Filter_Column_19 == null ? 0 : myFilter.Filter_Column_19.Count()) > 0)
            {
                List<Aktive_WP> array_Filter = new List<Aktive_WP>();
                ColumnValueString = myFilter.Filter_Column_19;
                string[] tokes = ColumnValueString.Split(',');
                int[] IdArray = Array.ConvertAll<string, int>(tokes, int.Parse);
                for (int n = 0; n < IdArray.Count(); n++)
                {
                    int a = IdArray[n];
                    string u = db.Aktive_WP.Where(x => x.Id == a).FirstOrDefault().Column_19;
                    if (u == null || u == "")
                    {
                        array_Filter.AddRange(db.Aktive_WP.Where(s => s.Column_19 == null).DefaultIfEmpty().ToList());
                        array_Filter.AddRange(db.Aktive_WP.Where(s => s.Column_19 == "").DefaultIfEmpty().ToList());
                    }
                    else
                    {
                        array_Filter.AddRange(db.Aktive_WP.Where(s => s.Column_19 == u).DefaultIfEmpty().ToList());
                    }
                }
                List<Aktive_WP> k = array.Intersect(array_Filter).ToList();
                array = k;
            }
            if ((myFilter.Filter_Column_20 == null ? 0 : myFilter.Filter_Column_20.Count()) > 0)
            {
                List<Aktive_WP> array_Filter = new List<Aktive_WP>();
                ColumnValueString = myFilter.Filter_Column_20;
                string[] tokes = ColumnValueString.Split(',');
                int[] IdArray = Array.ConvertAll<string, int>(tokes, int.Parse);
                for (int n = 0; n < IdArray.Count(); n++)
                {
                    int a = IdArray[n];
                    string u = db.Aktive_WP.Where(x => x.Id == a).FirstOrDefault().Column_20;
                    if (u == null || u == "")
                    {
                        array_Filter.AddRange(db.Aktive_WP.Where(s => s.Column_20 == null).DefaultIfEmpty().ToList());
                        array_Filter.AddRange(db.Aktive_WP.Where(s => s.Column_20 == "").DefaultIfEmpty().ToList());
                    }
                    else
                    {
                        array_Filter.AddRange(db.Aktive_WP.Where(s => s.Column_20 == u).DefaultIfEmpty().ToList());
                    }
                }
                List<Aktive_WP> k = array.Intersect(array_Filter).ToList();
                array = k;
            }
            if ((myFilter.Filter_Column_21 == null ? 0 : myFilter.Filter_Column_21.Count()) > 0)
            {
                List<Aktive_WP> array_Filter = new List<Aktive_WP>();
                ColumnValueString = myFilter.Filter_Column_21;
                string[] tokes = ColumnValueString.Split(',');
                int[] IdArray = Array.ConvertAll<string, int>(tokes, int.Parse);
                for (int n = 0; n < IdArray.Count(); n++)
                {
                    int a = IdArray[n];
                    string u = db.Aktive_WP.Where(x => x.Id == a).FirstOrDefault().Column_21;
                    if (u == null || u == "")
                    {
                        array_Filter.AddRange(db.Aktive_WP.Where(s => s.Column_21 == null).DefaultIfEmpty().ToList());
                        array_Filter.AddRange(db.Aktive_WP.Where(s => s.Column_21 == "").DefaultIfEmpty().ToList());
                    }
                    else
                    {
                        array_Filter.AddRange(db.Aktive_WP.Where(s => s.Column_21 == u).DefaultIfEmpty().ToList());
                    }
                }
                List<Aktive_WP> k = array.Intersect(array_Filter).ToList();
                array = k;
            }
            if ((myFilter.Filter_Column_22 == null ? 0 : myFilter.Filter_Column_22.Count()) > 0)
            {
                List<Aktive_WP> array_Filter = new List<Aktive_WP>();
                ColumnValueString = myFilter.Filter_Column_22;
                string[] tokes = ColumnValueString.Split(',');
                int[] IdArray = Array.ConvertAll<string, int>(tokes, int.Parse);
                for (int n = 0; n < IdArray.Count(); n++)
                {
                    int a = IdArray[n];
                    string u = db.Aktive_WP.Where(x => x.Id == a).FirstOrDefault().Column_22;
                    if (u == null || u == "")
                    {
                        array_Filter.AddRange(db.Aktive_WP.Where(s => s.Column_22 == null).DefaultIfEmpty().ToList());
                        array_Filter.AddRange(db.Aktive_WP.Where(s => s.Column_22 == "").DefaultIfEmpty().ToList());
                    }
                    else
                    {
                        array_Filter.AddRange(db.Aktive_WP.Where(s => s.Column_22 == u).DefaultIfEmpty().ToList());
                    }
                }
                List<Aktive_WP> k = array.Intersect(array_Filter).ToList();
                array = k;
            }
            if ((myFilter.Filter_Column_23 == null ? 0 : myFilter.Filter_Column_23.Count()) > 0)
            {
                List<Aktive_WP> array_Filter = new List<Aktive_WP>();
                ColumnValueString = myFilter.Filter_Column_23;
                string[] tokes = ColumnValueString.Split(',');
                int[] IdArray = Array.ConvertAll<string, int>(tokes, int.Parse);
                for (int n = 0; n < IdArray.Count(); n++)
                {
                    int a = IdArray[n];
                    string u = db.Aktive_WP.Where(x => x.Id == a).FirstOrDefault().Column_23;
                    if (u == null || u == "")
                    {
                        array_Filter.AddRange(db.Aktive_WP.Where(s => s.Column_23 == null).DefaultIfEmpty().ToList());
                        array_Filter.AddRange(db.Aktive_WP.Where(s => s.Column_23 == "").DefaultIfEmpty().ToList());
                    }
                    else
                    {
                        array_Filter.AddRange(db.Aktive_WP.Where(s => s.Column_23 == u).DefaultIfEmpty().ToList());
                    }
                }
                List<Aktive_WP> k = array.Intersect(array_Filter).ToList();
                array = k;
            }
            if ((myFilter.Filter_Column_24 == null ? 0 : myFilter.Filter_Column_24.Count()) > 0)
            {
                List<Aktive_WP> array_Filter = new List<Aktive_WP>();
                ColumnValueString = myFilter.Filter_Column_24;
                string[] tokes = ColumnValueString.Split(',');
                int[] IdArray = Array.ConvertAll<string, int>(tokes, int.Parse);
                for (int n = 0; n < IdArray.Count(); n++)
                {
                    int a = IdArray[n];
                    string u = db.Aktive_WP.Where(x => x.Id == a).FirstOrDefault().Column_24;
                    if (u == null || u == "")
                    {
                        array_Filter.AddRange(db.Aktive_WP.Where(s => s.Column_24 == null).DefaultIfEmpty().ToList());
                        array_Filter.AddRange(db.Aktive_WP.Where(s => s.Column_24 == "").DefaultIfEmpty().ToList());
                    }
                    else
                    {
                        array_Filter.AddRange(db.Aktive_WP.Where(s => s.Column_24 == u).DefaultIfEmpty().ToList());
                    }
                }
                List<Aktive_WP> k = array.Intersect(array_Filter).ToList();
                array = k;
            }
            if ((myFilter.Filter_Column_25 == null ? 0 : myFilter.Filter_Column_25.Count()) > 0)
            {
                List<Aktive_WP> array_Filter = new List<Aktive_WP>();
                ColumnValueString = myFilter.Filter_Column_25;
                string[] tokes = ColumnValueString.Split(',');
                int[] IdArray = Array.ConvertAll<string, int>(tokes, int.Parse);
                for (int n = 0; n < IdArray.Count(); n++)
                {
                    int a = IdArray[n];
                    string u = db.Aktive_WP.Where(x => x.Id == a).FirstOrDefault().Column_25;
                    if (u == null || u == "")
                    {
                        array_Filter.AddRange(db.Aktive_WP.Where(s => s.Column_25 == null).DefaultIfEmpty().ToList());
                        array_Filter.AddRange(db.Aktive_WP.Where(s => s.Column_25 == "").DefaultIfEmpty().ToList());
                    }
                    else
                    {
                        array_Filter.AddRange(db.Aktive_WP.Where(s => s.Column_25 == u).DefaultIfEmpty().ToList());
                    }
                }
                List<Aktive_WP> k = array.Intersect(array_Filter).ToList();
                array = k;
            }
            if ((myFilter.Filter_Column_26 == null ? 0 : myFilter.Filter_Column_26.Count()) > 0)
            {
                List<Aktive_WP> array_Filter = new List<Aktive_WP>();
                ColumnValueString = myFilter.Filter_Column_26;
                string[] tokes = ColumnValueString.Split(',');
                int[] IdArray = Array.ConvertAll<string, int>(tokes, int.Parse);
                for (int n = 0; n < IdArray.Count(); n++)
                {
                    int a = IdArray[n];
                    string u = db.Aktive_WP.Where(x => x.Id == a).FirstOrDefault().Column_26;
                    if (u == null || u == "")
                    {
                        array_Filter.AddRange(db.Aktive_WP.Where(s => s.Column_26 == null).DefaultIfEmpty().ToList());
                        array_Filter.AddRange(db.Aktive_WP.Where(s => s.Column_26 == "").DefaultIfEmpty().ToList());
                    }
                    else
                    {
                        array_Filter.AddRange(db.Aktive_WP.Where(s => s.Column_26 == u).DefaultIfEmpty().ToList());
                    }
                }
                List<Aktive_WP> k = array.Intersect(array_Filter).ToList();
                array = k;
            }
            if ((myFilter.Filter_Column_27 == null ? 0 : myFilter.Filter_Column_27.Count()) > 0)
            {
                List<Aktive_WP> array_Filter = new List<Aktive_WP>();
                ColumnValueString = myFilter.Filter_Column_27;
                string[] tokes = ColumnValueString.Split(',');
                int[] IdArray = Array.ConvertAll<string, int>(tokes, int.Parse);
                for (int n = 0; n < IdArray.Count(); n++)
                {
                    int a = IdArray[n];
                    string u = db.Aktive_WP.Where(x => x.Id == a).FirstOrDefault().Column_27;
                    if (u == null || u == "")
                    {
                        array_Filter.AddRange(db.Aktive_WP.Where(s => s.Column_27 == null).DefaultIfEmpty().ToList());
                        array_Filter.AddRange(db.Aktive_WP.Where(s => s.Column_27 == "").DefaultIfEmpty().ToList());
                    }
                    else
                    {
                        array_Filter.AddRange(db.Aktive_WP.Where(s => s.Column_27 == u).DefaultIfEmpty().ToList());
                    }
                }
                List<Aktive_WP> k = array.Intersect(array_Filter).ToList();
                array = k;
            }
            if ((myFilter.Filter_Column_28 == null ? 0 : myFilter.Filter_Column_28.Count()) > 0)
            {
                List<Aktive_WP> array_Filter = new List<Aktive_WP>();
                ColumnValueString = myFilter.Filter_Column_28;
                string[] tokes = ColumnValueString.Split(',');
                int[] IdArray = Array.ConvertAll<string, int>(tokes, int.Parse);
                for (int n = 0; n < IdArray.Count(); n++)
                {
                    int a = IdArray[n];
                    string u = db.Aktive_WP.Where(x => x.Id == a).FirstOrDefault().Column_29;
                    if (u == null || u == "")
                    {
                        array_Filter.AddRange(db.Aktive_WP.Where(s => s.Column_28 == null).DefaultIfEmpty().ToList());
                        array_Filter.AddRange(db.Aktive_WP.Where(s => s.Column_28 == "").DefaultIfEmpty().ToList());
                    }
                    else
                    {
                        array_Filter.AddRange(db.Aktive_WP.Where(s => s.Column_28 == u).DefaultIfEmpty().ToList());
                    }
                }
                List<Aktive_WP> k = array.Intersect(array_Filter).ToList();
                array = k;
            }
            if ((myFilter.Filter_Column_29 == null ? 0 : myFilter.Filter_Column_29.Count()) > 0)
            {
                List<Aktive_WP> array_Filter = new List<Aktive_WP>();
                ColumnValueString = myFilter.Filter_Column_29;
                string[] tokes = ColumnValueString.Split(',');
                int[] IdArray = Array.ConvertAll<string, int>(tokes, int.Parse);
                for (int n = 0; n < IdArray.Count(); n++)
                {
                    int a = IdArray[n];
                    string u = db.Aktive_WP.Where(x => x.Id == a).FirstOrDefault().Column_29;
                    if (u == null || u == "")
                    {
                        array_Filter.AddRange(db.Aktive_WP.Where(s => s.Column_29 == null).DefaultIfEmpty().ToList());
                        array_Filter.AddRange(db.Aktive_WP.Where(s => s.Column_29 == "").DefaultIfEmpty().ToList());
                    }
                    else
                    {
                        array_Filter.AddRange(db.Aktive_WP.Where(s => s.Column_29 == u).DefaultIfEmpty().ToList());
                    }
                }
                List<Aktive_WP> k = array.Intersect(array_Filter).ToList();
                array = k;
            }
            if ((myFilter.Filter_Column_30 == null ? 0 : myFilter.Filter_Column_30.Count()) > 0)
            {
                List<Aktive_WP> array_Filter = new List<Aktive_WP>();
                ColumnValueString = myFilter.Filter_Column_30;
                string[] tokes = ColumnValueString.Split(',');
                int[] IdArray = Array.ConvertAll<string, int>(tokes, int.Parse);
                for (int n = 0; n < IdArray.Count(); n++)
                {
                    int a = IdArray[n];
                    string u = db.Aktive_WP.Where(x => x.Id == a).FirstOrDefault().Column_30;
                    if (u == null || u == "")
                    {
                        array_Filter.AddRange(db.Aktive_WP.Where(s => s.Column_30 == null).DefaultIfEmpty().ToList());
                        array_Filter.AddRange(db.Aktive_WP.Where(s => s.Column_30 == "").DefaultIfEmpty().ToList());
                    }
                    else
                    {
                        array_Filter.AddRange(db.Aktive_WP.Where(s => s.Column_30 == u).DefaultIfEmpty().ToList());
                    }
                }
                List<Aktive_WP> k = array.Intersect(array_Filter).ToList();
                array = k;
            }
            List<Aktive_WP> result = array.Distinct().ToList();




            foreach (var item in result)
            {
                Row r = new Row { RowIndex = i };
                i++;

                Cell cell1 = new Cell();
                cell1.DataType = CellValues.String;
                cell1.CellValue = new CellValue(item.Column_1);
                r.Append(cell1);

                Cell cell2 = new Cell();
                cell2.DataType = CellValues.String;
                cell2.CellValue = new CellValue(item.Column_2);
                r.Append(cell2);

                if (((db.Setting_Group.Where(k => k.User_Group == "All").FirstOrDefault().Column_3_Show) == true) ||
                        (IsAdmin && (db.Setting_Group.Where(k => k.User_Group == "Admin").FirstOrDefault().Column_3_Show) == true) ||
                        (IsEditor && (db.Setting_Group.Where(k => k.User_Group == "Editor").FirstOrDefault().Column_3_Show) == true)
                        )
                {
                    Cell cell3 = new Cell();
                    cell3.DataType = CellValues.String;
                    cell3.CellValue = new CellValue(item.Column_3);
                    r.Append(cell3);
                }

                if (((db.Setting_Group.Where(k => k.User_Group == "All").FirstOrDefault().Column_4_Show) == true) ||
                        (IsAdmin && (db.Setting_Group.Where(k => k.User_Group == "Admin").FirstOrDefault().Column_4_Show) == true) ||
                        (IsEditor && (db.Setting_Group.Where(k => k.User_Group == "Editor").FirstOrDefault().Column_4_Show) == true)
                        )
                {
                    Cell cell4 = new Cell();
                    cell4.DataType = CellValues.String;
                    cell4.CellValue = new CellValue(item.Column_4);
                    r.Append(cell4);
                }

                if (((db.Setting_Group.Where(k => k.User_Group == "All").FirstOrDefault().Column_5_Show) == true) ||
                        (IsAdmin && (db.Setting_Group.Where(k => k.User_Group == "Admin").FirstOrDefault().Column_5_Show) == true) ||
                        (IsEditor && (db.Setting_Group.Where(k => k.User_Group == "Editor").FirstOrDefault().Column_5_Show) == true)
                        )
                {
                    Cell cell5 = new Cell();
                    cell5.DataType = CellValues.String;
                    cell5.CellValue = new CellValue(item.Column_5);
                    r.Append(cell5);
                }

                if (((db.Setting_Group.Where(k => k.User_Group == "All").FirstOrDefault().Column_6_Show) == true) ||
                        (IsAdmin && (db.Setting_Group.Where(k => k.User_Group == "Admin").FirstOrDefault().Column_6_Show) == true) ||
                        (IsEditor && (db.Setting_Group.Where(k => k.User_Group == "Editor").FirstOrDefault().Column_6_Show) == true)
                        )
                {
                    Cell cell6 = new Cell();
                    cell6.DataType = CellValues.String;
                    cell6.CellValue = new CellValue(item.Column_6);
                    r.Append(cell6);
                }

                if (((db.Setting_Group.Where(k => k.User_Group == "All").FirstOrDefault().Column_7_Show) == true) ||
                        (IsAdmin && (db.Setting_Group.Where(k => k.User_Group == "Admin").FirstOrDefault().Column_7_Show) == true) ||
                        (IsEditor && (db.Setting_Group.Where(k => k.User_Group == "Editor").FirstOrDefault().Column_7_Show) == true)
                        )
                {
                    Cell cell7 = new Cell();
                    cell7.DataType = CellValues.String;
                    cell7.CellValue = new CellValue(item.Column_7);
                    r.Append(cell7);
                }

                if (((db.Setting_Group.Where(k => k.User_Group == "All").FirstOrDefault().Column_8_Show) == true) ||
                        (IsAdmin && (db.Setting_Group.Where(k => k.User_Group == "Admin").FirstOrDefault().Column_8_Show) == true) ||
                        (IsEditor && (db.Setting_Group.Where(k => k.User_Group == "Editor").FirstOrDefault().Column_8_Show) == true)
                        )
                {
                    Cell cell8 = new Cell();
                    cell8.DataType = CellValues.String;
                    cell8.CellValue = new CellValue(item.Column_8);
                    r.Append(cell8);
                }

                if (((db.Setting_Group.Where(k => k.User_Group == "All").FirstOrDefault().Column_9_Show) == true) ||
                        (IsAdmin && (db.Setting_Group.Where(k => k.User_Group == "Admin").FirstOrDefault().Column_9_Show) == true) ||
                        (IsEditor && (db.Setting_Group.Where(k => k.User_Group == "Editor").FirstOrDefault().Column_9_Show) == true)
                        )
                {
                    Cell cell9 = new Cell();
                    cell9.DataType = CellValues.String;
                    cell9.CellValue = new CellValue(item.Column_9);
                    r.Append(cell9);
                }

                if (((db.Setting_Group.Where(k => k.User_Group == "All").FirstOrDefault().Column_10_Show) == true) ||
                        (IsAdmin && (db.Setting_Group.Where(k => k.User_Group == "Admin").FirstOrDefault().Column_10_Show) == true) ||
                        (IsEditor && (db.Setting_Group.Where(k => k.User_Group == "Editor").FirstOrDefault().Column_10_Show) == true)
                        )
                {
                    Cell cell10 = new Cell();
                    cell10.DataType = CellValues.String;
                    cell10.CellValue = new CellValue(item.Column_10);
                    r.Append(cell10);
                }

                if (((db.Setting_Group.Where(k => k.User_Group == "All").FirstOrDefault().Column_11_Show) == true) ||
                        (IsAdmin && (db.Setting_Group.Where(k => k.User_Group == "Admin").FirstOrDefault().Column_11_Show) == true) ||
                        (IsEditor && (db.Setting_Group.Where(k => k.User_Group == "Editor").FirstOrDefault().Column_11_Show) == true)
                        )
                {
                    Cell cell11 = new Cell();
                    cell11.DataType = CellValues.String;
                    cell11.CellValue = new CellValue(item.Column_11);
                    r.Append(cell11);
                }

                if (((db.Setting_Group.Where(k => k.User_Group == "All").FirstOrDefault().Column_12_Show) == true) ||
                        (IsAdmin && (db.Setting_Group.Where(k => k.User_Group == "Admin").FirstOrDefault().Column_12_Show) == true) ||
                        (IsEditor && (db.Setting_Group.Where(k => k.User_Group == "Editor").FirstOrDefault().Column_12_Show) == true)
                        )
                {
                    Cell cell12 = new Cell();
                    cell12.DataType = CellValues.String;
                    cell12.CellValue = new CellValue(item.Column_12);
                    r.Append(cell12);
                }

                if (((db.Setting_Group.Where(k => k.User_Group == "All").FirstOrDefault().Column_13_Show) == true) ||
                        (IsAdmin && (db.Setting_Group.Where(k => k.User_Group == "Admin").FirstOrDefault().Column_13_Show) == true) ||
                        (IsEditor && (db.Setting_Group.Where(k => k.User_Group == "Editor").FirstOrDefault().Column_13_Show) == true)
                        )
                {
                    Cell cell13 = new Cell();
                    cell13.DataType = CellValues.String;
                    cell13.CellValue = new CellValue(item.Column_13);
                    r.Append(cell13);
                }

                if (((db.Setting_Group.Where(k => k.User_Group == "All").FirstOrDefault().Column_14_Show) == true) ||
                        (IsAdmin && (db.Setting_Group.Where(k => k.User_Group == "Admin").FirstOrDefault().Column_14_Show) == true) ||
                        (IsEditor && (db.Setting_Group.Where(k => k.User_Group == "Editor").FirstOrDefault().Column_14_Show) == true)
                        )
                {
                    Cell cell14 = new Cell();
                    cell14.DataType = CellValues.String;
                    cell14.CellValue = new CellValue(item.Column_14);
                    r.Append(cell14);
                }

                if (((db.Setting_Group.Where(k => k.User_Group == "All").FirstOrDefault().Column_15_Show) == true) ||
                        (IsAdmin && (db.Setting_Group.Where(k => k.User_Group == "Admin").FirstOrDefault().Column_15_Show) == true) ||
                        (IsEditor && (db.Setting_Group.Where(k => k.User_Group == "Editor").FirstOrDefault().Column_15_Show) == true)
                        )
                {
                    Cell cell15 = new Cell();
                    cell15.DataType = CellValues.String;
                    cell15.CellValue = new CellValue(item.Column_15);
                    r.Append(cell15);
                }

                if (((db.Setting_Group.Where(k => k.User_Group == "All").FirstOrDefault().Column_16_Show) == true) ||
                        (IsAdmin && (db.Setting_Group.Where(k => k.User_Group == "Admin").FirstOrDefault().Column_16_Show) == true) ||
                        (IsEditor && (db.Setting_Group.Where(k => k.User_Group == "Editor").FirstOrDefault().Column_16_Show) == true)
                        )
                {
                    Cell cell16 = new Cell();
                    cell16.DataType = CellValues.String;
                    cell16.CellValue = new CellValue(item.Column_16);
                    r.Append(cell16);
                }

                if (((db.Setting_Group.Where(k => k.User_Group == "All").FirstOrDefault().Column_17_Show) == true) ||
                        (IsAdmin && (db.Setting_Group.Where(k => k.User_Group == "Admin").FirstOrDefault().Column_17_Show) == true) ||
                        (IsEditor && (db.Setting_Group.Where(k => k.User_Group == "Editor").FirstOrDefault().Column_17_Show) == true)
                        )
                {
                    Cell cell17 = new Cell();
                    cell17.DataType = CellValues.String;
                    cell17.CellValue = new CellValue(item.Column_17);
                    r.Append(cell17);
                }

                if (((db.Setting_Group.Where(k => k.User_Group == "All").FirstOrDefault().Column_18_Show) == true) ||
                        (IsAdmin && (db.Setting_Group.Where(k => k.User_Group == "Admin").FirstOrDefault().Column_18_Show) == true) ||
                        (IsEditor && (db.Setting_Group.Where(k => k.User_Group == "Editor").FirstOrDefault().Column_18_Show) == true)
                        )
                {
                    Cell cell18 = new Cell();
                    cell18.DataType = CellValues.String;
                    cell18.CellValue = new CellValue(item.Column_18);
                    r.Append(cell18);
                }

                if (((db.Setting_Group.Where(k => k.User_Group == "All").FirstOrDefault().Column_19_Show) == true) ||
                        (IsAdmin && (db.Setting_Group.Where(k => k.User_Group == "Admin").FirstOrDefault().Column_19_Show) == true) ||
                        (IsEditor && (db.Setting_Group.Where(k => k.User_Group == "Editor").FirstOrDefault().Column_19_Show) == true)
                        )
                {
                    Cell cell19 = new Cell();
                    cell19.DataType = CellValues.String;
                    cell19.CellValue = new CellValue(item.Column_19);
                    r.Append(cell19);
                }

                if (((db.Setting_Group.Where(k => k.User_Group == "All").FirstOrDefault().Column_20_Show) == true) ||
                        (IsAdmin && (db.Setting_Group.Where(k => k.User_Group == "Admin").FirstOrDefault().Column_20_Show) == true) ||
                        (IsEditor && (db.Setting_Group.Where(k => k.User_Group == "Editor").FirstOrDefault().Column_20_Show) == true)
                        )
                {
                    Cell cell20 = new Cell();
                    cell20.DataType = CellValues.String;
                    cell20.CellValue = new CellValue(item.Column_20);
                    r.Append(cell20);
                }

                if (((db.Setting_Group.Where(k => k.User_Group == "All").FirstOrDefault().Column_21_Show) == true) ||
                        (IsAdmin && (db.Setting_Group.Where(k => k.User_Group == "Admin").FirstOrDefault().Column_21_Show) == true) ||
                        (IsEditor && (db.Setting_Group.Where(k => k.User_Group == "Editor").FirstOrDefault().Column_21_Show) == true)
                        )
                {
                    Cell cell21 = new Cell();
                    cell21.DataType = CellValues.String;
                    cell21.CellValue = new CellValue(item.Column_21);
                    r.Append(cell21);
                }

                if (((db.Setting_Group.Where(k => k.User_Group == "All").FirstOrDefault().Column_22_Show) == true) ||
                        (IsAdmin && (db.Setting_Group.Where(k => k.User_Group == "Admin").FirstOrDefault().Column_22_Show) == true) ||
                        (IsEditor && (db.Setting_Group.Where(k => k.User_Group == "Editor").FirstOrDefault().Column_22_Show) == true)
                        )
                {
                    Cell cell22 = new Cell();
                    cell22.DataType = CellValues.String;
                    cell22.CellValue = new CellValue(item.Column_22);
                    r.Append(cell22);
                }

                if (((db.Setting_Group.Where(k => k.User_Group == "All").FirstOrDefault().Column_23_Show) == true) ||
                        (IsAdmin && (db.Setting_Group.Where(k => k.User_Group == "Admin").FirstOrDefault().Column_23_Show) == true) ||
                        (IsEditor && (db.Setting_Group.Where(k => k.User_Group == "Editor").FirstOrDefault().Column_23_Show) == true)
                        )
                {
                    Cell cell23 = new Cell();
                    cell23.DataType = CellValues.String;
                    cell23.CellValue = new CellValue(item.Column_23);
                    r.Append(cell23);
                }

                if (((db.Setting_Group.Where(k => k.User_Group == "All").FirstOrDefault().Column_24_Show) == true) ||
                        (IsAdmin && (db.Setting_Group.Where(k => k.User_Group == "Admin").FirstOrDefault().Column_24_Show) == true) ||
                        (IsEditor && (db.Setting_Group.Where(k => k.User_Group == "Editor").FirstOrDefault().Column_24_Show) == true)
                        )
                {
                    Cell cell24 = new Cell();
                    cell24.DataType = CellValues.String;
                    cell24.CellValue = new CellValue(item.Column_24);
                    r.Append(cell24);
                }

                if (((db.Setting_Group.Where(k => k.User_Group == "All").FirstOrDefault().Column_25_Show) == true) ||
                        (IsAdmin && (db.Setting_Group.Where(k => k.User_Group == "Admin").FirstOrDefault().Column_25_Show) == true) ||
                        (IsEditor && (db.Setting_Group.Where(k => k.User_Group == "Editor").FirstOrDefault().Column_25_Show) == true)
                        )
                {
                    Cell cell25 = new Cell();
                    cell25.DataType = CellValues.String;
                    cell25.CellValue = new CellValue(item.Column_25);
                    r.Append(cell25);
                }

                if (((db.Setting_Group.Where(k => k.User_Group == "All").FirstOrDefault().Column_26_Show) == true) ||
                        (IsAdmin && (db.Setting_Group.Where(k => k.User_Group == "Admin").FirstOrDefault().Column_26_Show) == true) ||
                        (IsEditor && (db.Setting_Group.Where(k => k.User_Group == "Editor").FirstOrDefault().Column_26_Show) == true)
                        )
                {
                    Cell cell26 = new Cell();
                    cell26.DataType = CellValues.String;
                    cell26.CellValue = new CellValue(item.Column_26);
                    r.Append(cell26);
                }

                if (((db.Setting_Group.Where(k => k.User_Group == "All").FirstOrDefault().Column_27_Show) == true) ||
                        (IsAdmin && (db.Setting_Group.Where(k => k.User_Group == "Admin").FirstOrDefault().Column_27_Show) == true) ||
                        (IsEditor && (db.Setting_Group.Where(k => k.User_Group == "Editor").FirstOrDefault().Column_27_Show) == true)
                        )
                {
                    Cell cell27 = new Cell();
                    cell27.DataType = CellValues.String;
                    cell27.CellValue = new CellValue(item.Column_27);
                    r.Append(cell27);
                }

                if (((db.Setting_Group.Where(k => k.User_Group == "All").FirstOrDefault().Column_28_Show) == true) ||
                        (IsAdmin && (db.Setting_Group.Where(k => k.User_Group == "Admin").FirstOrDefault().Column_28_Show) == true) ||
                        (IsEditor && (db.Setting_Group.Where(k => k.User_Group == "Editor").FirstOrDefault().Column_28_Show) == true)
                        )
                {
                    Cell cell28 = new Cell();
                    cell28.DataType = CellValues.String;
                    cell28.CellValue = new CellValue(item.Column_28);
                    r.Append(cell28);
                }

                if (((db.Setting_Group.Where(k => k.User_Group == "All").FirstOrDefault().Column_29_Show) == true) ||
                        (IsAdmin && (db.Setting_Group.Where(k => k.User_Group == "Admin").FirstOrDefault().Column_29_Show) == true) ||
                        (IsEditor && (db.Setting_Group.Where(k => k.User_Group == "Editor").FirstOrDefault().Column_29_Show) == true)
                        )
                {
                    Cell cell29 = new Cell();
                    cell29.DataType = CellValues.String;
                    cell29.CellValue = new CellValue(item.Column_29);
                    r.Append(cell29);
                }

                if (((db.Setting_Group.Where(k => k.User_Group == "All").FirstOrDefault().Column_30_Show) == true) ||
                        (IsAdmin && (db.Setting_Group.Where(k => k.User_Group == "Admin").FirstOrDefault().Column_30_Show) == true) ||
                        (IsEditor && (db.Setting_Group.Where(k => k.User_Group == "Editor").FirstOrDefault().Column_30_Show) == true)
                        )
                {
                    Cell cell30 = new Cell();
                    cell30.DataType = CellValues.String;
                    cell30.CellValue = new CellValue(item.Column_30);
                    r.Append(cell30);
                }

                sd.Append(r);
            }

            ws.Append(sd);
            wsp.Worksheet = ws;
            wsp.Worksheet.Save();
            Sheets sheets = new Sheets();
            Sheet sheet = new Sheet();
            sheet.Name = "TA aktivní";
            sheet.SheetId = 1;
            sheet.Id = wbp.GetIdOfPart(wsp);
            sheets.Append(sheet);
            wb.Append(fv);
            wb.Append(sheets);

            xl.WorkbookPart.Workbook = wb;
            xl.WorkbookPart.Workbook.Save();
            xl.Close();
            string fileName = "MANTA-" + DateTime.Now.ToString("yyyy_MM_dd HH.mm.ss") + ".xlsx";
            Response.Clear();
            byte[] dt = ms.ToArray();

            Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
            Response.AddHeader("Content-Disposition", string.Format("attachment; filename={0}", fileName));
            Response.BinaryWrite(dt);
            Response.End();
            return RedirectToAction("WP");
        }
        
        public JsonResult GetColumn1List (string searchTerm)
        {
            var ColumnList = db.Aktive_WP.ToList();

            if (searchTerm != null)
            {
                ColumnList = db.Aktive_WP.Where(x => x.Column_1.Contains(searchTerm)).ToList();
            }
            var modifiedData = ColumnList.Select(x => new
            {
                id = x.Id,
                text = x.Column_1
            }).ToList().GroupBy(p => p.text).Select(x => x.First()).OrderBy(x => x.text).ToList();
            return Json(modifiedData, JsonRequestBehavior.AllowGet);
        }
        public JsonResult GetColumn2List(string searchTerm)
        {
            var ColumnList = db.Aktive_WP.ToList();

            if (searchTerm != null)
            {
                ColumnList = db.Aktive_WP.Where(x => x.Column_2.Contains(searchTerm)).ToList();
            }
            var modifiedData = ColumnList.Select(x => new
            {
                id = x.Id,
                text = x.Column_2
            }).ToList().GroupBy(p => p.text).Select(x => x.First()).OrderBy(x => x.text).ToList();
            return Json(modifiedData, JsonRequestBehavior.AllowGet);
        }
        public JsonResult GetColumn3List(string searchTerm)
        {
            var ColumnList = db.Aktive_WP.ToList();

            if (searchTerm != null)
            {
                ColumnList = db.Aktive_WP.Where(x => x.Column_3.Contains(searchTerm)).ToList();
            }
            var modifiedData = ColumnList.Select(x => new
            {
                id = x.Id,
                text = x.Column_3
            }).ToList().GroupBy(p => p.text).Select(x => x.First()).OrderBy(x => x.text).ToList();
            return Json(modifiedData, JsonRequestBehavior.AllowGet);
        }
        public JsonResult GetColumn4List(string searchTerm)
        {
            var ColumnList = db.Aktive_WP.ToList();

            if (searchTerm != null)
            {
                ColumnList = db.Aktive_WP.Where(x => x.Column_4.Contains(searchTerm)).ToList();
            }
            var modifiedData = ColumnList.Select(x => new
            {
                id = x.Id,
                text = x.Column_4
            }).ToList().GroupBy(p => p.text).Select(x => x.First()).OrderBy(x => x.text).ToList();
            return Json(modifiedData, JsonRequestBehavior.AllowGet);
        }
        public JsonResult GetColumn5List(string searchTerm)
        {
            var ColumnList = db.Aktive_WP.ToList();

            if (searchTerm != null)
            {
                ColumnList = db.Aktive_WP.Where(x => x.Column_5.Contains(searchTerm)).ToList();
            }
            var modifiedData = ColumnList.Select(x => new
            {
                id = x.Id,
                text = x.Column_5
            }).ToList().GroupBy(p => p.text).Select(x => x.First()).OrderBy(x => x.text).ToList();
            return Json(modifiedData, JsonRequestBehavior.AllowGet);
        }
        public JsonResult GetColumn6List(string searchTerm)
        {
            var ColumnList = db.Aktive_WP.ToList();

            if (searchTerm != null)
            {
                ColumnList = db.Aktive_WP.Where(x => x.Column_6.Contains(searchTerm)).ToList();
            }
            var modifiedData = ColumnList.Select(x => new
            {
                id = x.Id,
                text = x.Column_6
            }).ToList().GroupBy(p => p.text).Select(x => x.First()).OrderBy(x => x.text).ToList();
            return Json(modifiedData, JsonRequestBehavior.AllowGet);
        }
        public JsonResult GetColumn7List(string searchTerm)
        {
            var ColumnList = db.Aktive_WP.ToList();

            if (searchTerm != null)
            {
                ColumnList = db.Aktive_WP.Where(x => x.Column_7.Contains(searchTerm)).ToList();
            }
            var modifiedData = ColumnList.Select(x => new
            {
                id = x.Id,
                text = x.Column_7
            }).ToList().GroupBy(p => p.text).Select(x => x.First()).OrderBy(x => x.text).ToList();
            //var modifiedData = ColumnList.GroupBy(p => p.Column_7).Select(x => x.First()).OrderBy(x => x.text).ToList();
            return Json(modifiedData, JsonRequestBehavior.AllowGet);
        }
        public JsonResult GetColumn8List(string searchTerm)
        {
            var ColumnList = db.Aktive_WP.ToList();

            if (searchTerm != null)
            {
                ColumnList = db.Aktive_WP.Where(x => x.Column_8.Contains(searchTerm)).ToList();
            }
            var modifiedData = ColumnList.Select(x => new
            {
                id = x.Id,
                text = x.Column_8
            }).ToList().GroupBy(p => p.text).Select(x => x.First()).OrderBy(x => x.text).ToList();
            return Json(modifiedData, JsonRequestBehavior.AllowGet);
        }
        public JsonResult GetColumn9List(string searchTerm)
        {
            var ColumnList = db.Aktive_WP.ToList();

            if (searchTerm != null)
            {
                ColumnList = db.Aktive_WP.Where(x => x.Column_9.Contains(searchTerm)).ToList();
            }
            var modifiedData = ColumnList.Select(x => new
            {
                id = x.Id,
                text = x.Column_9
            }).ToList().GroupBy(p => p.text).Select(x => x.First()).OrderBy(x => x.text).ToList();
            return Json(modifiedData, JsonRequestBehavior.AllowGet);
        }
        public JsonResult GetColumn10List(string searchTerm)
        {
            var ColumnList = db.Aktive_WP.ToList();

            if (searchTerm != null)
            {
                ColumnList = db.Aktive_WP.Where(x => x.Column_10.Contains(searchTerm)).ToList();
            }
            var modifiedData = ColumnList.Select(x => new
            {
                id = x.Id,
                text = x.Column_10
            }).ToList().GroupBy(p => p.text).Select(x => x.First()).OrderBy(x => x.text).ToList();
            return Json(modifiedData, JsonRequestBehavior.AllowGet);
        }
        public JsonResult GetColumn11List(string searchTerm)
        {
            var ColumnList = db.Aktive_WP.ToList();

            if (searchTerm != null)
            {
                ColumnList = db.Aktive_WP.Where(x => x.Column_11.Contains(searchTerm)).ToList();
            }
            var modifiedData = ColumnList.Select(x => new
            {
                id = x.Id,
                text = x.Column_11
            }).ToList().GroupBy(p => p.text).Select(x => x.First()).OrderBy(x => x.text).ToList();
            return Json(modifiedData, JsonRequestBehavior.AllowGet);
        }
        public JsonResult GetColumn12List(string searchTerm)
        {
            var ColumnList = db.Aktive_WP.ToList();

            if (searchTerm != null)
            {
                ColumnList = db.Aktive_WP.Where(x => x.Column_12.Contains(searchTerm)).ToList();
            }
            var modifiedData = ColumnList.Select(x => new
            {
                id = x.Id,
                text = x.Column_12
            }).ToList().GroupBy(p => p.text).Select(x => x.First()).OrderBy(x => x.text).ToList();
            return Json(modifiedData, JsonRequestBehavior.AllowGet);
        }
        public JsonResult GetColumn13List(string searchTerm)
        {
            var ColumnList = db.Aktive_WP.ToList();

            if (searchTerm != null)
            {
                ColumnList = db.Aktive_WP.Where(x => x.Column_13.Contains(searchTerm)).ToList();
            }
            var modifiedData = ColumnList.Select(x => new
            {
                id = x.Id,
                text = x.Column_13
            }).ToList().GroupBy(p => p.text).Select(x => x.First()).OrderBy(x => x.text).ToList();
            return Json(modifiedData, JsonRequestBehavior.AllowGet);
        }
        public JsonResult GetColumn14List(string searchTerm)
        {
            var ColumnList = db.Aktive_WP.ToList();

            if (searchTerm != null)
            {
                ColumnList = db.Aktive_WP.Where(x => x.Column_14.Contains(searchTerm)).ToList();
            }
            var modifiedData = ColumnList.Select(x => new
            {
                id = x.Id,
                text = x.Column_14
            }).ToList().GroupBy(p => p.text).Select(x => x.First()).OrderBy(x => x.text).ToList();
            return Json(modifiedData, JsonRequestBehavior.AllowGet);
        }
        public JsonResult GetColumn15List(string searchTerm)
        {
            var ColumnList = db.Aktive_WP.ToList();

            if (searchTerm != null)
            {
                ColumnList = db.Aktive_WP.Where(x => x.Column_15.Contains(searchTerm)).ToList();
            }
            var modifiedData = ColumnList.Select(x => new
            {
                id = x.Id,
                text = x.Column_15
            }).ToList().GroupBy(p => p.text).Select(x => x.First()).OrderBy(x => x.text).ToList();
            return Json(modifiedData, JsonRequestBehavior.AllowGet);
        }
        public JsonResult GetColumn16List(string searchTerm)
        {
            var ColumnList = db.Aktive_WP.ToList();

            if (searchTerm != null)
            {
                ColumnList = db.Aktive_WP.Where(x => x.Column_16.Contains(searchTerm)).ToList();
            }
            var modifiedData = ColumnList.Select(x => new
            {
                id = x.Id,
                text = x.Column_16
            }).ToList().GroupBy(p => p.text).Select(x => x.First()).OrderBy(x => x.text).ToList();
            return Json(modifiedData, JsonRequestBehavior.AllowGet);
        }
        public JsonResult GetColumn17List(string searchTerm)
        {
            var ColumnList = db.Aktive_WP.ToList();

            if (searchTerm != null)
            {
                ColumnList = db.Aktive_WP.Where(x => x.Column_17.Contains(searchTerm)).ToList();
            }
            var modifiedData = ColumnList.Select(x => new
            {
                id = x.Id,
                text = x.Column_17
            }).ToList().GroupBy(p => p.text).Select(x => x.First()).OrderBy(x => x.text).ToList();
            return Json(modifiedData, JsonRequestBehavior.AllowGet);
        }
        public JsonResult GetColumn18List(string searchTerm)
        {
            var ColumnList = db.Aktive_WP.ToList();

            if (searchTerm != null)
            {
                ColumnList = db.Aktive_WP.Where(x => x.Column_18.Contains(searchTerm)).ToList();
            }
            var modifiedData = ColumnList.Select(x => new
            {
                id = x.Id,
                text = x.Column_18
            }).ToList().GroupBy(p => p.text).Select(x => x.First()).OrderBy(x => x.text).ToList();
            return Json(modifiedData, JsonRequestBehavior.AllowGet);
        }
        public JsonResult GetColumn19List(string searchTerm)
        {
            var ColumnList = db.Aktive_WP.ToList();

            if (searchTerm != null)
            {
                ColumnList = db.Aktive_WP.Where(x => x.Column_19.Contains(searchTerm)).ToList();
            }
            var modifiedData = ColumnList.Select(x => new
            {
                id = x.Id,
                text = x.Column_19
            }).ToList().GroupBy(p => p.text).Select(x => x.First()).OrderBy(x => x.text).ToList();
            return Json(modifiedData, JsonRequestBehavior.AllowGet);
        }
        public JsonResult GetColumn20List(string searchTerm)
        {
            var ColumnList = db.Aktive_WP.ToList();

            if (searchTerm != null)
            {
                ColumnList = db.Aktive_WP.Where(x => x.Column_20.Contains(searchTerm)).ToList();
            }
            var modifiedData = ColumnList.Select(x => new
            {
                id = x.Id,
                text = x.Column_20
            }).ToList().GroupBy(p => p.text).Select(x => x.First()).OrderBy(x => x.text).ToList();
            return Json(modifiedData, JsonRequestBehavior.AllowGet);
        }
        public JsonResult GetColumn21List(string searchTerm)
        {
            var ColumnList = db.Aktive_WP.ToList();

            if (searchTerm != null)
            {
                ColumnList = db.Aktive_WP.Where(x => x.Column_21.Contains(searchTerm)).ToList();
            }
            var modifiedData = ColumnList.Select(x => new
            {
                id = x.Id,
                text = x.Column_21
            }).ToList().GroupBy(p => p.text).Select(x => x.First()).OrderBy(x => x.text).ToList();
            return Json(modifiedData, JsonRequestBehavior.AllowGet);
        }
        public JsonResult GetColumn22List(string searchTerm)
        {
            var ColumnList = db.Aktive_WP.ToList();

            if (searchTerm != null)
            {
                ColumnList = db.Aktive_WP.Where(x => x.Column_22.Contains(searchTerm)).ToList();
            }
            var modifiedData = ColumnList.Select(x => new
            {
                id = x.Id,
                text = x.Column_22
            }).ToList().GroupBy(p => p.text).Select(x => x.First()).OrderBy(x => x.text).ToList();
            return Json(modifiedData, JsonRequestBehavior.AllowGet);
        }
        public JsonResult GetColumn23List(string searchTerm)
        {
            var ColumnList = db.Aktive_WP.ToList();

            if (searchTerm != null)
            {
                ColumnList = db.Aktive_WP.Where(x => x.Column_23.Contains(searchTerm)).ToList();
            }
            var modifiedData = ColumnList.Select(x => new
            {
                id = x.Id,
                text = x.Column_23
            }).ToList().GroupBy(p => p.text).Select(x => x.First()).OrderBy(x => x.text).ToList();
            return Json(modifiedData, JsonRequestBehavior.AllowGet);
        }
        public JsonResult GetColumn24List(string searchTerm)
        {
            var ColumnList = db.Aktive_WP.ToList();

            if (searchTerm != null)
            {
                ColumnList = db.Aktive_WP.Where(x => x.Column_24.Contains(searchTerm)).ToList();
            }
            var modifiedData = ColumnList.Select(x => new
            {
                id = x.Id,
                text = x.Column_24
            }).ToList().GroupBy(p => p.text).Select(x => x.First()).OrderBy(x => x.text).ToList();
            return Json(modifiedData, JsonRequestBehavior.AllowGet);
        }
        public JsonResult GetColumn25List(string searchTerm)
        {
            var ColumnList = db.Aktive_WP.ToList();

            if (searchTerm != null)
            {
                ColumnList = db.Aktive_WP.Where(x => x.Column_25.Contains(searchTerm)).ToList();
            }
            var modifiedData = ColumnList.Select(x => new
            {
                id = x.Id,
                text = x.Column_25
            }).ToList().GroupBy(p => p.text).Select(x => x.First()).OrderBy(x => x.text).ToList();
            return Json(modifiedData, JsonRequestBehavior.AllowGet);
        }
        public JsonResult GetColumn26List(string searchTerm)
        {
            var ColumnList = db.Aktive_WP.ToList();

            if (searchTerm != null)
            {
                ColumnList = db.Aktive_WP.Where(x => x.Column_26.Contains(searchTerm)).ToList();
            }
            var modifiedData = ColumnList.Select(x => new
            {
                id = x.Id,
                text = x.Column_26
            }).ToList().GroupBy(p => p.text).Select(x => x.First()).OrderBy(x => x.text).ToList();
            return Json(modifiedData, JsonRequestBehavior.AllowGet);
        }
        public JsonResult GetColumn27List(string searchTerm)
        {
            var ColumnList = db.Aktive_WP.ToList();

            if (searchTerm != null)
            {
                ColumnList = db.Aktive_WP.Where(x => x.Column_27.Contains(searchTerm)).ToList();
            }
            var modifiedData = ColumnList.Select(x => new
            {
                id = x.Id,
                text = x.Column_27
            }).ToList().GroupBy(p => p.text).Select(x => x.First()).OrderBy(x => x.text).ToList();
            return Json(modifiedData, JsonRequestBehavior.AllowGet);
        }
        public JsonResult GetColumn28List(string searchTerm)
        {
            var ColumnList = db.Aktive_WP.ToList();

            if (searchTerm != null)
            {
                ColumnList = db.Aktive_WP.Where(x => x.Column_28.Contains(searchTerm)).ToList();
            }
            var modifiedData = ColumnList.Select(x => new
            {
                id = x.Id,
                text = x.Column_28
            }).ToList().GroupBy(p => p.text).Select(x => x.First()).OrderBy(x => x.text).ToList();
            return Json(modifiedData, JsonRequestBehavior.AllowGet);
        }
        public JsonResult GetColumn29List(string searchTerm)
        {
            var ColumnList = db.Aktive_WP.ToList();

            if (searchTerm != null)
            {
                ColumnList = db.Aktive_WP.Where(x => x.Column_29.Contains(searchTerm)).ToList();
            }
            var modifiedData = ColumnList.Select(x => new
            {
                id = x.Id,
                text = x.Column_29
            }).ToList().GroupBy(p => p.text).Select(x => x.First()).OrderBy(x => x.text).ToList();
            return Json(modifiedData, JsonRequestBehavior.AllowGet);
        }
        public JsonResult GetColumn30List(string searchTerm)
        {
            var ColumnList = db.Aktive_WP.ToList();

            if (searchTerm != null)
            {
                ColumnList = db.Aktive_WP.Where(x => x.Column_30.Contains(searchTerm)).ToList();
            }
            var modifiedData = ColumnList.Select(x => new
            {
                id = x.Id,
                text = x.Column_30
            }).ToList().GroupBy(p => p.text).Select(x => x.First()).OrderBy(x => x.text).ToList();
            return Json(modifiedData, JsonRequestBehavior.AllowGet);
        }

        public JsonResult GetWP(int? WPs)
        {
            var ColumnList = db.Aktive_WP.ToList();

            if (WPs != null)
            {
                ColumnList = null;
            }
            var modifiedData = ColumnList.Select(x => new
            {
                id = x.Id,
                text = x.Column_1
            }).ToList().GroupBy(p => p.text).Select(x => x.First()).OrderBy(x => x.text).ToList();
            return Json(modifiedData, JsonRequestBehavior.AllowGet);
        }

        [HttpGet]
        public ActionResult GetOptionColumn1()
        {
            string userString = User.Identity.Name.ToString().Split('\\').Last();    
            List<Settings_Filter> myFiltres = db.Settings_Filter.Where(w => w.User == userString).ToList();
            string ColumnValueNum = "";
            int IdValue = 0;
            string NameValue = "";
            List<Option> results = new List<Option>();
            if (myFiltres[0].Filter_Column_1 != null) ColumnValueNum = myFiltres[0].Filter_Column_1.ToString();
            if (ColumnValueNum != "")
            {
                string[] tokes = ColumnValueNum.Split(',');
                int[] IdArray = Array.ConvertAll<string, int>(tokes, int.Parse);
                for (int h = 0; h < IdArray.Length; h++)
                {
                    IdValue = IdArray[h];
                    NameValue = db.Aktive_WP.Where(x => x.Id == IdValue).FirstOrDefault().Column_1.ToString();
                    Option result = new Option();
                    result.text = NameValue;
                    result.id = IdValue;
                    results.Add(result);
                }
            }
            return Json(results, JsonRequestBehavior.AllowGet);
        }

        [HttpGet]
        public ActionResult GetOptionColumn2()
        {
            string userString = User.Identity.Name.ToString().Split('\\').Last();    
            List<Settings_Filter> myFiltres = db.Settings_Filter.Where(w => w.User == userString).ToList();
            string ColumnValueNum = "";
            int IdValue = 0;
            string NameValue = "";
            List<Option> results = new List<Option>();
            if (myFiltres[0].Filter_Column_2 != null) ColumnValueNum = myFiltres[0].Filter_Column_2.ToString();
            if (ColumnValueNum != "")
            {
                string[] tokes = ColumnValueNum.Split(',');
                int[] IdArray = Array.ConvertAll<string, int>(tokes, int.Parse);
                for (int h = 0; h < IdArray.Length; h++)
                {
                    IdValue = IdArray[h];
                    if (db.Aktive_WP.Where(x => x.Id == IdValue).FirstOrDefault().Column_2 != null)
                        NameValue = db.Aktive_WP.Where(x => x.Id == IdValue).FirstOrDefault().Column_2.ToString();
                    Option result = new Option();
                    result.text = NameValue;
                    result.id = IdValue;
                    results.Add(result);
                }
            }
            return Json(results, JsonRequestBehavior.AllowGet);
        }

        [HttpGet]
        public ActionResult GetOptionColumn3()
        {
            string userString = User.Identity.Name.ToString().Split('\\').Last();    
            List<Settings_Filter> myFiltres = db.Settings_Filter.Where(w => w.User == userString).ToList();
            string ColumnValueNum = "";
            int IdValue = 0;
            string NameValue = "";
            List<Option> results = new List<Option>();
            if (myFiltres[0].Filter_Column_3 != null) ColumnValueNum = myFiltres[0].Filter_Column_3.ToString();
            if (ColumnValueNum != "")
            {
                string[] tokes = ColumnValueNum.Split(',');
                int[] IdArray = Array.ConvertAll<string, int>(tokes, int.Parse);
                for (int h = 0; h < IdArray.Length; h++)
                {
                    IdValue = IdArray[h];
                    if (db.Aktive_WP.Where(x => x.Id == IdValue).FirstOrDefault().Column_3 != null)
                        NameValue = db.Aktive_WP.Where(x => x.Id == IdValue).FirstOrDefault().Column_3.ToString();
                    Option result = new Option();
                    result.text = NameValue;
                    result.id = IdValue;
                    results.Add(result);
                }
            }
            return Json(results, JsonRequestBehavior.AllowGet);
        }

        [HttpGet]
        public ActionResult GetOptionColumn4()
        {
            string userString = User.Identity.Name.ToString().Split('\\').Last();    
            List<Settings_Filter> myFiltres = db.Settings_Filter.Where(w => w.User == userString).ToList();
            string ColumnValueNum = "";
            int IdValue = 0;
            string NameValue = "";
            List<Option> results = new List<Option>();
            if (myFiltres[0].Filter_Column_4 != null) ColumnValueNum = myFiltres[0].Filter_Column_4.ToString();
            if (ColumnValueNum != "")
            {
                string[] tokes = ColumnValueNum.Split(',');
                int[] IdArray = Array.ConvertAll<string, int>(tokes, int.Parse);
                for (int h = 0; h < IdArray.Length; h++)
                {
                    IdValue = IdArray[h];
                    if (db.Aktive_WP.Where(x => x.Id == IdValue).FirstOrDefault().Column_4 != null)
                        NameValue = db.Aktive_WP.Where(x => x.Id == IdValue).FirstOrDefault().Column_4.ToString();
                    Option result = new Option();
                    result.text = NameValue;
                    result.id = IdValue;
                    results.Add(result);
                }
            }
            return Json(results, JsonRequestBehavior.AllowGet);
        }

        [HttpGet]
        public ActionResult GetOptionColumn5()
        {
            string userString = User.Identity.Name.ToString().Split('\\').Last();    
            List<Settings_Filter> myFiltres = db.Settings_Filter.Where(w => w.User == userString).ToList();
            string ColumnValueNum = "";
            int IdValue = 0;
            string NameValue = "";
            List<Option> results = new List<Option>();
            if (myFiltres[0].Filter_Column_5 != null) ColumnValueNum = myFiltres[0].Filter_Column_5.ToString();
            if (ColumnValueNum != "")
            {
                string[] tokes = ColumnValueNum.Split(',');
                int[] IdArray = Array.ConvertAll<string, int>(tokes, int.Parse);
                for (int h = 0; h < IdArray.Length; h++)
                {
                    IdValue = IdArray[h];
                    if (db.Aktive_WP.Where(x => x.Id == IdValue).FirstOrDefault().Column_5 != null)
                        NameValue = db.Aktive_WP.Where(x => x.Id == IdValue).FirstOrDefault().Column_5.ToString();
                    Option result = new Option();
                    result.text = NameValue;
                    result.id = IdValue;
                    results.Add(result);
                }
            }
            return Json(results, JsonRequestBehavior.AllowGet);
        }

        [HttpGet]
        public ActionResult GetOptionColumn6()
        {
            string userString = User.Identity.Name.ToString().Split('\\').Last();    
            List<Settings_Filter> myFiltres = db.Settings_Filter.Where(w => w.User == userString).ToList();
            string ColumnValueNum = "";
            int IdValue = 0;
            string NameValue = "";
            List<Option> results = new List<Option>();
            if (myFiltres[0].Filter_Column_6 != null) ColumnValueNum = myFiltres[0].Filter_Column_6.ToString();
            if (ColumnValueNum != "")
            {
                string[] tokes = ColumnValueNum.Split(',');
                int[] IdArray = Array.ConvertAll<string, int>(tokes, int.Parse);
                for (int h = 0; h < IdArray.Length; h++)
                {
                    IdValue = IdArray[h];
                    if (db.Aktive_WP.Where(x => x.Id == IdValue).FirstOrDefault().Column_6 != null)
                        NameValue = db.Aktive_WP.Where(x => x.Id == IdValue).FirstOrDefault().Column_6.ToString();
                    Option result = new Option();
                    result.text = NameValue;
                    result.id = IdValue;
                    results.Add(result);
                }
            }
            return Json(results, JsonRequestBehavior.AllowGet);
        }

        [HttpGet]
        public ActionResult GetOptionColumn7()
        {
            string userString = User.Identity.Name.ToString().Split('\\').Last();    
            List<Settings_Filter> myFiltres = db.Settings_Filter.Where(w => w.User == userString).ToList();
            string ColumnValueNum = "";
            int IdValue = 0;
            string NameValue = "";
            List<Option> results = new List<Option>();
            if (myFiltres[0].Filter_Column_7 != null) ColumnValueNum = myFiltres[0].Filter_Column_7.ToString();
            if (ColumnValueNum != "")
            {
                string[] tokes = ColumnValueNum.Split(',');
                int[] IdArray = Array.ConvertAll<string, int>(tokes, int.Parse);
                for (int h = 0; h < IdArray.Length; h++)
                {
                    IdValue = IdArray[h];
                    if (db.Aktive_WP.Where(x => x.Id == IdValue).FirstOrDefault().Column_7 != null)
                        NameValue = db.Aktive_WP.Where(x => x.Id == IdValue).FirstOrDefault().Column_7.ToString();
                    Option result = new Option();
                    result.text = NameValue;
                    result.id = IdValue;
                    results.Add(result);
                }
            }
            return Json(results, JsonRequestBehavior.AllowGet);
        }

        [HttpGet]
        public ActionResult GetOptionColumn8()
        {
            string userString = User.Identity.Name.ToString().Split('\\').Last();    
            List<Settings_Filter> myFiltres = db.Settings_Filter.Where(w => w.User == userString).ToList();
            string ColumnValueNum = "";
            int IdValue = 0;
            string NameValue = "";
            List<Option> results = new List<Option>();
            if (myFiltres[0].Filter_Column_8 != null) ColumnValueNum = myFiltres[0].Filter_Column_8.ToString();
            if (ColumnValueNum != "")
            {
                string[] tokes = ColumnValueNum.Split(',');
                int[] IdArray = Array.ConvertAll<string, int>(tokes, int.Parse);
                for (int h = 0; h < IdArray.Length; h++)
                {
                    IdValue = IdArray[h];
                    if (db.Aktive_WP.Where(x => x.Id == IdValue).FirstOrDefault().Column_8 != null)
                        NameValue = db.Aktive_WP.Where(x => x.Id == IdValue).FirstOrDefault().Column_8.ToString();
                    Option result = new Option();
                    result.text = NameValue;
                    result.id = IdValue;
                    results.Add(result);
                }
            }
            return Json(results, JsonRequestBehavior.AllowGet);
        }

        [HttpGet]
        public ActionResult GetOptionColumn9()
        {
            string userString = User.Identity.Name.ToString().Split('\\').Last();    
            List<Settings_Filter> myFiltres = db.Settings_Filter.Where(w => w.User == userString).ToList();
            string ColumnValueNum = "";
            int IdValue = 0;
            string NameValue = "";
            List<Option> results = new List<Option>();
            if (myFiltres[0].Filter_Column_9 != null) ColumnValueNum = myFiltres[0].Filter_Column_9.ToString();
            if (ColumnValueNum != "")
            {
                string[] tokes = ColumnValueNum.Split(',');
                int[] IdArray = Array.ConvertAll<string, int>(tokes, int.Parse);
                for (int h = 0; h < IdArray.Length; h++)
                {
                    IdValue = IdArray[h];
                    if (db.Aktive_WP.Where(x => x.Id == IdValue).FirstOrDefault().Column_9 != null)
                        NameValue = db.Aktive_WP.Where(x => x.Id == IdValue).FirstOrDefault().Column_9.ToString();
                    Option result = new Option();
                    result.text = NameValue;
                    result.id = IdValue;
                    results.Add(result);
                }
            }
            return Json(results, JsonRequestBehavior.AllowGet);
        }

        [HttpGet]
        public ActionResult GetOptionColumn10()
        {
            string userString = User.Identity.Name.ToString().Split('\\').Last();    
            List<Settings_Filter> myFiltres = db.Settings_Filter.Where(w => w.User == userString).ToList();
            string ColumnValueNum = "";
            int IdValue = 0;
            string NameValue = "";
            List<Option> results = new List<Option>();
            if (myFiltres[0].Filter_Column_10 != null) ColumnValueNum = myFiltres[0].Filter_Column_10.ToString();
            if (ColumnValueNum != "")
            {
                string[] tokes = ColumnValueNum.Split(',');
                int[] IdArray = Array.ConvertAll<string, int>(tokes, int.Parse);
                for (int h = 0; h < IdArray.Length; h++)
                {
                    IdValue = IdArray[h];
                    if (db.Aktive_WP.Where(x => x.Id == IdValue).FirstOrDefault().Column_10 != null)
                        NameValue = db.Aktive_WP.Where(x => x.Id == IdValue).FirstOrDefault().Column_10.ToString();
                    Option result = new Option();
                    result.text = NameValue;
                    result.id = IdValue;
                    results.Add(result);
                }
            }
            return Json(results, JsonRequestBehavior.AllowGet);
        }

        [HttpGet]
        public ActionResult GetOptionColumn11()
        {
            string userString = User.Identity.Name.ToString().Split('\\').Last();    
            List<Settings_Filter> myFiltres = db.Settings_Filter.Where(w => w.User == userString).ToList();
            string ColumnValueNum = "";
            int IdValue = 0;
            string NameValue = "";
            List<Option> results = new List<Option>();
            if (myFiltres[0].Filter_Column_11 != null) ColumnValueNum = myFiltres[0].Filter_Column_11.ToString();
            if (ColumnValueNum != "")
            {
                string[] tokes = ColumnValueNum.Split(',');
                int[] IdArray = Array.ConvertAll<string, int>(tokes, int.Parse);
                for (int h = 0; h < IdArray.Length; h++)
                {
                    IdValue = IdArray[h];
                    if (db.Aktive_WP.Where(x => x.Id == IdValue).FirstOrDefault().Column_11 != null)
                        NameValue = db.Aktive_WP.Where(x => x.Id == IdValue).FirstOrDefault().Column_11.ToString();
                    Option result = new Option();
                    result.text = NameValue;
                    result.id = IdValue;
                    results.Add(result);
                }
            }
            return Json(results, JsonRequestBehavior.AllowGet);
        }

        [HttpGet]
        public ActionResult GetOptionColumn12()
        {
            string userString = User.Identity.Name.ToString().Split('\\').Last();    
            List<Settings_Filter> myFiltres = db.Settings_Filter.Where(w => w.User == userString).ToList();
            string ColumnValueNum = "";
            int IdValue = 0;
            string NameValue = "";
            List<Option> results = new List<Option>();
            if (myFiltres[0].Filter_Column_12 != null) ColumnValueNum = myFiltres[0].Filter_Column_12.ToString();
            if (ColumnValueNum != "")
            {
                string[] tokes = ColumnValueNum.Split(',');
                int[] IdArray = Array.ConvertAll<string, int>(tokes, int.Parse);
                for (int h = 0; h < IdArray.Length; h++)
                {
                    IdValue = IdArray[h];
                    if (db.Aktive_WP.Where(x => x.Id == IdValue).FirstOrDefault().Column_12 != null)
                        NameValue = db.Aktive_WP.Where(x => x.Id == IdValue).FirstOrDefault().Column_12.ToString();
                    Option result = new Option();
                    result.text = NameValue;
                    result.id = IdValue;
                    results.Add(result);
                }
            }
            return Json(results, JsonRequestBehavior.AllowGet);
        }

        [HttpGet]
        public ActionResult GetOptionColumn13()
        {
            string userString = User.Identity.Name.ToString().Split('\\').Last();    
            List<Settings_Filter> myFiltres = db.Settings_Filter.Where(w => w.User == userString).ToList();
            string ColumnValueNum = "";
            int IdValue = 0;
            string NameValue = "";
            List<Option> results = new List<Option>();
            if (myFiltres[0].Filter_Column_13 != null) ColumnValueNum = myFiltres[0].Filter_Column_13.ToString();
            if (ColumnValueNum != "")
            {
                string[] tokes = ColumnValueNum.Split(',');
                int[] IdArray = Array.ConvertAll<string, int>(tokes, int.Parse);
                for (int h = 0; h < IdArray.Length; h++)
                {
                    IdValue = IdArray[h];
                    if (db.Aktive_WP.Where(x => x.Id == IdValue).FirstOrDefault().Column_13 != null)
                        NameValue = db.Aktive_WP.Where(x => x.Id == IdValue).FirstOrDefault().Column_13.ToString();
                    Option result = new Option();
                    result.text = NameValue;
                    result.id = IdValue;
                    results.Add(result);
                }
            }
            return Json(results, JsonRequestBehavior.AllowGet);
        }

        [HttpGet]
        public ActionResult GetOptionColumn14()
        {
            string userString = User.Identity.Name.ToString().Split('\\').Last();    
            List<Settings_Filter> myFiltres = db.Settings_Filter.Where(w => w.User == userString).ToList();
            string ColumnValueNum = "";
            int IdValue = 0;
            string NameValue = "";
            List<Option> results = new List<Option>();
            if (myFiltres[0].Filter_Column_14 != null) ColumnValueNum = myFiltres[0].Filter_Column_14.ToString();
            if (ColumnValueNum != "")
            {
                string[] tokes = ColumnValueNum.Split(',');
                int[] IdArray = Array.ConvertAll<string, int>(tokes, int.Parse);
                for (int h = 0; h < IdArray.Length; h++)
                {
                    IdValue = IdArray[h];
                    if (db.Aktive_WP.Where(x => x.Id == IdValue).FirstOrDefault().Column_14 != null)
                        NameValue = db.Aktive_WP.Where(x => x.Id == IdValue).FirstOrDefault().Column_14.ToString();
                    Option result = new Option();
                    result.text = NameValue;
                    result.id = IdValue;
                    results.Add(result);
                }
            }
            return Json(results, JsonRequestBehavior.AllowGet);
        }

        [HttpGet]
        public ActionResult GetOptionColumn15()
        {
            string userString = User.Identity.Name.ToString().Split('\\').Last();    
            List<Settings_Filter> myFiltres = db.Settings_Filter.Where(w => w.User == userString).ToList();
            string ColumnValueNum = "";
            int IdValue = 0;
            string NameValue = "";
            List<Option> results = new List<Option>();
            if (myFiltres[0].Filter_Column_15 != null) ColumnValueNum = myFiltres[0].Filter_Column_15.ToString();
            if (ColumnValueNum != "")
            {
                string[] tokes = ColumnValueNum.Split(',');
                int[] IdArray = Array.ConvertAll<string, int>(tokes, int.Parse);
                for (int h = 0; h < IdArray.Length; h++)
                {
                    IdValue = IdArray[h];
                    if (db.Aktive_WP.Where(x => x.Id == IdValue).FirstOrDefault().Column_15 != null)
                        NameValue = db.Aktive_WP.Where(x => x.Id == IdValue).FirstOrDefault().Column_15.ToString();
                    Option result = new Option();
                    result.text = NameValue;
                    result.id = IdValue;
                    results.Add(result);
                }
            }
            return Json(results, JsonRequestBehavior.AllowGet);
        }

        [HttpGet]
        public ActionResult GetOptionColumn16()
        {
            string userString = User.Identity.Name.ToString().Split('\\').Last();    
            List<Settings_Filter> myFiltres = db.Settings_Filter.Where(w => w.User == userString).ToList();
            string ColumnValueNum = "";
            int IdValue = 0;
            string NameValue = "";
            List<Option> results = new List<Option>();
            if (myFiltres[0].Filter_Column_16 != null) ColumnValueNum = myFiltres[0].Filter_Column_16.ToString();
            if (ColumnValueNum != "")
            {
                string[] tokes = ColumnValueNum.Split(',');
                int[] IdArray = Array.ConvertAll<string, int>(tokes, int.Parse);
                for (int h = 0; h < IdArray.Length; h++)
                {
                    IdValue = IdArray[h];
                    if (db.Aktive_WP.Where(x => x.Id == IdValue).FirstOrDefault().Column_16 != null)
                        NameValue = db.Aktive_WP.Where(x => x.Id == IdValue).FirstOrDefault().Column_16.ToString();
                    Option result = new Option();
                    result.text = NameValue;
                    result.id = IdValue;
                    results.Add(result);
                }
            }
            return Json(results, JsonRequestBehavior.AllowGet);
        }

        [HttpGet]
        public ActionResult GetOptionColumn17()
        {
            string userString = User.Identity.Name.ToString().Split('\\').Last();    
            List<Settings_Filter> myFiltres = db.Settings_Filter.Where(w => w.User == userString).ToList();
            string ColumnValueNum = "";
            int IdValue = 0;
            string NameValue = "";
            List<Option> results = new List<Option>();
            if (myFiltres[0].Filter_Column_17 != null) ColumnValueNum = myFiltres[0].Filter_Column_17.ToString();
            if (ColumnValueNum != "")
            {
                string[] tokes = ColumnValueNum.Split(',');
                int[] IdArray = Array.ConvertAll<string, int>(tokes, int.Parse);
                for (int h = 0; h < IdArray.Length; h++)
                {
                    IdValue = IdArray[h];
                    if (db.Aktive_WP.Where(x => x.Id == IdValue).FirstOrDefault().Column_17 != null)
                        NameValue = db.Aktive_WP.Where(x => x.Id == IdValue).FirstOrDefault().Column_17.ToString();
                    Option result = new Option();
                    result.text = NameValue;
                    result.id = IdValue;
                    results.Add(result);
                }
            }
            return Json(results, JsonRequestBehavior.AllowGet);
        }

        [HttpGet]
        public ActionResult GetOptionColumn18()
        {
            string userString = User.Identity.Name.ToString().Split('\\').Last();    
            List<Settings_Filter> myFiltres = db.Settings_Filter.Where(w => w.User == userString).ToList();
            string ColumnValueNum = "";
            int IdValue = 0;
            string NameValue = "";
            List<Option> results = new List<Option>();
            if (myFiltres[0].Filter_Column_18 != null) ColumnValueNum = myFiltres[0].Filter_Column_18.ToString();
            if (ColumnValueNum != "")
            {
                string[] tokes = ColumnValueNum.Split(',');
                int[] IdArray = Array.ConvertAll<string, int>(tokes, int.Parse);
                for (int h = 0; h < IdArray.Length; h++)
                {
                    IdValue = IdArray[h];
                    if (db.Aktive_WP.Where(x => x.Id == IdValue).FirstOrDefault().Column_18 != null)
                        NameValue = db.Aktive_WP.Where(x => x.Id == IdValue).FirstOrDefault().Column_18.ToString();
                    Option result = new Option();
                    result.text = NameValue;
                    result.id = IdValue;
                    results.Add(result);
                }
            }
            return Json(results, JsonRequestBehavior.AllowGet);
        }

        [HttpGet]
        public ActionResult GetOptionColumn19()
        {
            string userString = User.Identity.Name.ToString().Split('\\').Last();    
            List<Settings_Filter> myFiltres = db.Settings_Filter.Where(w => w.User == userString).ToList();
            string ColumnValueNum = "";
            int IdValue = 0;
            string NameValue = "";
            List<Option> results = new List<Option>();
            if (myFiltres[0].Filter_Column_19 != null) ColumnValueNum = myFiltres[0].Filter_Column_19.ToString();
            if (ColumnValueNum != "")
            {
                string[] tokes = ColumnValueNum.Split(',');
                int[] IdArray = Array.ConvertAll<string, int>(tokes, int.Parse);
                for (int h = 0; h < IdArray.Length; h++)
                {
                    IdValue = IdArray[h];
                    if (db.Aktive_WP.Where(x => x.Id == IdValue).FirstOrDefault().Column_19 != null)
                        NameValue = db.Aktive_WP.Where(x => x.Id == IdValue).FirstOrDefault().Column_19.ToString();
                    Option result = new Option();
                    result.text = NameValue;
                    result.id = IdValue;
                    results.Add(result);
                }
            }
            return Json(results, JsonRequestBehavior.AllowGet);
        }

        [HttpGet]
        public ActionResult GetOptionColumn20()
        {
            string userString = User.Identity.Name.ToString().Split('\\').Last();    
            List<Settings_Filter> myFiltres = db.Settings_Filter.Where(w => w.User == userString).ToList();
            string ColumnValueNum = "";
            int IdValue = 0;
            string NameValue = "";
            List<Option> results = new List<Option>();
            if (myFiltres[0].Filter_Column_20 != null) ColumnValueNum = myFiltres[0].Filter_Column_20.ToString();
            if (ColumnValueNum != "")
            {
                string[] tokes = ColumnValueNum.Split(',');
                int[] IdArray = Array.ConvertAll<string, int>(tokes, int.Parse);
                for (int h = 0; h < IdArray.Length; h++)
                {
                    IdValue = IdArray[h];
                    if (db.Aktive_WP.Where(x => x.Id == IdValue).FirstOrDefault().Column_20 != null)
                        NameValue = db.Aktive_WP.Where(x => x.Id == IdValue).FirstOrDefault().Column_20.ToString();
                    Option result = new Option();
                    result.text = NameValue;
                    result.id = IdValue;
                    results.Add(result);
                }
            }
            return Json(results, JsonRequestBehavior.AllowGet);
        }

        [HttpGet]
        public ActionResult GetOptionColumn21()
        {
            string userString = User.Identity.Name.ToString().Split('\\').Last();    
            List<Settings_Filter> myFiltres = db.Settings_Filter.Where(w => w.User == userString).ToList();
            string ColumnValueNum = "";
            int IdValue = 0;
            string NameValue = "";
            List<Option> results = new List<Option>();
            if (myFiltres[0].Filter_Column_21 != null) ColumnValueNum = myFiltres[0].Filter_Column_21.ToString();
            if (ColumnValueNum != "")
            {
                string[] tokes = ColumnValueNum.Split(',');
                int[] IdArray = Array.ConvertAll<string, int>(tokes, int.Parse);
                for (int h = 0; h < IdArray.Length; h++)
                {
                    IdValue = IdArray[h];
                    if (db.Aktive_WP.Where(x => x.Id == IdValue).FirstOrDefault().Column_21 != null)
                        NameValue = db.Aktive_WP.Where(x => x.Id == IdValue).FirstOrDefault().Column_21.ToString();
                    Option result = new Option();
                    result.text = NameValue;
                    result.id = IdValue;
                    results.Add(result);
                }
            }
            return Json(results, JsonRequestBehavior.AllowGet);
        }

        [HttpGet]
        public ActionResult GetOptionColumn22()
        {
            string userString = User.Identity.Name.ToString().Split('\\').Last();    
            List<Settings_Filter> myFiltres = db.Settings_Filter.Where(w => w.User == userString).ToList();
            string ColumnValueNum = "";
            int IdValue = 0;
            string NameValue = "";
            List<Option> results = new List<Option>();
            if (myFiltres[0].Filter_Column_22 != null) ColumnValueNum = myFiltres[0].Filter_Column_22.ToString();
            if (ColumnValueNum != "")
            {
                string[] tokes = ColumnValueNum.Split(',');
                int[] IdArray = Array.ConvertAll<string, int>(tokes, int.Parse);
                for (int h = 0; h < IdArray.Length; h++)
                {
                    IdValue = IdArray[h];
                    if (db.Aktive_WP.Where(x => x.Id == IdValue).FirstOrDefault().Column_22 != null)
                        NameValue = db.Aktive_WP.Where(x => x.Id == IdValue).FirstOrDefault().Column_22.ToString();
                    Option result = new Option();
                    result.text = NameValue;
                    result.id = IdValue;
                    results.Add(result);
                }
            }
            return Json(results, JsonRequestBehavior.AllowGet);
        }

        [HttpGet]
        public ActionResult GetOptionColumn23()
        {
            string userString = User.Identity.Name.ToString().Split('\\').Last();    
            List<Settings_Filter> myFiltres = db.Settings_Filter.Where(w => w.User == userString).ToList();
            string ColumnValueNum = "";
            int IdValue = 0;
            string NameValue = "";
            List<Option> results = new List<Option>();
            if (myFiltres[0].Filter_Column_23 != null) ColumnValueNum = myFiltres[0].Filter_Column_23.ToString();
            if (ColumnValueNum != "")
            {
                string[] tokes = ColumnValueNum.Split(',');
                int[] IdArray = Array.ConvertAll<string, int>(tokes, int.Parse);
                for (int h = 0; h < IdArray.Length; h++)
                {
                    IdValue = IdArray[h];
                    if (db.Aktive_WP.Where(x => x.Id == IdValue).FirstOrDefault().Column_23 != null)
                        NameValue = db.Aktive_WP.Where(x => x.Id == IdValue).FirstOrDefault().Column_23.ToString();
                    Option result = new Option();
                    result.text = NameValue;
                    result.id = IdValue;
                    results.Add(result);
                }
            }
            return Json(results, JsonRequestBehavior.AllowGet);
        }

        [HttpGet]
        public ActionResult GetOptionColumn24()
        {
            string userString = User.Identity.Name.ToString().Split('\\').Last();    
            List<Settings_Filter> myFiltres = db.Settings_Filter.Where(w => w.User == userString).ToList();
            string ColumnValueNum = "";
            int IdValue = 0;
            string NameValue = "";
            List<Option> results = new List<Option>();
            if (myFiltres[0].Filter_Column_24 != null) ColumnValueNum = myFiltres[0].Filter_Column_24.ToString();
            if (ColumnValueNum != "")
            {
                string[] tokes = ColumnValueNum.Split(',');
                int[] IdArray = Array.ConvertAll<string, int>(tokes, int.Parse);
                for (int h = 0; h < IdArray.Length; h++)
                {
                    IdValue = IdArray[h];
                    if (db.Aktive_WP.Where(x => x.Id == IdValue).FirstOrDefault().Column_24 != null)
                        NameValue = db.Aktive_WP.Where(x => x.Id == IdValue).FirstOrDefault().Column_24.ToString();
                    Option result = new Option();
                    result.text = NameValue;
                    result.id = IdValue;
                    results.Add(result);
                }
            }
            return Json(results, JsonRequestBehavior.AllowGet);
        }

        [HttpGet]
        public ActionResult GetOptionColumn25()
        {
            string userString = User.Identity.Name.ToString().Split('\\').Last();    
            List<Settings_Filter> myFiltres = db.Settings_Filter.Where(w => w.User == userString).ToList();
            string ColumnValueNum = "";
            int IdValue = 0;
            string NameValue = "";
            List<Option> results = new List<Option>();
            if (myFiltres[0].Filter_Column_25 != null) ColumnValueNum = myFiltres[0].Filter_Column_25.ToString();
            if (ColumnValueNum != "")
            {
                string[] tokes = ColumnValueNum.Split(',');
                int[] IdArray = Array.ConvertAll<string, int>(tokes, int.Parse);
                for (int h = 0; h < IdArray.Length; h++)
                {
                    IdValue = IdArray[h];
                    if (db.Aktive_WP.Where(x => x.Id == IdValue).FirstOrDefault().Column_25 != null)
                        NameValue = db.Aktive_WP.Where(x => x.Id == IdValue).FirstOrDefault().Column_25.ToString();
                    Option result = new Option();
                    result.text = NameValue;
                    result.id = IdValue;
                    results.Add(result);
                }
            }
            return Json(results, JsonRequestBehavior.AllowGet);
        }

        [HttpGet]
        public ActionResult GetOptionColumn26()
        {
            string userString = User.Identity.Name.ToString().Split('\\').Last();    
            List<Settings_Filter> myFiltres = db.Settings_Filter.Where(w => w.User == userString).ToList();
            string ColumnValueNum = "";
            int IdValue = 0;
            string NameValue = "";
            List<Option> results = new List<Option>();
            if (myFiltres[0].Filter_Column_26 != null) ColumnValueNum = myFiltres[0].Filter_Column_26.ToString();
            if (ColumnValueNum != "")
            {
                string[] tokes = ColumnValueNum.Split(',');
                int[] IdArray = Array.ConvertAll<string, int>(tokes, int.Parse);
                for (int h = 0; h < IdArray.Length; h++)
                {
                    IdValue = IdArray[h];
                    if (db.Aktive_WP.Where(x => x.Id == IdValue).FirstOrDefault().Column_26 != null)
                        NameValue = db.Aktive_WP.Where(x => x.Id == IdValue).FirstOrDefault().Column_26.ToString();
                    Option result = new Option();
                    result.text = NameValue;
                    result.id = IdValue;
                    results.Add(result);
                }
            }
            return Json(results, JsonRequestBehavior.AllowGet);
        }

        [HttpGet]
        public ActionResult GetOptionColumn27()
        {
            string userString = User.Identity.Name.ToString().Split('\\').Last();    
            List<Settings_Filter> myFiltres = db.Settings_Filter.Where(w => w.User == userString).ToList();
            string ColumnValueNum = "";
            int IdValue = 0;
            string NameValue = "";
            List<Option> results = new List<Option>();
            if (myFiltres[0].Filter_Column_27 != null) ColumnValueNum = myFiltres[0].Filter_Column_27.ToString();
            if (ColumnValueNum != "")
            {
                string[] tokes = ColumnValueNum.Split(',');
                int[] IdArray = Array.ConvertAll<string, int>(tokes, int.Parse);
                for (int h = 0; h < IdArray.Length; h++)
                {
                    IdValue = IdArray[h];
                    if (db.Aktive_WP.Where(x => x.Id == IdValue).FirstOrDefault().Column_27 != null)
                        NameValue = db.Aktive_WP.Where(x => x.Id == IdValue).FirstOrDefault().Column_27.ToString();
                    Option result = new Option();
                    result.text = NameValue;
                    result.id = IdValue;
                    results.Add(result);
                }
            }
            return Json(results, JsonRequestBehavior.AllowGet);
        }

        [HttpGet]
        public ActionResult GetOptionColumn28()
        {
            string userString = User.Identity.Name.ToString().Split('\\').Last();    
            List<Settings_Filter> myFiltres = db.Settings_Filter.Where(w => w.User == userString).ToList();
            string ColumnValueNum = "";
            int IdValue = 0;
            string NameValue = "";
            List<Option> results = new List<Option>();
            if (myFiltres[0].Filter_Column_28 != null) ColumnValueNum = myFiltres[0].Filter_Column_28.ToString();
            if (ColumnValueNum != "")
            {
                string[] tokes = ColumnValueNum.Split(',');
                int[] IdArray = Array.ConvertAll<string, int>(tokes, int.Parse);
                for (int h = 0; h < IdArray.Length; h++)
                {
                    IdValue = IdArray[h];
                    if (db.Aktive_WP.Where(x => x.Id == IdValue).FirstOrDefault().Column_28 != null)
                        NameValue = db.Aktive_WP.Where(x => x.Id == IdValue).FirstOrDefault().Column_28.ToString();
                    Option result = new Option();
                    result.text = NameValue;
                    result.id = IdValue;
                    results.Add(result);
                }
            }
            return Json(results, JsonRequestBehavior.AllowGet);
        }

        [HttpGet]
        public ActionResult GetOptionColumn29()
        {
            string userString = User.Identity.Name.ToString().Split('\\').Last();    
            List<Settings_Filter> myFiltres = db.Settings_Filter.Where(w => w.User == userString).ToList();
            string ColumnValueNum = "";
            int IdValue = 0;
            string NameValue = "";
            List<Option> results = new List<Option>();
            if (myFiltres[0].Filter_Column_29 != null) ColumnValueNum = myFiltres[0].Filter_Column_29.ToString();
            if (ColumnValueNum != "")
            {
                string[] tokes = ColumnValueNum.Split(',');
                int[] IdArray = Array.ConvertAll<string, int>(tokes, int.Parse);
                for (int h = 0; h < IdArray.Length; h++)
                {
                    IdValue = IdArray[h];
                    if (db.Aktive_WP.Where(x => x.Id == IdValue).FirstOrDefault().Column_29 != null)
                        NameValue = db.Aktive_WP.Where(x => x.Id == IdValue).FirstOrDefault().Column_29.ToString();
                    Option result = new Option();
                    result.text = NameValue;
                    result.id = IdValue;
                    results.Add(result);
                }
            }
            return Json(results, JsonRequestBehavior.AllowGet);
        }

        [HttpGet]
        public ActionResult GetOptionColumn30()
        {
            string userString = User.Identity.Name.ToString().Split('\\').Last();    
            List<Settings_Filter> myFiltres = db.Settings_Filter.Where(w => w.User == userString).ToList();
            string ColumnValueNum = "";
            int IdValue = 0;
            string NameValue = "";
            List<Option> results = new List<Option>();
            if (myFiltres[0].Filter_Column_30 != null) ColumnValueNum = myFiltres[0].Filter_Column_30.ToString();
            if (ColumnValueNum != "")
            {
                string[] tokes = ColumnValueNum.Split(',');
                int[] IdArray = Array.ConvertAll<string, int>(tokes, int.Parse);
                for (int h = 0; h < IdArray.Length; h++)
                {
                    IdValue = IdArray[h];
                    if (db.Aktive_WP.Where(x => x.Id == IdValue).FirstOrDefault().Column_30 != null)
                        NameValue = db.Aktive_WP.Where(x => x.Id == IdValue).FirstOrDefault().Column_30.ToString();
                    Option result = new Option();
                    result.text = NameValue;
                    result.id = IdValue;
                    results.Add(result);
                }
            }
            return Json(results, JsonRequestBehavior.AllowGet);
        }
    }
}
