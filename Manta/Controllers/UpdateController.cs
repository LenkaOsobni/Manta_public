using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using Manta_dev_Data;

namespace Manta.Controllers
{
     public class UpdateController : Controller
    {
        private Manta_dev_DataSet_DWH db_DWH = new Manta_dev_DataSet_DWH();
        private Manta_dev_DataSet db = new Manta_dev_DataSet();


        [HttpGet]
        public ActionResult UpdateWPs()
        {
            DateTime DateNow = DateTime.Now;
            string updateCalendarREL = "RE" + db.Settings_Name_Columns.First().Current_Release; ;
            List<Milestone_DWH> DeployMilestone = db_DWH.Milestones_DWH.Where(w => w.MILESTONE_NAME == "Instalace na produkční prostředí").OrderBy(x => x.RELEASE_NAME).ToList();
            for (int i = 0; i < DeployMilestone.Count; i++)
            {
                if (DeployMilestone[i].RELEASE_NAME == "RE04") // přeskoč divné záznamy
                {
                    for (int j = i; j < DeployMilestone.Count; j++) //projdi relevantní záznamy
                    {
                        if (DateNow <= DeployMilestone[j].MILESTONE_STARTTIME)// první datum, které ještě nenastalo
                        {
                            updateCalendarREL = DeployMilestone[j].RELEASE_NAME;
                            break;
                        }
                        else
                        {
                            updateCalendarREL = "RE" + db.Settings_Name_Columns.First().Current_Release;
                        }
                    }
                    break;//už bylo vše prohledáno
                } 
            }

            //najdi REL pro update (ručně zadávaná verze)
            //  Settings_Name_Columns Sett = new Settings_Name_Columns();
            // string numberCalendarREL = db.Settings_Name_Columns.First().Current_Release;
            // string updateCalendarREL = "RE" + numberCalendarREL;

            //stahni data z DWH o hledanem REL
            List<Milestone_DWH> updateCalendar = db_DWH.Milestones_DWH.Where(w => w.RELEASE_NAME == updateCalendarREL).ToList();
            //vymaz dosavadni data o REL
            List<Calendar> delCalendar = db.Calendar.Where(w => w.RELEASE == updateCalendarREL).ToList();
            for (int i = 0; i < delCalendar.Count; i++)
            {
                db.Calendar.Remove(delCalendar[i]);
                db.SaveChanges();
            }
            //vloz nova data
            foreach (Milestone_DWH mil in db_DWH.Milestones_DWH.Where(w => w.RELEASE_NAME == updateCalendarREL))
            {
                Calendar newCalendar = new Calendar();
                newCalendar.Id = mil.Guid;
                newCalendar.CAPTION = mil.MILESTONE_NAME;
                newCalendar.STARTTIME = mil.MILESTONE_STARTTIME;
                newCalendar.ENDTIME = mil.MILESTONE_ENDTIME;
                newCalendar.RELEASE = mil.RELEASE_NAME;

                db.Calendar.Add(newCalendar);
                db.SaveChanges();
            }
            db.Settings_Name_Columns.First().Current_Release = updateCalendarREL.Substring(2);

            //aktualizace WP
            //WP k aktualizaci
            List<Aktive_WP> list_Aktiv_WP = db.Aktive_WP.ToList();
            for (int i = 0; i < list_Aktiv_WP.Count; i++)
            {
                string awp = list_Aktiv_WP[i].Column_1;
                WP_DWH newInformation = db_DWH.WPs_DWH.Where(w => w.WP == awp).FirstOrDefault();
                if (newInformation == null) continue; //REQ....
                Aktive_WP update_WP = db.Aktive_WP.Where(w => w.Column_1 == awp).FirstOrDefault();
                update_WP.Column_2 = newInformation.WP_NAME;
                update_WP.Column_3 = newInformation.PROJECT;
                update_WP.Column_4 = newInformation.PROJECT_NAME;
                update_WP.Column_7 = newInformation.PROJECT_MANAGER_NAME;
                update_WP.Column_10 = newInformation.WP_STEP_NAME;


                db.Entry(update_WP).State = EntityState.Modified;
                db.SaveChanges();
            }
            return RedirectToAction("WP", "WP");
        }
    }
}