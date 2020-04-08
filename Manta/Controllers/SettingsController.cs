using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using Manta_dev_ViewModel;
using Manta_dev_Data;
using System.Data.Entity;

namespace Manta_dev.Controllers
{
    public class SettingsController : Controller
    {
        private Manta_dev_DataSet db = new Manta_dev_DataSet();

        [HttpGet]
        public ActionResult Settings()
        {
            Setting_Column_Name_ViewModel vm = new Setting_Column_Name_ViewModel { };
            vm.HandleRequest();
            return View(vm);
        }

        [HttpGet]
        public ActionResult Settings_Edit_Column(int id)
        {
            var editName = db.Settings_Name_Columns.Where(w => w.User == "All").First();
            if (editName == null)
            {
                return HttpNotFound();
            }
            return View(editName);
        }

        [HttpPost]
        public ActionResult Settings_Edit_Column(Settings_Name_Columns sett)
        {
            sett.User = "All";
            if (ModelState.IsValid)
            {
                db.Entry(sett).State = EntityState.Modified;

                db.SaveChanges();
                return RedirectToAction("Settings");
            }
            return View(sett);
        }

        [HttpGet]
        public ActionResult Settings_Edit_Autorization(int id)
        {
            ViewBag.Settings = db.Settings_Name_Columns.First();
            var autor = db.Setting_Group.Where(w => w.Id == id).First();
            if (autor == null)
            {
                return HttpNotFound();
            }
            return View(autor);
        }

        [HttpPost]
        public ActionResult Settings_Edit_Autorization(Setting_Group sett)
        {
            if (ModelState.IsValid)
            {
                db.Entry(sett).State = EntityState.Modified;
                db.SaveChanges();
                return RedirectToAction("Settings");
            }
            return View(sett);
        }
    }
}