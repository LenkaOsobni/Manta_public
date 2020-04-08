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
using PagedList;
using PagedList.Mvc;

namespace Manta_dev.Controllers
{
    public class ArchivController : Controller
    {
        private Manta_dev_DataSet db = new Manta_dev_DataSet();
        public readonly string Username = WindowsIdentity.GetCurrent().Name.ToString();
        public bool IsEditor = false;
        public bool IsAdmin = false;
        public ArchivController()
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

        }

        [HttpGet]
        public ActionResult Archiv(string searching, bool? all, bool? c1ch, bool? c2ch, bool? c3ch, bool? c4ch, bool? c5ch, bool? c6ch, bool? c7ch, bool? c8ch, bool? c9ch, bool? c10ch, bool? c11ch, bool? c12ch, bool? c13ch, bool? c14ch, bool? c15ch, bool? c16ch, bool? c17ch, bool? c18ch, bool? c19ch, bool? c20ch, bool? c21ch, bool? c22ch, bool? c23ch, bool? c24ch, bool? c25ch, bool? c26ch, bool? c27ch, bool? c28ch, bool? c29ch, bool? c30ch, int? page)
        {
            Archiv_ViewModel vm = new Archiv_ViewModel { };
            vm.HandleRequest(page);
            List<Archiv> array = db.Archiv.ToList();
            if ((searching == null ? 0 : searching.Length) > 0)
            {
                List<Archiv> pom = new List<Archiv>();
                for (int i = 0; i < array.Count() ; i++)
                {
                    int a = array[i].Id;
                    if((array[i].Column_1 != null && (all == true || c1ch == true) && array[i].Column_1.ToLower().Contains(searching.ToLower())) ||
                       (array[i].Column_2 != null && (all == true || c2ch == true) && array[i].Column_2.ToLower().Contains(searching.ToLower())) ||
                       (array[i].Column_3 != null && (all == true || c3ch == true) && array[i].Column_3.ToLower().Contains(searching.ToLower())) ||
                       (array[i].Column_4 != null && (all == true || c4ch == true) && array[i].Column_4.ToLower().Contains(searching.ToLower())) ||
                       (array[i].Column_5 != null && (all == true || c5ch == true) && array[i].Column_5.ToLower().Contains(searching.ToLower())) ||
                       (array[i].Column_6 != null && (all == true || c6ch == true) && array[i].Column_6.ToLower().Contains(searching.ToLower())) ||
                       (array[i].Column_7 != null && (all == true || c7ch == true) && array[i].Column_7.ToLower().Contains(searching.ToLower())) ||
                       (array[i].Column_8 != null && (all == true || c8ch == true) && array[i].Column_8.ToLower().Contains(searching.ToLower())) ||
                       (array[i].Column_9 != null && (all == true || c9ch == true) && array[i].Column_9.ToLower().Contains(searching.ToLower())) ||
                       (array[i].Column_10 != null && (all == true || c10ch == true) && array[i].Column_10.ToLower().Contains(searching.ToLower())) ||
                       (array[i].Column_11 != null && (all == true || c11ch == true) && array[i].Column_11.ToLower().Contains(searching.ToLower())) ||
                       (array[i].Column_12 != null && (all == true || c12ch == true) && array[i].Column_12.ToLower().Contains(searching.ToLower())) ||
                       (array[i].Column_13 != null && (all == true || c13ch == true) && array[i].Column_13.ToLower().Contains(searching.ToLower())) ||
                       (array[i].Column_14 != null && (all == true || c14ch == true) && array[i].Column_14.ToLower().Contains(searching.ToLower())) ||
                       (array[i].Column_15 != null && (all == true || c15ch == true) && array[i].Column_15.ToLower().Contains(searching.ToLower())) ||
                       (array[i].Column_16 != null && (all == true || c16ch == true) && array[i].Column_16.ToLower().Contains(searching.ToLower())) ||
                       (array[i].Column_17 != null && (all == true || c17ch == true) && array[i].Column_17.ToLower().Contains(searching.ToLower())) ||
                       (array[i].Column_18 != null && (all == true || c18ch == true) && array[i].Column_18.ToLower().Contains(searching.ToLower())) ||
                       (array[i].Column_19 != null && (all == true || c19ch == true) && array[i].Column_19.ToLower().Contains(searching.ToLower())) ||
                       (array[i].Column_20 != null && (all == true || c20ch == true) && array[i].Column_20.ToLower().Contains(searching.ToLower())) ||
                       (array[i].Column_21 != null && (all == true || c21ch == true) && array[i].Column_21.ToLower().Contains(searching.ToLower())) ||
                       (array[i].Column_22 != null && (all == true || c22ch == true) && array[i].Column_22.ToLower().Contains(searching.ToLower())) ||
                       (array[i].Column_23 != null && (all == true || c23ch == true) && array[i].Column_23.ToLower().Contains(searching.ToLower())) ||
                       (array[i].Column_24 != null && (all == true || c24ch == true) && array[i].Column_24.ToLower().Contains(searching.ToLower())) ||
                       (array[i].Column_25 != null && (all == true || c25ch == true) && array[i].Column_25.ToLower().Contains(searching.ToLower())) ||
                       (array[i].Column_26 != null && (all == true || c26ch == true) && array[i].Column_26.ToLower().Contains(searching.ToLower())) ||
                       (array[i].Column_27 != null && (all == true || c27ch == true) && array[i].Column_27.ToLower().Contains(searching.ToLower())) ||
                       (array[i].Column_28 != null && (all == true || c28ch == true) && array[i].Column_28.ToLower().Contains(searching.ToLower())) ||
                       (array[i].Column_29 != null && (all == true || c29ch == true) && array[i].Column_29.ToLower().Contains(searching.ToLower())) ||
                       (array[i].Column_30 != null && (all == true || c30ch == true) && array[i].Column_30.ToLower().Contains(searching.ToLower())) )
                                          
                    {
                        pom.AddRange(db.Archiv.Where(s => s.Id == a).DefaultIfEmpty().ToList());
                    }
                }
                List<Archiv> k = array.Intersect(pom).ToList();
                array = k;
            }

            vm.DataCollection_Archiv = array.Distinct().ToList().ToPagedList(page ?? 1, 10);
            return View(vm);
        }

        [HttpGet]
        public ActionResult Add(int Id)
        {
            var archivedWP = db.Aktive_WP.Find(Id);
            var archiv = new Archiv();

            archiv.Id_Aktive_WP = archivedWP.Id;
            archiv.Color_1 = archivedWP.Color_1;
            archiv.Color_2 = archivedWP.Color_2;
            archiv.Color_3 = archivedWP.Color_3;
            archiv.Color_4 = archivedWP.Color_4;
            archiv.Color_5 = archivedWP.Color_5;
            archiv.Color_6 = archivedWP.Color_6;
            archiv.Column_1 = archivedWP.Column_1;
            archiv.Column_2 = archivedWP.Column_2;
            archiv.Column_3 = archivedWP.Column_3;
            archiv.Column_4 = archivedWP.Column_4;
            archiv.Column_5 = archivedWP.Column_5;
            archiv.Column_6 = archivedWP.Column_6;
            archiv.Column_7 = archivedWP.Column_7;
            archiv.Column_8 = archivedWP.Column_8;
            archiv.Column_9 = archivedWP.Column_9;
            archiv.Column_10 = archivedWP.Column_10;
            archiv.Column_11 = archivedWP.Column_11;
            archiv.Column_12 = archivedWP.Column_12;
            archiv.Column_13 = archivedWP.Column_13;
            archiv.Column_14 = archivedWP.Column_14;
            archiv.Column_15 = archivedWP.Column_15;
            archiv.Column_16 = archivedWP.Column_16;
            archiv.Column_17 = archivedWP.Column_17;
            archiv.Column_18 = archivedWP.Column_18;
            archiv.Column_19 = archivedWP.Column_19;
            archiv.Column_20 = archivedWP.Column_20;
            archiv.Column_21 = archivedWP.Column_21;
            archiv.Column_22 = archivedWP.Column_22;
            archiv.Column_23 = archivedWP.Column_23;
            archiv.Column_24 = archivedWP.Column_24;
            archiv.Column_25 = archivedWP.Column_25;
            archiv.Column_26 = archivedWP.Column_26;
            archiv.Column_27 = archivedWP.Column_27;
            archiv.Column_28 = archivedWP.Column_28;
            archiv.Column_29 = archivedWP.Column_29;
            archiv.Column_30 = archivedWP.Column_30;

            db.Archiv.Add(archiv);
            db.Aktive_WP.Remove(archivedWP);
            db.SaveChanges();
            return RedirectToAction("WP", "WP");
        }

        [HttpGet]
        public ActionResult Export(string searching, bool? all, bool? c1ch, bool? c2ch, bool? c3ch, bool? c4ch, bool? c5ch, bool? c6ch, bool? c7ch, bool? c8ch, bool? c9ch, bool? c10ch, bool? c11ch, bool? c12ch, bool? c13ch, bool? c14ch, bool? c15ch, bool? c16ch, bool? c17ch, bool? c18ch, bool? c19ch, bool? c20ch, bool? c21ch, bool? c22ch, bool? c23ch, bool? c24ch, bool? c25ch, bool? c26ch, bool? c27ch, bool? c28ch, bool? c29ch, bool? c30ch)
        {
            List<Archiv> ExportArray = db.Archiv.ToList();
            if ((searching == null ? 0 : searching.Length) > 0)
            {
                List<Archiv> pom = new List<Archiv>();
                for (int j = 0; j < ExportArray.Count(); j++)
                {
                    int a = ExportArray[j].Id;
                    if ((ExportArray[j].Column_1 != null && (all == true || c1ch == true) && ExportArray[j].Column_1.ToLower().Contains(searching.ToLower())) ||
                       (ExportArray[j].Column_2 != null && (all == true || c2ch == true) && ExportArray[j].Column_2.ToLower().Contains(searching.ToLower())) ||
                       (ExportArray[j].Column_3 != null && (all == true || c3ch == true) && ExportArray[j].Column_3.ToLower().Contains(searching.ToLower())) ||
                       (ExportArray[j].Column_4 != null && (all == true || c4ch == true) && ExportArray[j].Column_4.ToLower().Contains(searching.ToLower())) ||
                       (ExportArray[j].Column_5 != null && (all == true || c5ch == true) && ExportArray[j].Column_5.ToLower().Contains(searching.ToLower())) ||
                       (ExportArray[j].Column_6 != null && (all == true || c6ch == true) && ExportArray[j].Column_6.ToLower().Contains(searching.ToLower())) ||
                       (ExportArray[j].Column_7 != null && (all == true || c7ch == true) && ExportArray[j].Column_7.ToLower().Contains(searching.ToLower())) ||
                       (ExportArray[j].Column_8 != null && (all == true || c8ch == true) && ExportArray[j].Column_8.ToLower().Contains(searching.ToLower())) ||
                       (ExportArray[j].Column_9 != null && (all == true || c9ch == true) && ExportArray[j].Column_9.ToLower().Contains(searching.ToLower())) ||
                       (ExportArray[j].Column_10 != null && (all == true || c10ch == true) && ExportArray[j].Column_10.ToLower().Contains(searching.ToLower())) ||
                       (ExportArray[j].Column_11 != null && (all == true || c11ch == true) && ExportArray[j].Column_11.ToLower().Contains(searching.ToLower())) ||
                       (ExportArray[j].Column_12 != null && (all == true || c12ch == true) && ExportArray[j].Column_12.ToLower().Contains(searching.ToLower())) ||
                       (ExportArray[j].Column_13 != null && (all == true || c13ch == true) && ExportArray[j].Column_13.ToLower().Contains(searching.ToLower())) ||
                       (ExportArray[j].Column_14 != null && (all == true || c14ch == true) && ExportArray[j].Column_14.ToLower().Contains(searching.ToLower())) ||
                       (ExportArray[j].Column_15 != null && (all == true || c15ch == true) && ExportArray[j].Column_15.ToLower().Contains(searching.ToLower())) ||
                       (ExportArray[j].Column_16 != null && (all == true || c16ch == true) && ExportArray[j].Column_16.ToLower().Contains(searching.ToLower())) ||
                       (ExportArray[j].Column_17 != null && (all == true || c17ch == true) && ExportArray[j].Column_17.ToLower().Contains(searching.ToLower())) ||
                       (ExportArray[j].Column_18 != null && (all == true || c18ch == true) && ExportArray[j].Column_18.ToLower().Contains(searching.ToLower())) ||
                       (ExportArray[j].Column_19 != null && (all == true || c19ch == true) && ExportArray[j].Column_19.ToLower().Contains(searching.ToLower())) ||
                       (ExportArray[j].Column_20 != null && (all == true || c20ch == true) && ExportArray[j].Column_20.ToLower().Contains(searching.ToLower())) ||
                       (ExportArray[j].Column_21 != null && (all == true || c21ch == true) && ExportArray[j].Column_21.ToLower().Contains(searching.ToLower())) ||
                       (ExportArray[j].Column_22 != null && (all == true || c22ch == true) && ExportArray[j].Column_22.ToLower().Contains(searching.ToLower())) ||
                       (ExportArray[j].Column_23 != null && (all == true || c23ch == true) && ExportArray[j].Column_23.ToLower().Contains(searching.ToLower())) ||
                       (ExportArray[j].Column_24 != null && (all == true || c24ch == true) && ExportArray[j].Column_24.ToLower().Contains(searching.ToLower())) ||
                       (ExportArray[j].Column_25 != null && (all == true || c25ch == true) && ExportArray[j].Column_25.ToLower().Contains(searching.ToLower())) ||
                       (ExportArray[j].Column_26 != null && (all == true || c26ch == true) && ExportArray[j].Column_26.ToLower().Contains(searching.ToLower())) ||
                       (ExportArray[j].Column_27 != null && (all == true || c27ch == true) && ExportArray[j].Column_27.ToLower().Contains(searching.ToLower())) ||
                       (ExportArray[j].Column_28 != null && (all == true || c28ch == true) && ExportArray[j].Column_28.ToLower().Contains(searching.ToLower())) ||
                       (ExportArray[j].Column_29 != null && (all == true || c29ch == true) && ExportArray[j].Column_29.ToLower().Contains(searching.ToLower())) ||
                       (ExportArray[j].Column_30 != null && (all == true || c30ch == true) && ExportArray[j].Column_30.ToLower().Contains(searching.ToLower())))

                    {
                        pom.AddRange(db.Archiv.Where(s => s.Id == a).DefaultIfEmpty().ToList());
                    }
                }
                List<Archiv> k = ExportArray.Intersect(pom).ToList();
                ExportArray = k;
            }




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

            foreach (var item in ExportArray)
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
            sheet.Name = "Archiv najdi=" + searching;
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
    }
}