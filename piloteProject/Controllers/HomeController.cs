
using Pechkin;
using piloteProject.Context;
using piloteProject.Models;
using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.IO;
using System.Linq;
using System.Web.Mvc;


using System.Web.UI.WebControls;


namespace piloteProject.Controllers
{
    public class HomeController : Controller
    {

        private MyDbContext db = new MyDbContext();
        // GET: Projects
        public ActionResult Index()
        {
            return View(db.Projects.ToList());
        }

        public ActionResult PTFProjectDecider()
        {
            using (db)
            {
                var projectsEcarter = db.Projects.SqlQuery("SELECT * FROM projects WHERE decision='ecarte'").ToList();
                ViewBag.projectsEcarter = projectsEcarter;
                var projectsDemarer = db.Projects.SqlQuery("SELECT * FROM projects WHERE decision='demarer'").ToList();
                ViewBag.projectsDemarer = projectsDemarer;
                var projectsEnAttente = db.Projects.SqlQuery("SELECT * FROM projects WHERE decision='miseenattente'").ToList();
                ViewBag.projectsEnAttente = projectsEnAttente;

            }
            ViewBag.selected = "selected";

            return View();
        }

        [HttpGet]
        public ActionResult AddProjectFormulaire()
        {
            if (Request.QueryString["Id"] != null)
            {
                int Id = Convert.ToInt32(Request.QueryString["Id"]);
                Project project = db.Projects.Find(Id);
                ViewBag.project = project;
            }
            else
            {
                ViewBag.project = null;
            }
            ViewBag.Directions = db.Directions.ToList();
            ViewBag.Services = db.Services.ToList();

            ViewBag.selected = "selected";
            ViewBag.Checked = "checked";

            return View();
        }

        [HttpPost]
        public ActionResult AddProjectFormulaire([Bind(Include = "Id,Name,Redacteur,Directions,Services,ProjectConsistance,DescriptifSouhait,DescriptifGain,DescriptifMoyen,DescriptifChefDeProjet,DescriptifDelais,DescriptifDifficulte,ProjectConsistance,Consistance,Initiative,Technologie,Applicatif,Transversalite,AutreInitiative,AutreDirections,SommeNote,NoteId,DateCreattion")] Project project)
        {
            if (project.Id != 0)
            {
                db.Entry(project).State = EntityState.Modified;
                db.SaveChanges();
            }
            else
            {
                project.DateCreattion = DateTime.Today;
                project.DateDemarrage = DateTime.Today;
                db.Projects.Add(project);
                db.SaveChanges();
            }
            return RedirectToAction("Index");
        }

        public ActionResult ProjectDecision()
        {
            ViewBag.ProjectId = Request.QueryString["Id"];
            return View();
        }

        [HttpPost]
        public ActionResult ProjectDecision(String Desicion)
        {
            string desicion = Request.Form["decision"];
            int projectId = Convert.ToInt32(Request.Form["projectId"]);
            Project project = db.Projects.Find(projectId);
            project.Decision = desicion;
            project.DateDemarrage = DateTime.Today;
            db.Entry(project).State = EntityState.Modified;
            db.SaveChanges();

            return RedirectToAction("Index");
        }

        public ActionResult ProjectStatus()
        {
            string status = Request.QueryString["status"];
            int projectId = Convert.ToInt32(Request.QueryString["Id"]);
            Project project = db.Projects.Find(projectId);
            project.Status = status;
            db.Entry(project).State = EntityState.Modified;
            db.SaveChanges();
            return RedirectToAction("Index", "PTFProjectDecider");
        }

        public ActionResult ProjectDelete()
        {
            int projectId = Convert.ToInt32(Request.QueryString["Id"]);
            Project project = db.Projects.Find(projectId);
            db.Projects.Remove(project);
            db.SaveChanges();
            return RedirectToAction("Index");
        }

        protected void btnClick_Click(object sender, EventArgs e)
        {
            DownloadAsPDF();
        }

        public ActionResult DownloadAsPDF()
        {
            List<Project> projectsDemarer = new List<Project>();
            using (db)
            {
                projectsDemarer = db.Projects.SqlQuery("SELECT * FROM projects WHERE decision='demarer'").ToList();
            }
            string direction = "";

            string html2 = "";
            string html = "<html style='position: relative;min-height: 100%;'><head></head><body><header></header><h3 style = 'text-align:center' > Mes projets </h3>" +
                "<table class='table table-striped' style='width: 100%;max-width: 100%;margin-bottom: 20px;margin-top: 20%;border-spacing:0;border-collapse: collapse;'>" +
                "<thead style = 'display: table-header-group;vertical-align: middle;border-color: inherit;' >" +
                "<tr style='display: table-row;vertical-align: inherit;border-color:inherit;'>" +
                " <th style = 'vertical-align: line-height: 1.42857143;padding: 8px;bottom;border-bottom: 2px solid #ddd;font-size: 15px;text-align: center;' class='tableauTitre'>Titre</th>" +
                "<th style = 'vertical-align: line-height: 1.42857143;padding: 8px;bottom;border-bottom: 2px solid #ddd;font-size: 15px;text-align: center;' class='tableauTitre'>Direction(s)</th>" +
                "<th style = 'vertical-align: line-height: 1.42857143;padding: 8px;bottom;border-bottom: 2px solid #ddd;font-size: 15px;text-align: center;' class='tableauTitre'>Note</th>" +
                "<th style = 'vertical-align: line-height: 1.42857143;padding: 8px;bottom;border-bottom: 2px solid #ddd;font-size: 15px;text-align: center;' class='tableauTitre'>Date de démarrage</th>" +
                "<th style = 'vertical-align: line-height: 1.42857143;padding: 8px;bottom;border-bottom: 2px solid #ddd;font-size: 15px;text-align: center;' class='tableauTitre'>Statut</th>" +
                "</tr>" +
                "</thead>" +
                "<tbody style = 'padding: 8px;' > ";
            foreach (var item in projectsDemarer)
            {
                if(item.Directions!= null)
                {
                    direction = item.Directions.Replace(";", " ");
                }
                html2 = html2 + "<tr style='display: table-row;vertical-align: inherit;border-color:inherit;'>" +
                "<td style='padding:8px;line-height:1.42857143;border-top:1px; border-left:3px; solid #ddd;vertical-align: bottom;border-bottom: 2px solid #ddd;font-size: 15px;text-align: center;' class='elementTableau'>" + item.Name + "</td>" +
                "<td style='padding:8px;line-height:1.42857143;border-top:1px border-left:3px; solid #ddd;vertical-align: bottom;border-bottom: 2px solid #ddd;font-size: 15px;text-align: center;' class='elementTableau'>" + direction + "</td>" +
                "<td style='padding:8px;line-height:1.42857143;border-top:1px border-left:3px; solid #ddd;vertical-align: bottom;border-bottom: 2px solid #ddd;font-size: 15px;text-align: center;' class='elementTableau'>" + item.SommeNote + "/20</td>" +
                "<td style='padding:8px;line-height:1.42857143;border-top:1px border-left:3px; solid #ddd;vertical-align: bottom;border-bottom: 2px solid #ddd;font-size: 15px;text-align: center;' class='elementTableau'>" + item.DateDemarrage.ToShortDateString() + "</td>" +
                "<td style='padding:8px;line-height:1.42857143;border-top:1px border-left:3px; solid #ddd;vertical-align: bottom;border-bottom: 2px solid #ddd;font-size: 15px;text-align: center;' class='elementTableau'>" + item.Status + "</td>" +
                "</tr>";
            }
            string html3 = html + html2 + "</tbody></table><footer style='position:absolute;bottom:0;width:100%;height:60px;'></footer></body></html>";
            byte[] pdfContent = new SimplePechkin(new GlobalConfig()).Convert(html3);
           
            Response.Clear();
            Response.ContentType = "application/pdf";
            Response.AddHeader("Content-Disposition",
                "attachment;filename=\"MesProjets.pdf\"");
            Response.BinaryWrite(pdfContent);
            Response.Flush();
            Response.End();


            return RedirectToAction("Index");
        }
    }
}