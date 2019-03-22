using piloteProject.Context;
using piloteProject.Models;
using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace piloteProject.Controllers
{
    public class NotesController : Controller
    {
        private MyDbContext db = new MyDbContext();
        // GET: Note
        public ActionResult Index()
        {
            return View();
        }

        [HttpGet]
        public ActionResult NoteProject()
        {
            string noteId = Request.QueryString["NoteId"];
            string ProjectId = Request.QueryString["ProjectId"];
            ViewBag.ProjectId = ProjectId;
            Note note = new Note();
            ViewBag.note = note;
            ViewBag.selected = "selected";
            if (Request.QueryString["NoteId"] != null)
            {
                int Id = Convert.ToInt32(Request.QueryString["NoteId"]);
                note = db.Note.Find(Id);
                ViewBag.note = note;
            }

            return View();
        }

        [HttpPost]
        public ActionResult NoteProject([Bind(Include = "Id,ProjectId,ImpactSi,MacrosImpact,Enjeux,Satisfaction,Effet,Somme")] Note note)
        {
            int somme = Int32.Parse(note.ImpactSi) + Int32.Parse(note.MacrosImpact) + Int32.Parse(note.Enjeux) + Int32.Parse(note.Satisfaction) + Int32.Parse(note.Effet);
            note.Somme = somme;
            Project project = db.Projects.Find(note.ProjectId);
            if (note.Id != 0)
            {
                db.Entry(note).State = EntityState.Modified;
                project.NoteId = note.Id;
            }
            else
            {
                //Save Note in db Note
                db.Note.Add(note);
                db.SaveChanges();
                project.NoteId = note.Id;
            }
            // save Note id in db Project

            project.SommeNote = somme;
            db.Entry(project).State = EntityState.Modified;
            db.SaveChanges();

            return RedirectToAction("Index", "Home");
        }
    }
}