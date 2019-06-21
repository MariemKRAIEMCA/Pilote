
using Pechkin;
using piloteProject.Context;
using piloteProject.Models;
using Syncfusion.Presentation;
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
            ViewBag.Checked = "checked";

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
                ViewBag.DateDemarrage = project.DateDemarrage.ToString("yyyy-MM-dd");
                ViewBag.DatedeFin = project.DatedeFin.ToString("yyyy-MM-dd");

            }
            else
            {
                ViewBag.project = null;
            }
            if (Request.QueryString["decide"] != null)
            {
                ViewBag.page = "decider";
            } else
            {
                ViewBag.page = null;
            }
            ViewBag.Directions = db.Directions.ToList();
            ViewBag.Services = db.Services.ToList();

            ViewBag.selected = "selected";
            ViewBag.Checked = "checked";

            return View();
        }

        [HttpPost]
        public ActionResult AddProjectFormulaire([Bind(Include = "Id,Name,Redacteur,Directions,Services,ProjectConsistance,DescriptifSouhait,DescriptifGain,DescriptifMoyen,DescriptifChefDeProjet,DescriptifDelais,DescriptifDifficulte,gainEtp,ChargeJh,ProjectConsistance,Consistance,Initiative,Technologie,Applicatif,Transversalite,AutreInitiative,AutreDirections,SommeNote,NoteId,DateCreattion,Decision,Status,DateDemarrage,Pilote,Sponsor,Commentaires,TypeProjet,Meteo,DatedeFin,Impact,Marche,InstanceSuivie,TransCR,Typologie")] Project project)
        {
            project.DateDemarrage = Convert.ToDateTime(project.DateDemarrage);
            project.DatedeFin = Convert.ToDateTime(project.DatedeFin);

            if (project.Id != 0)
            {
                db.Entry(project).State = EntityState.Modified;
                db.SaveChanges();
            }
            else
            {
                project.DateCreattion = DateTime.Today;
                db.Projects.Add(project);
                db.SaveChanges();
            }

            string page = Request.Form["Page"];
            if (page != "")
            {
                return RedirectToAction("PTFProjectDecider");
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
        public ActionResult ProjectMeteo()
        {
            string meteo = Request.QueryString["meteo"];
            int projectId = Convert.ToInt32(Request.QueryString["Id"]);
            Project project = db.Projects.Find(projectId);
            project.Meteo = meteo;
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

        public void addPremierRow(ITable table)
        {
            ICell cell = table[0, 0];
            IParagraph parg = cell.TextBody.AddParagraph("Projet \n (nom + descriptive)");
            parg.HorizontalAlignment = HorizontalAlignmentType.Center;
            IFont font1 = parg.Font;
            font1.FontName = "Century Gothic (En-têtes)";
            font1.FontSize = 12;
            
            cell.ColumnWidth = 100;
            cell = table[0, 1];
            cell.ColumnWidth = 100;
            parg =  cell.TextBody.AddParagraph("Sponsor(s)");
            parg.HorizontalAlignment = HorizontalAlignmentType.Center;
            font1 = parg.Font;
            font1.FontName = "Century Gothic (En-têtes)";
            font1.FontSize = 12;

            cell = table[0, 2];
            cell.ColumnWidth = 100;
            parg = cell.TextBody.AddParagraph("Pilote(s)");
            parg.HorizontalAlignment = HorizontalAlignmentType.Center;
            font1 = parg.Font;
            font1.FontName = "Century Gothic (En-têtes)";
            font1.FontSize = 12;

            cell = table[0, 3];
            cell.ColumnWidth = 100;
            parg = cell.TextBody.AddParagraph("Chef de projet");
            font1 = parg.Font;
            font1.FontName = "Century Gothic (En-têtes)";
            font1.FontSize = 12;

            cell = table[0, 3];
            cell.ColumnWidth = 100;
            cell = table[0, 4];
            cell.ColumnWidth = 100;
            parg = cell.TextBody.AddParagraph("Météo");
            font1 = parg.Font;
            font1.FontName = "Century Gothic (En-têtes)";
            font1.FontSize = 12;

            cell = table[0, 3];
            cell.ColumnWidth = 100;
            cell = table[0, 5];
            cell.ColumnWidth = 100;
            parg = cell.TextBody.AddParagraph("Étapes C/A/R/T/E/S");
            font1 = parg.Font;
            font1.FontName = "Century Gothic (En-têtes)";
            font1.FontSize = 12;

            cell = table[0, 3];
            cell.ColumnWidth = 100;
            cell = table[0, 6];
            cell.ColumnWidth = 100;
            parg = cell.TextBody.AddParagraph("Commentaire");//\n(principales réalisations, prochaines étapres, alerte, point de vigilance
            font1 = parg.Font;
            font1.FontName = "Century Gothic (En-têtes)";
            font1.FontSize = 12;
            cell = table[0, 3];
            cell.ColumnWidth = 100;

        }

        public void header(IPresentation pptxDoc, int slidenb)
        {
            //titre
            IShape EtpText = pptxDoc.Slides[slidenb].AddTextBox(250, 20, 500, 200);

            IParagraph EtpPa = EtpText.TextBody.AddParagraph();
            ITextPart EtpPar = EtpPa.AddTextPart();
            EtpPar.Text = "Météo des Projets suivis au CQEF "+DateTime.Now.ToShortDateString();
            IFont font1 = EtpPar.Font;
            font1.FontName = "Century Gothic (En-têtes)";
            font1.FontSize = 20;
            //meteo
            IShape MeteoText = pptxDoc.Slides[slidenb].AddTextBox(50, 10, 50, 50);
            IParagraph MeteoPa = MeteoText.TextBody.AddParagraph();
            ITextPart MeteoPar = MeteoPa.AddTextPart();
            MeteoPar.Text = "Météo";
            IFont font2 = MeteoPar.Font;
            font2.FontName = "Century Gothic (En-têtes)";
            font2.FontSize = 12;
            font2.Color = ColorObject.LightBlue;
            // Tres bien pic
            FileStream pictureStreamTB = new FileStream(@"C:\Users\Factory\Documents\GitProjects\Pilote\PilotePicMeteo\TresBien.png", FileMode.Open);//Background pic
            pptxDoc.Slides[slidenb].Shapes.AddPicture(pictureStreamTB, 20, 30, 21, 21);
            pictureStreamTB.Close();
            // bien pic 
            FileStream pictureStreamB = new FileStream(@"C:\Users\Factory\Documents\GitProjects\Pilote\PilotePicMeteo\Bien.png", FileMode.Open);//Background pic
            pptxDoc.Slides[slidenb].Shapes.AddPicture(pictureStreamB, 45, 30, 21, 21);
            pictureStreamB.Close();
            //Moyen Pic
            FileStream pictureStreamMo = new FileStream(@"C:\Users\Factory\Documents\GitProjects\Pilote\PilotePicMeteo\MoyenMeteo.png", FileMode.Open);//Background pic
            pptxDoc.Slides[slidenb].Shapes.AddPicture(pictureStreamMo, 70, 30, 21, 21);
            pictureStreamMo.Close();
            //Mauvais pic
            FileStream pictureStreamMa = new FileStream(@"C:\Users\Factory\Documents\GitProjects\Pilote\PilotePicMeteo\Mauvais.png", FileMode.Open);//Background pic
            pptxDoc.Slides[slidenb].Shapes.AddPicture(pictureStreamMa, 95, 30, 21, 21);
            pictureStreamMa.Close();


        }
        public void Addtext(IPresentation pptxDoc, int slidenb, int x, int y , int largeX, int largeY, int taille, String text )
        {
            IShape CltText = pptxDoc.Slides[slidenb].AddTextBox(x, y, largeX, largeY);
            IParagraph CltPa = CltText.TextBody.AddParagraph();
            ITextPart CltPar = CltPa.AddTextPart();
            CltPar.Text = text;
            IFont font1 = CltPar.Font;
            font1.FontName = "Century Gothic (En-têtes)";
            font1.FontSize = taille;
        }
        public void tableauPlanning(IPresentation pptxDoc, int slidenb)
        {
            /***************************NvOffre***************************/
            IShape NvOffre = pptxDoc.Slides[slidenb].Shapes.AddShape(AutoShapeType.Rectangle, 580, 10, 120, 20);
            NvOffre.Fill.FillType = FillType.Solid;
            NvOffre.Fill.SolidFill.Color = ColorObject.FromArgb(1, 091, 155, 213);
            NvOffre.LineFormat.Fill.SolidFill.Color = ColorObject.Transparent;
            NvOffre.LineFormat.Weight = 0.01;
            IParagraph paragNv = NvOffre.TextBody.AddParagraph("Nouvelle offre");
            IFont fontP1 = paragNv.Font;
            fontP1.FontSize = 10;
            paragNv.HorizontalAlignment = HorizontalAlignmentType.Center;

            /***************************NvOutil***************************/
            IShape NvOtil = pptxDoc.Slides[slidenb].Shapes.AddShape(AutoShapeType.Rectangle, 705, 10, 120, 20);
            NvOtil.Fill.FillType = FillType.Solid;
            NvOtil.Fill.SolidFill.Color = ColorObject.FromArgb(1, 237, 125, 049);
            NvOtil.LineFormat.Fill.SolidFill.Color = ColorObject.Transparent;
            NvOtil.LineFormat.Weight = 0.1;
            IParagraph paragNvO =  NvOtil.TextBody.AddParagraph("Nouvel outil");
            fontP1 = paragNvO.Font;
            fontP1.FontSize = 10;
            paragNvO.HorizontalAlignment = HorizontalAlignmentType.Center;
            /***************************NvReg***************************/
            IShape NvReg = pptxDoc.Slides[slidenb].Shapes.AddShape(AutoShapeType.Rectangle, 580, 35, 120, 20);
            NvReg.Fill.FillType = FillType.Solid;
            NvReg.Fill.SolidFill.Color = ColorObject.FromArgb(1, 255, 192, 000);
            NvReg.LineFormat.Fill.SolidFill.Color = ColorObject.Transparent;
            NvReg.LineFormat.Weight = 0.1;
            IParagraph paragNReg = NvReg.TextBody.AddParagraph("Nouvel Réglementation");
            fontP1 = paragNReg.Font;
            fontP1.FontSize = 10;
            paragNReg.HorizontalAlignment = HorizontalAlignmentType.Center;
            /***************************Digital***************************/
            IShape Digital = pptxDoc.Slides[slidenb].Shapes.AddShape(AutoShapeType.Rectangle, 705, 35, 120, 20);
            Digital.Fill.FillType = FillType.Solid;
            Digital.Fill.SolidFill.Color = ColorObject.FromArgb(1, 112, 173, 071);
            Digital.LineFormat.Fill.SolidFill.Color = ColorObject.Transparent;
            Digital.LineFormat.Weight = 0.1;
            IParagraph paragDig = Digital.TextBody.AddParagraph("Digital");
            fontP1 = paragDig.Font;
            fontP1.FontSize = 10;
            paragDig.HorizontalAlignment = HorizontalAlignmentType.Center;
            /******************************************************/
            Addtext(pptxDoc, slidenb, 830, 10, 130, 25, 9, "Niveau de transformation en CR :");

            FileStream pictureStreamMa = new FileStream(@"C:\Users\Factory\Documents\GitProjects\Pilote\PilotePicMeteo\fort.png", FileMode.Open);//Background pic
            pptxDoc.Slides[slidenb].Shapes.AddPicture(pictureStreamMa, 835, 35, 8, 18);
            pictureStreamMa.Close();
            Addtext(pptxDoc, slidenb, 840, 40, 30, 10, 6, "Fort");

            pictureStreamMa = new FileStream(@"C:\Users\Factory\Documents\GitProjects\Pilote\PilotePicMeteo\moyen.png", FileMode.Open);//Background pic
            pptxDoc.Slides[slidenb].Shapes.AddPicture(pictureStreamMa, 870, 35, 8, 18);
            pictureStreamMa.Close();
            Addtext(pptxDoc, slidenb, 875, 40, 50, 10, 6, "Moyen");

            pictureStreamMa = new FileStream(@"C:\Users\Factory\Documents\GitProjects\Pilote\PilotePicMeteo\faible.png", FileMode.Open);//Background pic
            pptxDoc.Slides[slidenb].Shapes.AddPicture(pictureStreamMa, 910, 35, 8, 18);
            pictureStreamMa.Close();
            Addtext(pptxDoc, slidenb, 915, 40, 50, 10, 6, "Faible");
            /**************************C/R/S***********************/
            Addtext(pptxDoc, slidenb, 5, 165, 65, 30, 12, "Clients");
            IShape shapeLine = pptxDoc.Slides[slidenb].Shapes.AddShape(AutoShapeType.Line, 75, 245, 880, 0);
            shapeLine.LineFormat.DashStyle = LineDashStyle.Dash;
            shapeLine.LineFormat.Weight = 1.5;
            Addtext(pptxDoc, slidenb, 5, 295, 65, 30, 12, "Réseaux");
            shapeLine = pptxDoc.Slides[slidenb].Shapes.AddShape(AutoShapeType.Line, 75, 375, 880, 0);
            shapeLine.LineFormat.DashStyle = LineDashStyle.Dash;
            shapeLine.LineFormat.Weight = 1.5;
            Addtext(pptxDoc, slidenb, 5, 425, 50, 30, 12, "Site");

            /*************************Dates****************************/
            Addtext(pptxDoc, slidenb, 75, 80, 65, 25, 12, "Janvier");
            Addtext(pptxDoc, slidenb, 150, 80, 65, 25, 12, "Février");
            Addtext(pptxDoc, slidenb, 225, 80, 65, 25, 12, "mars");
            Addtext(pptxDoc, slidenb, 300, 80, 65, 25, 12, "Avril");
            Addtext(pptxDoc, slidenb, 375, 80, 65, 25, 12, "Mai");
            Addtext(pptxDoc, slidenb, 450, 80, 65, 25, 12, "Juin");
            Addtext(pptxDoc, slidenb, 525, 80, 65, 25, 12, "Juillet");
            Addtext(pptxDoc, slidenb, 600, 80, 65, 25, 12, "Aout");

            Addtext(pptxDoc, slidenb, 675, 80, 65, 25, 10, "Septembre");
            Addtext(pptxDoc, slidenb, 750, 80, 65, 25, 10, "Octobre");
            Addtext(pptxDoc, slidenb, 825, 80, 65, 25, 10, "Novembre");
            Addtext(pptxDoc, slidenb, 900, 80, 65, 25, 10, "Décembre");
        }   

        public void addProjectWithDate(List<Project> projectList, IPresentation pptxDoc, int pptNb , int y)
        {
            int pointDebut = 50;
            int margeDebut = 0;
            int R =0, G=0, B=0;
            int yInitiale = y;
            List<Block> blocksList = new List<Block>();
            double TailleDurée =0 ;
            foreach (Project project in projectList)
            {
                y = yInitiale;
                int totalDay = Convert.ToInt32((project.DatedeFin - project.DateDemarrage).TotalDays);
                int day = project.DateDemarrage.Day;
                int mois = project.DateDemarrage.Month;
               
                if (day < 10)
                {
                    margeDebut = 0;
                }
                else if (day < 20)
                {
                    margeDebut = 21;
                }
                else
                {
                    margeDebut = 43;
                }

                if (project.Typologie.Equals("NouvelleOffre"))
                {
                    R = 091;
                    G = 155;
                    B = 213;

                }
                else if (project.Typologie.Equals("NouvelOutil"))
                {
                    R = 237;
                    G = 125;
                    B = 049;

                }
                else if (project.Typologie.Equals("NouvelleReglementation"))
                {
                    R = 255;
                    G = 192;
                    B = 000;
                }
                else if (project.Typologie.Equals("Digital"))
                {
                    R = 112;
                    G = 173;
                    B = 071;
                }
                else
                {
                    R = 150;
                    G = 150;
                    B = 150;
                }

                pointDebut = (mois * 75) + margeDebut;// point de debut
                TailleDurée = 2.5 * totalDay; // durée
                                              /****************************************verif Possibilité********************************************/
                bool verif = true;
                do
                {
                    y += 22;
                    verif = verifFree(blocksList, pointDebut, y, TailleDurée);
                }
                while(!verif);

                Block blockCordonnée = new Block();
                blockCordonnée.X = pointDebut;
                blockCordonnée.Y = y;
                blockCordonnée.Duree = TailleDurée;
                blocksList.Add(blockCordonnée);

                /*********************************************block*********************************************/
                IShape NvReg = pptxDoc.Slides[pptNb].Shapes.AddShape(AutoShapeType.Rectangle, pointDebut, y, TailleDurée, 20);
                //ajouter l'incone niveau de transaformation
                if(project.TransCR != null)
                {
                    FileStream pictureStreamMa = new FileStream(@"C:\Users\Factory\Documents\GitProjects\Pilote\PilotePicMeteo\"+project.TransCR+".png", FileMode.Open);//Background pic
                    pptxDoc.Slides[pptNb].Shapes.AddPicture(pictureStreamMa, pointDebut + 2, y+2, 7, 15);
                    pictureStreamMa.Close();
                }


                NvReg.Fill.FillType = FillType.Solid;
                NvReg.Fill.SolidFill.Color = ColorObject.FromArgb(1, R, G, B);
                NvReg.LineFormat.Fill.SolidFill.Color = ColorObject.Transparent;
                NvReg.LineFormat.Weight = 0.1;
                IParagraph paragNReg = NvReg.TextBody.AddParagraph(project.Name);
                IFont fontP1 = paragNReg.Font;
                fontP1 = paragNReg.Font;
                fontP1.FontSize = 10;
                paragNReg.HorizontalAlignment = HorizontalAlignmentType.Center;
            }
        }
        public ActionResult DownloadAsPDF()
        {
            List<Project> projectsPrivatifDemarer;
            List<Project> projectsFactoryDemarer;
            List<Project> projectsNatiauxDemarer;

            List<Project> projectsProfClient;
            List<Project> projectsProfRes;
            List<Project> projectsProfSite;

            List<Project> projectsPartClient;
            List<Project> projectsPartRes;
            List<Project> projectsPartSite;

            List<Project> projectsEntClient;
            List<Project> projectsEntRes;
            List<Project> projectsEntSite;




            using (db)
            {
                projectsPrivatifDemarer = db.Projects.SqlQuery("SELECT * FROM projects WHERE decision='demarer' AND TypeProjet ='Projet privatif'").ToList();
                projectsFactoryDemarer = db.Projects.SqlQuery("SELECT * FROM projects WHERE decision='demarer' AND TypeProjet ='Outil FACTORY' ").ToList();//
                projectsNatiauxDemarer = db.Projects.SqlQuery("SELECT * FROM projects WHERE decision='demarer' AND TypeProjet ='Projet national'").ToList();

                projectsProfClient = db.Projects.SqlQuery("SELECT * FROM projects WHERE decision='demarer' AND Marche ='Professionnels' AND Impact ='Client'").ToList();
                projectsProfRes = db.Projects.SqlQuery("SELECT * FROM projects WHERE decision='demarer' AND Marche ='Professionnels' AND Impact ='Reseau'").ToList();
                projectsProfSite = db.Projects.SqlQuery("SELECT * FROM projects WHERE decision='demarer' AND Marche ='Professionnels' AND Impact ='Site'").ToList();

                projectsPartClient = db.Projects.SqlQuery("SELECT * FROM projects WHERE decision='demarer' AND Marche ='Particuliers' AND Impact ='Client'").ToList();
                projectsPartRes = db.Projects.SqlQuery("SELECT * FROM projects WHERE decision='demarer' AND Marche ='Particuliers' AND Impact ='Reseau'").ToList();
                projectsPartSite = db.Projects.SqlQuery("SELECT * FROM projects WHERE decision='demarer' AND Marche ='Particuliers' AND Impact ='Site'").ToList();

                projectsEntClient = db.Projects.SqlQuery("SELECT * FROM projects WHERE decision='demarer' AND Marche ='Entreprises' AND Impact ='Client'").ToList();
                projectsEntRes = db.Projects.SqlQuery("SELECT * FROM projects WHERE decision='demarer' AND Marche ='Entreprises' AND Impact ='Reseau'").ToList();
                projectsEntSite = db.Projects.SqlQuery("SELECT * FROM projects WHERE decision='demarer' AND Marche ='Entreprises' AND Impact ='Site'").ToList();

            }


            IPresentation pptxDoc = Presentation.Create();
            //Add slide to the presentation
            addFullTable(pptxDoc, projectsPrivatifDemarer, "Privatif", BuiltInTableStyle.LightStyle1Accent6);
            addFullTable(pptxDoc, projectsFactoryDemarer, "Factory", BuiltInTableStyle.LightStyle1Accent4);
            addFullTable(pptxDoc, projectsNatiauxDemarer, "Nationaux", BuiltInTableStyle.LightStyle1Accent1);


            // slide Marché prefessionnels
            ISlide slide4 = pptxDoc.Slides.Add(SlideLayoutType.Blank);
            tableauPlanning(pptxDoc, 3);
            IShape MarcheText = pptxDoc.Slides[3].AddTextBox(70, 20, 300, 50);
            IParagraph MarchPa = MarcheText.TextBody.AddParagraph();
            ITextPart MarchPar = MarchPa.AddTextPart();
            MarchPar.Text = "Marché des Professionnels";
            IFont font1 = MarchPar.Font;
            font1.FontName = "Century Gothic (En-têtes)";
            font1.FontSize = 22;
            font1.Color = ColorObject.FromArgb(1, 148, 194, 117);
            font1.Bold = true;
            // Marché prefesionnels
            addProjectWithDate(projectsProfClient, pptxDoc, 3, 78); //projectsProfClient
            addProjectWithDate(projectsProfRes, pptxDoc, 3, 228);//projectsProfRes
            addProjectWithDate(projectsProfSite, pptxDoc, 3, 355);//projectsProfSite
            

            // slide Marché Particuliers
            ISlide slide5 = pptxDoc.Slides.Add(SlideLayoutType.Blank);
            tableauPlanning(pptxDoc, 4);
            IShape MarcheText2 = pptxDoc.Slides[4].AddTextBox(70, 20, 300, 50);
            IParagraph MarchPa2 = MarcheText2.TextBody.AddParagraph();
            ITextPart MarchPar2 = MarchPa2.AddTextPart();
            MarchPar2.Text = "Marché des Particuliers";
            IFont font2 = MarchPar2.Font;
            font2.FontName = "Century Gothic (En-têtes)";
            font2.FontSize = 22;
            font2.Color = ColorObject.FromArgb(1, 255, 192, 000);
            font2.Bold = true;
            //Marché Particuliers

            addProjectWithDate(projectsPartClient, pptxDoc, 4, 78); //projectsProfClient
            addProjectWithDate(projectsPartRes, pptxDoc, 4, 228);//projectsProfRes
            addProjectWithDate(projectsPartSite, pptxDoc, 4, 355);//projectsProfSite

            // Marche entreprises
            ISlide slide6 = pptxDoc.Slides.Add(SlideLayoutType.Blank);
            tableauPlanning(pptxDoc, 5);
            IShape MarcheText3 = pptxDoc.Slides[5].AddTextBox(70, 20, 300, 50);
            IParagraph MarchPa3 = MarcheText3.TextBody.AddParagraph();
            ITextPart MarchPar3 = MarchPa3.AddTextPart();
            MarchPar3.Text = "Marché des Entreprises";
            IFont font3 = MarchPar3.Font;
            font3.FontName = "Century Gothic (En-têtes)";
            font3.FontSize = 22;
            font3.Color = ColorObject.FromArgb(1, 091,155,213);
            font3.Bold = true;
            //Marché Entreprise

            addProjectWithDate(projectsEntClient, pptxDoc, 3, 78); //projectsProfClient
            addProjectWithDate(projectsEntRes, pptxDoc, 3, 228);//projectsProfRes
            addProjectWithDate(projectsEntSite, pptxDoc, 3, 355);//projectsProfSite

            //Save the presentation
            pptxDoc.Save(@"C:\Users\Factory\Documents\GitProjects\Pilote\testPilotses.pptx");
            
            // debut de telechargement

            MemoryStream ms = new MemoryStream();

            using ( ms )
            using (FileStream file = new FileStream(@"C:\Users\Factory\Documents\GitProjects\Pilote\testPilotses.pptx", FileMode.Open, FileAccess.Read))
            {
                byte[] bytes = new byte[file.Length];
                file.Read(bytes, 0, (int)file.Length);
                ms.Write(bytes, 0, (int)file.Length);
                ms.Position = 0;
                byte[] buffer = new byte[(int)ms.Length];
                ms.Read(buffer, 0, (int)ms.Length);
                System.Web.HttpContext.Current.Response.Clear();
                System.Web.HttpContext.Current.Response.Buffer = true;
                System.Web.HttpContext.Current.Response.AddHeader("Content-Type", "application/pptx");
                System.Web.HttpContext.Current.Response.AddHeader("Content-Disposition", "attachment;filename=testPilotses.pptx;");

                System.Web.HttpContext.Current.Response.BinaryWrite(buffer);
                System.Web.HttpContext.Current.Response.Flush();
                System.Web.HttpContext.Current.Response.Close();
                pptxDoc.Close();
            }


            //fin de telechargement









            return RedirectToAction("Index");
        }

        public void addFullTable(IPresentation pptxDoc, List<Project> listP , string type , BuiltInTableStyle style )
        {
            ISlide slide = pptxDoc.Slides.Add(SlideLayoutType.Blank);
            int nbSlide = pptxDoc.Slides.Count();
            // projet centre loire 
            IShape MarcheText = pptxDoc.Slides[nbSlide -1].AddTextBox(90, 70, 400, 50);
            IParagraph MarchPa = MarcheText.TextBody.AddParagraph();
            ITextPart MarchPar = MarchPa.AddTextPart();
            MarchPar.Text = "Projets " + type + " Centre Loire :";
            IFont font1 = MarchPar.Font;
            font1.FontName = "Calibri (Corps)";
            font1.FontSize = 18;
            font1.Color = ColorObject.FromArgb(1, 75, 75, 75);
            font1.Bold = true;



            ITable table = slide.Shapes.AddTable(7, 7, 80, 120, 400, 250);
            table.BuiltInStyle = style;
            header(pptxDoc, nbSlide - 1);
            addPremierRow(table);
            int rowIndex = 1;
            int iconeIndex = 175;

            foreach(Project project in listP)
            {
                IRows rows = table.Rows;
                ICells cells = rows[rowIndex].Cells;
                IParagraph parg = cells[0].TextBody.AddParagraph(project.Name);
                textAlignment(parg);
                parg = cells[1].TextBody.AddParagraph(project.Sponsor);
                textAlignment(parg);
                parg = cells[2].TextBody.AddParagraph(project.Pilote);
                textAlignment(parg);
                parg = cells[3].TextBody.AddParagraph(project.DescriptifChefDeProjet);
                textAlignment(parg);
                if (project.Meteo != null)
                {
                    if (project.Meteo.Equals("TresBien"))
                    {
                        // Tres bien pic
                        FileStream pictureStreamTB = new FileStream(@"C:\Users\Factory\Documents\GitProjects\Pilote\PilotePicMeteo\TresBien.png", FileMode.Open);//Background pic
                        pptxDoc.Slides[nbSlide - 1].Shapes.AddPicture(pictureStreamTB, 525, iconeIndex, 25, 25);
                        pictureStreamTB.Close();

                    }
                    else if (project.Meteo.Equals("Bien"))
                    {
                        FileStream pictureStreamTB = new FileStream(@"C:\Users\Factory\Documents\GitProjects\Pilote\PilotePicMeteo\Bien.png", FileMode.Open);//Background pic
                        pptxDoc.Slides[nbSlide - 1].Shapes.AddPicture(pictureStreamTB, 525, iconeIndex, 25, 25);
                        pictureStreamTB.Close();

                    }
                    else if (project.Meteo.Equals("Moyen"))
                    {
                        FileStream pictureStreamTB = new FileStream(@"C:\Users\Factory\Documents\GitProjects\Pilote\PilotePicMeteo\MoyenMeteo.png", FileMode.Open);//Background pic
                        pptxDoc.Slides[nbSlide - 1].Shapes.AddPicture(pictureStreamTB, 525, iconeIndex, 25, 25);
                        pictureStreamTB.Close();

                    }
                    else if (project.Meteo.Equals("Mauvais"))
                    {
                        FileStream pictureStreamTB = new FileStream(@"C:\Users\Factory\Documents\GitProjects\Pilote\PilotePicMeteo\Mauvais.png", FileMode.Open);//Background pic
                        pptxDoc.Slides[nbSlide - 1].Shapes.AddPicture(pictureStreamTB, 525, iconeIndex, 25, 25);
                        pictureStreamTB.Close();
                    }
                }



                parg = cells[5].TextBody.AddParagraph(project.Status);
                textAlignment(parg);
                parg = cells[6].TextBody.AddParagraph(project.Commentaires);
                textAlignment(parg);
                rowIndex++;
                iconeIndex += 32;
            }
        }
        public void textAlignment(IParagraph parg)
        {
            parg.HorizontalAlignment = HorizontalAlignmentType.Center;
           
            IFont font1 = parg.Font;
            font1.FontName = "Century Gothic (En-têtes)";
            font1.FontSize = 12;
        }

        public bool verifFree(List<Block> blocksList, int x, int y , double duree)
        {
            if(blocksList.Any())
            {
                foreach(Block block in blocksList)
                {
                    if(x == block.X && y == block.Y)
                    {
                        return false;
                    }
                    if(x < block.X + block.Duree && x> block.X && y == block.Y)
                    {
                        return false;
                    }
                    if(x + duree > block.X && y == block.Y)
                    {
                        return true;
                    }
                }
            }
            return true;
        }
    }
}






/*
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

*/
/*IPresentation pptxDoc = Syncfusion.Presentation.Presentation.Create();
ISlide slide1 = pptxDoc.Slides.Add(SlideLayoutType.TitleOnly);
ITable table = slide1.Shapes.AddTable(2, 2, 100, 120, 300, 200);
int rowIndex = 0, colIndex;
foreach (IRow rows in table.Rows)
{
    colIndex = 0;

    foreach (ICell cell in rows.Cells)
    {

        cell.TextBody.AddParagraph("(" + rowIndex.ToString() + " , " + colIndex.ToString() + ")");

        colIndex++;

    }

    rowIndex++;

}
FileStream outputStream = new FileStream(@"P:\testPilotes.pptx", FileMode.Create);
pptxDoc.Save(outputStream);
//Close the PowerPoint presentation
pptxDoc.Close();*/
