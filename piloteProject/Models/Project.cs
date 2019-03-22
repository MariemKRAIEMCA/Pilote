using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace piloteProject.Models
{
    public class Project
    {
        public int Id { get; set; }
        public string Name { get; set; }
        public string Redacteur { get; set; }
        public string Directions { get; set; }
        public string Services { get; set; }
        public DateTime DateCreattion { get; set; }
        /**********************************************************/
        public string DescriptifSouhait { get; set; }
        public string DescriptifGain { get; set; }
        public string DescriptifMoyen { get; set; }
        public string DescriptifChefDeProjet { get; set; }
        public string DescriptifDelais { get; set; }
        public string DescriptifDifficulte { get; set; }
        /**********************************************************/
        public string Consistance { get; set; }
        public string ProjectConsistance { get; set; }
        /**********************************************************/
        public string Initiative { get; set; }
        public string AutreInitiative { get; set; }
        /**********************************************************/
        public string Technologie { get; set; }
        /**********************************************************/
        public string Applicatif { get; set; }
        /**********************************************************/
        public string Transversalite { get; set; }
        public string AutreDirections { get; set; }
        /**********************************************************/
        public string Decision { get; set; }
        public int NoteId { get; set; }
        public DateTime DateDemarrage { get; set; }
        public string Status { get; set; }
        public int SommeNote { get; set; }
    }
}