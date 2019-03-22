using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace piloteProject.Models
{
    public class Note
    {
        public int Id { get; set; }
        public int ProjectId { get; set; }
        public string ImpactSi { get; set; }
        public string MacrosImpact { get; set; }
        public string Enjeux { get; set; }
        public string Satisfaction { get; set; }
        public string Effet { get; set; }
        public int Somme { get; set; }
    }
}