using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace piloteProject.Models
{
    public class User
    {
        public int Id { get; set; }
        public string Name { get; set; }
        public string Pseudo { get; set; }
        public string Email { get; set; }
        public string Pwd { get; set; }
        public string Sexe { get; set; }
    }
}