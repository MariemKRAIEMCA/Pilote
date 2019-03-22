using piloteProject.Models;
using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Linq;
using System.Web;

namespace piloteProject.Context
{
    public class MyDbContext : DbContext
    {
        public DbSet<Project> Projects { get; set; }
        public DbSet<Note> Note { get; set; }
        public DbSet<Service> Services { get; set; }
        public DbSet<Direction> Directions { get; set; }

        public MyDbContext() : base()
        {

        }
    }
}