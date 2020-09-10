using System;
using System.Collections.Generic;
using System.Text;
using Microsoft.AspNetCore.Identity.EntityFrameworkCore;
using Microsoft.EntityFrameworkCore;
using Project_V17.Data.Migrations;
using Project_V17.Models;

namespace Project_V17.Data
{
    public class ApplicationDbContext : IdentityDbContext
    {
       

        public ApplicationDbContext(DbContextOptions<ApplicationDbContext> options)
            : base(options)
        {
        }

      

        public DbSet<Project_V17.Models.FSApp> FSApp { get; set; }

        public DbSet<Project_V17.Models.Department> Department { get; set; }



    }
}
