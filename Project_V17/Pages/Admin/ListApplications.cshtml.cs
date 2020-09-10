using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.EntityFrameworkCore;
using Project_V17.Data;
using Project_V17.Models;

namespace Project_V17.Pages
{
    public class ListApplicationsModel : PageModel
    {
        private readonly Project_V17.Data.ApplicationDbContext _context;

        public ListApplicationsModel(Project_V17.Data.ApplicationDbContext context)
        {
            _context = context;
        }

        public IList<FSApp> FSApp { get;set; }

        public async Task OnGetAsync()
        {
            FSApp = await _context.FSApp.ToListAsync();
        }
    }
}
