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
    public class DetailsModel : PageModel
    {
        private readonly Project_V17.Data.ApplicationDbContext _context;

        public DetailsModel(Project_V17.Data.ApplicationDbContext context)
        {
            _context = context;
        }

        public FSApp FSApp { get; set; }

        public async Task<IActionResult> OnGetAsync(int? id)
        {
            if (id == null)
            {
                return NotFound();
            }

            FSApp = await _context.FSApp.FirstOrDefaultAsync(m => m.AppID == id);

            if (FSApp == null)
            {
                return NotFound();
            }
            return Page();
        }
    }
}
