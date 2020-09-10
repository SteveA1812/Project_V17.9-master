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
    public class DeleteModel : PageModel
    {
        private readonly Project_V17.Data.ApplicationDbContext _context;

        public DeleteModel(Project_V17.Data.ApplicationDbContext context)
        {
            _context = context;
        }

        [BindProperty]
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

        public async Task<IActionResult> OnPostAsync(int? id)
        {
            if (id == null)
            {
                return NotFound();
            }

            FSApp = await _context.FSApp.FindAsync(id);

            if (FSApp != null)
            {
                _context.FSApp.Remove(FSApp);
                await _context.SaveChangesAsync();
            }

            return RedirectToPage("./Index");
        }
    }
}
