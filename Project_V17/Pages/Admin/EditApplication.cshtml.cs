using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.AspNetCore.Mvc.Rendering;
using Microsoft.EntityFrameworkCore;
using Project_V17.Data;
using Project_V17.Models;

namespace Project_V17.Pages
{
    public class EditApplicationModel : PageModel
    {
        private readonly Project_V17.Data.ApplicationDbContext _context;

        public EditApplicationModel(Project_V17.Data.ApplicationDbContext context)
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

        public async Task<IActionResult> OnPostAsync()
        {
            if (!ModelState.IsValid)
            {
                return Page();
            }

            _context.Attach(FSApp).State = EntityState.Modified;

            try
            {
                await _context.SaveChangesAsync();
            }
            catch (DbUpdateConcurrencyException)
            {
                if (!FSAppExists(FSApp.AppID))
                {
                    return NotFound();
                }
                else
                {
                    throw;
                }
            }

            return RedirectToPage("./Index");
        }

        private bool FSAppExists(int id)
        {
            return _context.FSApp.Any(e => e.AppID == id);
        }
    }
}
