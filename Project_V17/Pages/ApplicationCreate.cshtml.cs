using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.AspNetCore.Mvc.Rendering;
using Project_V17.Data;
using Project_V17.Models;

namespace Project_V17.Pages
{
    public class ApplicationCreateModel : PageModel
    {
        private readonly Project_V17.Data.ApplicationDbContext _context;

        public ApplicationCreateModel(Project_V17.Data.ApplicationDbContext context)
        {
            _context = context;
        }

        public IActionResult OnGet()
        {
            return Page();
        }

        [BindProperty]
        public FSApp FSApp { get; set; }

        public async Task<IActionResult> OnPostAsync()
        {
            if (!ModelState.IsValid)
            {
                return Page();
            }

            _context.FSApp.Add(FSApp);
            await _context.SaveChangesAsync();

            return RedirectToPage("Confirmation");
        }
    }
}