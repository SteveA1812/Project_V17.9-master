using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Identity;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace Project_V17.Models
{

    [Authorize(Roles = "Admin")]
    public class AdminController:Controller
    {

        private readonly UserManager<IdentityUser>
          _userManager;

        public AdminController(
            UserManager<IdentityUser> userManager)
        {
            _userManager = userManager;
        }

        public async Task<IActionResult> Index()
        {
            var admins = (await _userManager
                .GetUsersInRoleAsync("Admin"))
                .ToArray();

            var everyone = await _userManager.Users
                .ToArrayAsync();

            var model = new ManageUsersViewModels
            {
                Administrators = admins,
                Everyone = everyone
            };

            return View(model);
        }
    }
}
