using Microsoft.AspNetCore.Mvc;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace ExcelColumnMenu.Controllers
{
    [Route("")]
    [Route("[controller]/[action]")]
    public class HomeController : Controller
    {
        public IActionResult Index()
        {
            var options = new ExcelColumnDropdownOptions();
            options.AddDefaultOptions();
            return View(options);
        }
    }
}
