using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace ExcelColumnMenu.Controllers
{
    public class ExcelColumnDropdownOptions
    {
        public string SheetName { get; set; }
        public string CoursePeriodOptions { get; set; }
        public string CourseOptions { get; set; }
        public string CourseInstitutionOptions { get; set; }

        public ExcelColumnDropdownOptions()
        {
            
        }
        public string[] CoursePeriods
        { get
            {
                return this.CoursePeriodOptions?.Split(",").ToArray();
            } 
        }
        public string[] Courses
        {
            get
            {
                return this.CourseOptions?.Split(",").ToArray();
            }
        }
        public string[] CourseInstitution
        {
            get
            {
                return this.CourseInstitutionOptions?.Split(",").ToArray();
            }
        }
        public void AddDefaultOptions()
        {

            this.SheetName = "coursedetails.xlsx";
            this.CoursePeriodOptions = "10 Days,2 Weeks,5 Weeks,2 Months,5 Months,6 Months,12 Months";
            this.CourseOptions = ".NET,NODE.JS,JAVA FULL STACK COURSE,ANGULAR 4/5/6,REACT JS,TYPESCRIPT 3 ,JAVASCRIPT,CSS,HTML,HTML 5 ,SQL,BOOTSTRAP 3,BOOTSRATP 4,ASP.NET CORE MVC,ASP.NET CORE WEB API,PYTHON";
            CourseInstitutionOptions = "ABC Technologies,Mahendar Techno School,XYZ Institute,Naresh Technologies";
        }

       
    }
}
