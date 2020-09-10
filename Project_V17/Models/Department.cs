using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Threading.Tasks;

namespace Project_V17.Models
{
    public class Department
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        public int DepID { get; set; }

        [Required]
        [Display(Name = "Department Name")]
        [RegularExpression(@"\w{2,50}", ErrorMessage = "The Department name must be more than two letters")]
        public string DepartmentName { get; set; } = "";

    }
}
