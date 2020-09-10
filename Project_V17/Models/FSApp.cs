using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Threading.Tasks;

namespace Project_V17.Models
{
    public class FSApp
    {

        #region Staff Details

        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        public int AppID { get; set; }

        [Required]
        [Display(Name = "First Name")]
        [RegularExpression(@"\w{2,50}", ErrorMessage = "The First Name must be more than two letters")]
        public string StaffFirstName { get; set; } = "";


        [Required]
        [Display(Name = " Surname")]
        [RegularExpression(@"[\w’-]{2,}", ErrorMessage = "The Surame must be at least two letters")]
        public string StaffSurname { get; set; } = "";

        [Required]
        [Display(Name = "Department")]

        public string Department { get; set; } = "";

        [Required]
        [Display(Name = "Function")]

        public string Function { get; set; } = "";


        #endregion



        #region Course Details


        [Required]
        [Display(Name = "Course Name")]
        [RegularExpression(@"\w{2,50}", ErrorMessage = "The Course Name must be more than two letters")]
        public string CourseName { get; set; } = "";


        [Required]
        [Display(Name = " Level")]

        public int Level { get; set; }

        [Required]
        [Display(Name = "Provider")]

        public string Provider { get; set; } = "";

        [Required]
        [Display(Name = "Details of Course")]

        public string Details { get; set; } = "";

        [Required]
        [Display(Name = "Start Year")]

        public int StartYear { get; set; }

        [Required]
        [Display(Name = "Mode of Study")]

        public string Mode { get; set; } = "";

        [Required]
        [Display(Name = "Duration of Course (Years)")]

        public int Duration { get; set; }

        [Required]
        [Display(Name = "Cost of Course per year (Round to nearest Euro")]
        [RegularExpression(@"\d*", ErrorMessage = "The entry must be a number")]
        public float Cost { get; set; }




        #endregion


        #region Supplimentary Questions
        [Required(ErrorMessage = "You Must Enter Details")]
        [Display(Name = "Why do you think this programme is of relevance/importance to your work at IT Sligo?")]

        public string Q1 { get; set; } = "";
        [Required(ErrorMessage = "You Must Enter Details")]
        [Display(Name = "Why do you think this programme is of relevance/importance to the mission and objectives of IT Sligo?")]

        public string Q2 { get; set; } = "";

        [Required(ErrorMessage = "You Must Enter Details")]
        [Display(Name = "What specific knowledge, skills and competencies do you expect to gain/ improve as a result of completing this programme of study?")]

        public string Q3 { get; set; } = "";

        [Required(ErrorMessage = "You Must Enter Details")]
        [Display(Name = "How will you use what you have learned from this programme of study to improve your performance and the performance of IT Sligo?")]

        public string Q4 { get; set; } = "";

        [Required(ErrorMessage = "You Must Enter Details")]
        [Display(Name = "Please provide details of other support (financial / non financial) which you would like IT to consider supporting")]

        public string Q5 { get; set; } = "";
        #endregion
    }
}
