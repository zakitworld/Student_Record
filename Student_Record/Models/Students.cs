using System.ComponentModel.DataAnnotations.Schema;
using System.ComponentModel.DataAnnotations;

namespace Student_Record.Models
{
    public class Students
    {
        [Key]
        public int Id { get; set; }

        [Required(ErrorMessage = "Student Name is required.")]
        [StringLength(100, ErrorMessage = "Student Name cannot be longer than 100 characters.")]
        [RegularExpression(@"^[A-Z][a-zA-Z\s-.']*$", ErrorMessage = "Name must start with a capital letter, cannot contain numbers, and must be in a valid format.")]
        [Display(Name = "Student Name")]
        public string StudentName { get; set; }

        [Required(ErrorMessage = "Father Name is required.")]
        [StringLength(100, ErrorMessage = "Father Name cannot be longer than 100 characters.")]
        [RegularExpression(@"^[A-Z][a-zA-Z\s-.']*$", ErrorMessage = "Name must start with a capital letter, cannot contain numbers, and must be in a valid format.")]
        [Display(Name = "Father's Name")]
        public string FatherName { get; set; }

        [Required(ErrorMessage = "Contact number is required.")]
        [RegularExpression(@"^0\d{9}$", ErrorMessage = "Invalid contact number. It must start with 0 and contain 10 digits.")]
        [Display(Name = "Father's Contact")]
        public string FatherContact { get; set; }

        [Required(ErrorMessage = "Mother Name is required.")]
        [StringLength(100, ErrorMessage = "Mother Name cannot be longer than 100 characters.")]
        [RegularExpression(@"^[A-Z][a-zA-Z\s-.']*$", ErrorMessage = "Name must start with a capital letter, cannot contain numbers, and must be in a valid format.")]
        [Display(Name = "Mother's Name")]
        public string MotherName { get; set; }

        [Required(ErrorMessage = "Contact number is required.")]
        [RegularExpression(@"^0\d{9}$", ErrorMessage = "Invalid contact number. It must start with 0 and contain 10 digits.")]
        [Display(Name = "Mother's Contact")]
        public string MotherContact { get; set; }

        [Required(ErrorMessage = "Date of Birth is required.")]
        [DataType(DataType.Date)]
        [Display(Name = "Date Of Birth")]
        public DateTime DateOfBirth { get; set; }

        [Required(ErrorMessage = "Gender is required")]
        public string Gender { get; set; }

        [Required(ErrorMessage = "Programmed enrolled is required.")]
        [StringLength(100, ErrorMessage = "Programme name cannot be longer than 100 characters")]
        [Display(Name = "Programmed Enrolled")]
        public string ProgrammeEnrolled { get; set; }
    }
}
