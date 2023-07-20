using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace EmployeeOperations2.Models
{
    [Table("EmployeeDetails")]
    public class Employee
    {
        [Key]
        public int Id { get; set; }
        
        [Display(Name ="Photo")]
        public string PhotoPath { get; set; }

        [NotMapped]
        public HttpPostedFileBase PhotoFile { get; set; }

        [Display(Name = "Resume")]
        public string ResumePath { get; set; }

        [NotMapped]
        public HttpPostedFileBase ResumeFile { get; set; }

        [Display(Name = "Certifications")]
        public string CertificatesPath { get; set; }

        [NotMapped]
        public IEnumerable<HttpPostedFileBase> CertificateFiles { get; set; }

        [Required]
        [Display(Name = "First Name")]
        public string FirstName { get; set; }

        [Required]
        [Display(Name = "Middle Name")]
        public string MiddleName { get; set; }

        [Required]
        [Display(Name = "Last Name")]
        public string LastName { get; set; }

        [Required]
        [Display(Name = "Email ID")]
        [DataType(DataType.EmailAddress)]
        [EmailAddress]
        public string EmailId { get; set; }

        [Required(ErrorMessage = "You must provide a phone number")]
        [Display(Name = "Mobile Number")]
        [DataType(DataType.PhoneNumber)]
        [RegularExpression(@"^\(?([0-9]{3})\)?[-. ]?([0-9]{3})[-. ]?([0-9]{4})$", ErrorMessage = "Not a valid phone number")]
        public string MobileNumber { get; set; }

        [Required]
        [Display(Name = "Gender")]
        [Range(1, int.MaxValue, ErrorMessage = "The Gender field is required.")]
        public Gender Gender { get; set; }

        [Required]
        [Display(Name = "Date of Birth")]
        public DateTime DateOfBirth { get; set; }
        
        [Display(Name = "Remarks")]
        public string Remarks { get; set; }

        

        [Display(Name = "Country")]
        public string Country { get; set; }

        [NotMapped]
        public int CountryId { get; set; }

        [Display(Name = "City")]
        public string City { get; set; }
        [NotMapped]
        public int CityId { get; set; }
        public string Hobbies { get; set; }
        [NotMapped]
        public List<string> HobbiesSelected { get; set; }
    }

    public enum Gender
    {
        Male = 1,
        Female = 2
    }
}