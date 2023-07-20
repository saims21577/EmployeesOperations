using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace EmployeeOperations2.Models
{
    public class City
    {
        public int Id { get; set; }
        public string Name { get; set; }
        public int countryId { get; set; }
    }
}