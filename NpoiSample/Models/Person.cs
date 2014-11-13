using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace NpoiSample.Models
{
    public class Person
    {
        public int PersonId { get; set; }

        public string Email { get; set; }

        public string FirstName { get; set; }
        public string LastName { get; set; }

        public int? Age { get; set; }
    }
}