using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace NpoiSample.Models
{
    public class PersonRepository
    {
        public List<Person> Get( )
        {
            return new List<Person>
            {
                new Person
                {
                    PersonId = 1,
                    Email = "person1@example.com",
                    FirstName = "PersonOne",
                    LastName = "One"
                },
                new Person
                {
                    PersonId = 2,
                    Email = "person2@example.com",
                    FirstName = "PersonTwo",
                    LastName = "Two"
                },
                new Person
                {
                    PersonId = 3,
                    Email = "person3@example.com",
                    FirstName = "PersonThree",
                    LastName = "Three",
                    Age = 5
                },
                new Person
                {
                    PersonId = 4,
                    Email = null,
                    FirstName = "PersonFour",
                    LastName = "Four",
                    Age = 7
                },
            };
        }
    }
}