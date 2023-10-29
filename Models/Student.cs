using System.ComponentModel.DataAnnotations;

namespace openxml.Models
{
    public class Student
    {
        public int  Id { get; set; }
        public string FirstName { get; set; } = string.Empty;
        public string LastName { get; set; }= string.Empty;
        public int Age { get; set; }
        public string City { get; set; } = string.Empty;
    }
}
