public class EmployeeEducationDto
{
    public int EmployeeId { get; set; }
    public string FullName { get; set; }
    public DateTime? DateOfBirth { get; set; }
    public string Gender { get; set; }
    public string Address { get; set; }
    public string City { get; set; }
    public string StateProvince { get; set; }
    public string ZipPostalCode { get; set; }
    public string Country { get; set; }
    public string PhoneNumber { get; set; }
    public string EmailAddress { get; set; }

    public int? EducationId { get; set; }
    public string Degree { get; set; }
    public string University { get; set; }
    public int? YearGraduated { get; set; }
}