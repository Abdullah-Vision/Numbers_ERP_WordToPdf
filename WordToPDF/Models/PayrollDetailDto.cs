
namespace WordToPDF.Models
{
    public class PayrollDetailDto
    {
        public int CompanyId { get; set; }
        public string CompanyName { get; set; }
        public string CPhone { get; set; }
        public byte[] Logo { get; set; }
        public string BranchId { get; set; }
        public string BranchName { get; set; }
        public string OrgId { get; set; }
        public string Department { get; set; }
        public int PayrollId { get; set; }
        public int EmpId { get; set; }
        public string EmpNo { get; set; }
        public string CardId { get; set; }
        public string EmpName { get; set; }
        public decimal CurrentPay { get; set; }
        public string EmpFatherName { get; set; }
        public string Doj { get; set; }
        public int WorkingDays { get; set; }
        public int Presents { get; set; }
        public decimal EarPay { get; set; }
        public decimal RatePerDay { get; set; }
        public string Designation { get; set; }
        public string EmployeeType { get; set; }
        public string Nic { get; set; }
        public string Ntn { get; set; }
        public decimal OtAmount { get; set; }
        public decimal OT_HRS { get; set; }
        public decimal G_PAY { get; set; }
        public decimal EAR_PAY { get; set; }
        public decimal WORK_DAYS { get; set; }
        public decimal LOAN { get; set; }
        public decimal BIKE_LOAN { get; set; }
        public decimal TOTAL_DED { get; set; }
        public decimal Incentive { get; set; }
        public decimal EarnedPay { get; set; }
        public string CURR_DATE { get; set; }
        public decimal DeductionAmount { get; set; }
        public decimal NetSalary { get; set; }
        public decimal NetPayable { get; set; }
        public string AmtInWords { get; set; }
        public string IncName { get; set; }
        public string DedName { get; set; }
        public string TranType { get; set; }
        public string TranName { get; set; }
        public decimal TranAmount { get; set; }
    }
}