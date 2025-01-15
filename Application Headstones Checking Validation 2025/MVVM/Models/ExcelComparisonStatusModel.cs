namespace Application_Headstones_Checking_Validation_2025.MVVM.Models
{
    internal class ExcelComparisonStatusModel
    {
        public string Fields { get; set; }

        // Count when the Value Changed
        public int Miscoded { get; set; }

        // Count when the empty/blank value has changed
        public int Uncoded { get; set; }

        // Count when a value changed to empty/blank
        public int Deleted { get; set; }

        // Total Errors
        public int TotalErrors { get; set; }

        public ExcelComparisonStatusModel()
        {
            Miscoded = 0;
            Uncoded = 0;
            Deleted = 0;
            TotalErrors = 0;
        }

        public int TotalErrorsCount()
        {
            return Miscoded + Uncoded + Deleted;
        }
    }
}
