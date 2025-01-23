using Application_Headstones_Checking_Validation_2025.MVVM.Models;
using Application_Headstones_Checking_Validation_2025.Utilities;
using DevExpress.Mvvm;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Threading.Tasks;

namespace Application_Headstones_Checking_Validation_2025.MVVM.ViewModels
{
    internal class MainViewModel : Abstract.ViewModelBase
    {
        #region Constants
        const string EXCEL_EXTENSION = "*.xlsx";
        #endregion

        /// <summary>
        /// Bindable properties for the view
        /// </summary>
        #region Properties
        private string _OldOutputTextFilePath;

        public string OldOutputTextFilePath
        {
            get { return _OldOutputTextFilePath; }
            set { _OldOutputTextFilePath = value; OnPropertyChanged(); }
        }

        private string _NewOutputTextFilePath;

        public string NewOutputTextFilePath
        {
            get { return _NewOutputTextFilePath; }
            set { _NewOutputTextFilePath = value; OnPropertyChanged(); }
        }
        #endregion

        /// <summary>
        /// Commands for the view
        /// </summary>
        #region Commands
        public DelegateCommand SelectOldOutputCommand { get; private set; }
        public DelegateCommand SelectNewOutputCommand { get; private set; }
        public AsyncCommand CompareChangesCommand { get; private set; }
        #endregion

        #region Fields
        private readonly ExcelHelper _excelHelper;
        #endregion

        public MainViewModel()
        {
            SelectOldOutputCommand = new DelegateCommand(SelectOldOutputExecute);
            SelectNewOutputCommand = new DelegateCommand(SelectNewOutputExecute);
            CompareChangesCommand = new AsyncCommand(CompareChangesExecuteAsync);

            _excelHelper = new ExcelHelper();
        }

        private async Task CompareChangesExecuteAsync()
        {
            try
            {
                if (!HasOutputTextFilePaths()) return;
                int resultSheetIndex = 1;
                int reportSheetIndex = 2;

                List<ExcelComparisonStatusModel> resultChanges = new List<ExcelComparisonStatusModel>();
                List<ExcelComparisonStatusReportModel> statusReport = new List<ExcelComparisonStatusReportModel>();
                // Get the data from the excel files
                IEnumerable<ExcelDataModel> oldOutputData = await _excelHelper.GetIEnumerableExcelData<ExcelDataModel>(OldOutputTextFilePath);
                IEnumerable<ExcelDataModel> newOutputData = await _excelHelper.GetIEnumerableExcelData<ExcelDataModel>(NewOutputTextFilePath);

                int findRowAtOldOutputDataIndex = 0;

                // Do the comparison here
                foreach (ExcelDataModel newData in newOutputData)
                {
                    // find the data from old output data by UID
                    ExcelDataModel oldDataItem = oldOutputData.ElementAtOrDefault(findRowAtOldOutputDataIndex);
                    findRowAtOldOutputDataIndex++;
                    if (oldDataItem == null) continue;

                    foreach (PropertyInfo property in newData.GetType().GetProperties())
                    {
                        if (property == null) continue;
                        if (property.Name.Equals("UID", StringComparison.OrdinalIgnoreCase)) continue;


                        bool hasError = false;
                        string errorName = string.Empty;

                        // Get property name as FieldName
                        // Get the Value of the property
                        string fieldName = property.Name.Trim();
                        string newValue = property.GetValue(newData).ToString().Replace("\u200b", " ").Trim();

                        // Find the old property from old data item by property name
                        PropertyInfo oldProperty = oldDataItem.GetType()
                                                              .GetProperties()
                                                              .FirstOrDefault(e => e.Name.Equals(fieldName, StringComparison.OrdinalIgnoreCase));


                        if (oldProperty == null) return;
                        string oldValue = oldProperty.GetValue(oldDataItem).ToString().Replace("\u200b", " ").Trim();

                        // Check first the resultChanges if the fieldName does exist
                        // Get the item
                        ExcelComparisonStatusModel existingStatus = resultChanges.FirstOrDefault(e => e.Fields == fieldName);

                        // If the item does not exist
                        // Add to resultChanges
                        if (existingStatus == null)
                        {
                            existingStatus = new ExcelComparisonStatusModel { Fields = fieldName };
                            resultChanges.Add(existingStatus);
                        }

                        // Check for changes
                        // Deleted: "When the value has changed to BLANK/EMPTY
                        // Uncoded: "When the BLANK/EMPTY value has changed
                        // Miscoded: "When the value has changed"

                        // Deleted
                        if (!string.IsNullOrWhiteSpace(oldValue) && string.IsNullOrWhiteSpace(newValue))
                        {
                            existingStatus.Deleted++;
                            hasError = true;
                            errorName = "Deleted";
                            //Debug.WriteLine($"{newData.Image_ID} {fieldName} {oldValue} {newValue} deleted");
                        }
                        // Uncoded
                        else if (string.IsNullOrWhiteSpace(oldValue) && !string.IsNullOrWhiteSpace(newValue))
                        {
                            existingStatus.Uncoded++;
                            hasError = true;
                            errorName = "Uncoded";
                            //Debug.WriteLine($"{newData.Image_ID} {fieldName} {oldValue} {newValue} uncoded");
                        }
                        // Miscoded
                        else if (!oldValue.Equals(newValue))
                        {
                            existingStatus.Miscoded++;
                            hasError = true;
                            errorName = "Miscoded";
                            //Debug.WriteLine($"{newData.Image_ID} {fieldName} {oldValue} {newValue} miscoded");
                        }

                        if (hasError)
                        {
                            statusReport.Add(new ExcelComparisonStatusReportModel
                            {
                                ImageNumber = newData.Image_ID,
                                Fields = fieldName,
                                Coded = oldValue,
                                Correction = newValue,
                                TypeError = errorName
                            });
                        }

                        existingStatus.TotalErrors = existingStatus.TotalErrorsCount();
                    }
                }

                // Tally
                int TotalMiscoded = 0;
                int TotalUncoded = 0;
                int TotalDeleted = 0;
                int TotalErrors = 0;

                TotalMiscoded = resultChanges.Select(e => e.Miscoded).Sum();
                TotalUncoded = resultChanges.Select(e => e.Uncoded).Sum();
                TotalDeleted = resultChanges.Select(e => e.Deleted).Sum();
                TotalErrors = resultChanges.Select(e => e.TotalErrors).Sum();

                resultChanges.Add(new ExcelComparisonStatusModel
                {
                    Fields = "Total Errors:",
                    Deleted = TotalDeleted,
                    Uncoded = TotalUncoded,
                    TotalErrors = TotalErrors,
                    Miscoded = TotalMiscoded,
                });

                // Ensure proper sorting and preserve original order
                List<ExcelComparisonStatusReportModel> sortedStatusReport = statusReport
                    .Select((model, index) => new { Model = model, Index = index })
                    .GroupBy(x => x.Model.Fields)
                    .OrderBy(g => g.Max(x => x.Index))
                    .SelectMany(g => g)
                    .Select(x => x.Model)
                    .ToList();

                #region Process Excel Sheet

                #region Process Errors
                // Add the resultChanges to a new sheet in new output file excel
                await _excelHelper.AddDataAtSheetIndex(NewOutputTextFilePath, resultSheetIndex, resultChanges);
                // activate sheet at index
                await _excelHelper.ActivateSheetAtIndex(NewOutputTextFilePath, resultSheetIndex);
                // modify font to bold
                // header and lastrow
                // at sheet 2
                await _excelHelper.ModifyFontToBoldAtFullRow(NewOutputTextFilePath, resultSheetIndex, 1, resultChanges.Count + 1);
                // modify text alignment to right
                // last row and col 1 at sheet 2
                await _excelHelper.ModifyTextHAlignmentAtCell(NewOutputTextFilePath, resultChanges.Count + 1, 1, Syncfusion.XlsIO.ExcelHAlign.HAlignRight, resultSheetIndex);
                #endregion

                #region Process Status Report
                await _excelHelper.AddDataAtSheetIndex(NewOutputTextFilePath, reportSheetIndex, sortedStatusReport);
                await _excelHelper.ActivateSheetAtIndex(NewOutputTextFilePath, reportSheetIndex);
                await _excelHelper.ModifyFontToBoldAtFullRow(NewOutputTextFilePath, reportSheetIndex, 1);
                #endregion

                #endregion

                // Prompt a message
                await InformationMessage("Done", "Successful");

                // open the file
                await OpenFile(NewOutputTextFilePath);

                NewOutputTextFilePath = string.Empty;
            }
            catch (Exception ex)
            {
                ErrorMessage(ex);
            }
        }

        private void SelectNewOutputExecute()
        {
            try
            {
                string newOutputFilePath = GetFilePath("Excel File (*.xlsx)|*.xlsx", EXCEL_EXTENSION, "Select New Output Excel File");

                if (HasOutputTextFile(newOutputFilePath))
                    NewOutputTextFilePath = newOutputFilePath;

            }
            catch (Exception ex)
            {
                ErrorMessage(ex);
            }
        }

        private void SelectOldOutputExecute()
        {
            try
            {
                string oldOutputFilePath = GetFilePath("Excel File (*.xlsx)|*.xlsx", EXCEL_EXTENSION, "Select Old Output Excel File");

                if (HasOutputTextFile(oldOutputFilePath))
                    OldOutputTextFilePath = oldOutputFilePath;

            }
            catch (Exception ex)
            {
                ErrorMessage(ex);
            }
        }


        #region Flags
        private bool HasOutputTextFilePaths()
        {
            return !string.IsNullOrWhiteSpace(OldOutputTextFilePath) && !string.IsNullOrWhiteSpace(NewOutputTextFilePath);
        }

        private bool HasOutputTextFile(string outputTextFilePath)
        {
            return !string.IsNullOrWhiteSpace(outputTextFilePath);
        }
        #endregion
    }
}
