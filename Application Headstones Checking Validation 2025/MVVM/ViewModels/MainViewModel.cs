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
                int sheetIndex = 1;

                List<ExcelComparisonStatusModel> resultChanges = new List<ExcelComparisonStatusModel>();
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

                        // Get property name as FieldName
                        // Get the Value of the property
                        string fieldName = property.Name.Trim();
                        string newValue = property.GetValue(newData).ToString().Trim();

                        // Find the old property from old data item by property name
                        PropertyInfo oldProperty = oldDataItem.GetType()
                                                              .GetProperties()
                                                              .FirstOrDefault(e => e.Name.Equals(fieldName, StringComparison.OrdinalIgnoreCase));


                        if (oldProperty == null) return;
                        string oldValue = oldProperty.GetValue(oldDataItem).ToString();

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
                        }
                        // Uncoded
                        else if (string.IsNullOrWhiteSpace(oldValue) && !string.IsNullOrWhiteSpace(newValue))
                        {
                            existingStatus.Uncoded++;
                        }
                        // Miscoded
                        else if (!oldValue.Equals(newValue))
                        {
                            existingStatus.Miscoded++;
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

                #region Process Excel Sheet
                // Add the resultChanges to a new sheet in new output file excel
                await _excelHelper.AddDataAtSheetIndex(NewOutputTextFilePath, sheetIndex, resultChanges);

                // activate sheet at index
                await _excelHelper.ActivateSheetAtIndex(NewOutputTextFilePath, sheetIndex);

                // modify font to bold
                // header and lastrow
                // at sheet 2
                await _excelHelper.ModifyFontToBoldAtFullRow(NewOutputTextFilePath, sheetIndex, 1, resultChanges.Count + 1);

                // modify text alignment to right
                // last row and col 1 at sheet 2
                await _excelHelper.ModifyTextHAlignmentAtCell(NewOutputTextFilePath, resultChanges.Count + 1, 1, Syncfusion.XlsIO.ExcelHAlign.HAlignRight, sheetIndex);
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
