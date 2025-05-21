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
        //const string EXCEL_EXTENSION = "*.xlsx";
        private static readonly string[] EXCEL_EXTENSION = { "*.xlsx", "*.xls", "*.xlsm", "*.xltx", "*.xltm", "*.xlsb" };
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

        /// <summary>
        /// Added AdditionalFields to the ExcelDataModel for dynamic/unexpected columns
        /// also added AdditionalFields to exceldatamodel
        /// </summary>
        /// <returns></returns>
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
                        string newValue = property.GetValue(newData)?.ToString().Replace("\u200b", " ").Trim() ?? string.Empty;
                        //string newValue = property.GetValue(newData).ToString().Replace("\u200b", " ").Trim();

                        // Find the old property from old data item by property name
                        PropertyInfo oldProperty = oldDataItem.GetType()
                                                              .GetProperties()
                                                              .FirstOrDefault(e => e.Name.Equals(fieldName, StringComparison.OrdinalIgnoreCase));


                        //if (oldProperty == null) return;
                        if (oldProperty == null) continue;
                        string oldValue = oldProperty.GetValue(oldDataItem)?.ToString().Replace("\u200b", " ").Trim() ?? string.Empty;
                        //string oldValue = oldProperty.GetValue(oldDataItem).ToString().Replace("\u200b", " ").Trim();

                        CompareAndAddStatus(fieldName, oldValue, newValue, GetPossibleImageId(newData), resultChanges, statusReport);

                        #region dump
                        //// Check first the resultChanges if the fieldName does exist
                        //// Get the item
                        //ExcelComparisonStatusModel existingStatus = resultChanges.FirstOrDefault(e => e.Fields == fieldName);

                        //// If the item does not exist
                        //// Add to resultChanges
                        //if (existingStatus == null)
                        //{
                        //    existingStatus = new ExcelComparisonStatusModel { Fields = fieldName };
                        //    resultChanges.Add(existingStatus);
                        //}

                        //// Check for changes
                        //// Deleted: "When the value has changed to BLANK/EMPTY
                        //// Uncoded: "When the BLANK/EMPTY value has changed
                        //// Miscoded: "When the value has changed"


                        //bool skipBecauseOfQuestionMark = (oldValue == "?" ^ newValue == "?");

                        //// Deleted
                        //if (!string.IsNullOrWhiteSpace(oldValue) && string.IsNullOrWhiteSpace(newValue))
                        //{
                        //    existingStatus.Deleted++;
                        //    hasError = true;
                        //    errorName = "Deleted";
                        //    //Debug.WriteLine($"{newData.Image_ID} {fieldName} {oldValue} {newValue} deleted");
                        //}
                        //// Uncoded
                        //else if (string.IsNullOrWhiteSpace(oldValue) && !string.IsNullOrWhiteSpace(newValue))
                        //{
                        //    existingStatus.Uncoded++;
                        //    hasError = true;
                        //    errorName = "Uncoded";
                        //    //Debug.WriteLine($"{newData.Image_ID} {fieldName} {oldValue} {newValue} uncoded");
                        //}
                        //// Miscoded
                        //else if (!oldValue.Equals(newValue))
                        //{
                        //    existingStatus.Miscoded++;
                        //    hasError = true;
                        //    errorName = "Miscoded";
                        //    //Debug.WriteLine($"{newData.Image_ID} {fieldName} {oldValue} {newValue} miscoded");
                        //}

                        //if (hasError)
                        //{
                        //    statusReport.Add(new ExcelComparisonStatusReportModel
                        //    {
                        //        ImageNumber = newData.Image_ID,
                        //        Fields = fieldName,
                        //        Coded = oldValue,
                        //        Correction = newValue,
                        //        TypeError = errorName
                        //    });
                        //}

                        //existingStatus.TotalErrors = existingStatus.TotalErrorsCount();
                        #endregion
                    }

                    //Compare AdditionalFields (dynamically-added columns)
                    foreach (var kvp in newData.AdditionalFields)
                    {
                        string fieldName = kvp.Key;
                        string newValue = kvp.Value?.Replace("\u200b", " ").Trim() ?? string.Empty;
                        string oldValue = oldDataItem.AdditionalFields.ContainsKey(fieldName)
                            ? oldDataItem.AdditionalFields[fieldName]?.Replace("\u200b", " ").Trim() ?? string.Empty
                            : string.Empty;

                        CompareAndAddStatus(fieldName, oldValue, newValue, GetPossibleImageId(newData), resultChanges, statusReport);
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


        /// <summary>
        /// Compare the old and new values of the field Validation
        /// Check for changes
        /// Deleted: "When the value has changed to BLANK/EMPTY
        /// Uncoded: "When the BLANK/EMPTY value has changed
        /// Miscoded: "When the value has changed"
        /// Skip comparison if change is between "?" and a letter (in either direction)
        /// </summary>
        /// <param name="fieldName"></param>
        /// <param name="oldValue"></param>
        /// <param name="newValue"></param>
        /// <param name="imageId"></param>
        /// <param name="resultChanges"></param>
        /// <param name="statusReport"></param>
        private void CompareAndAddStatus(string fieldName, string oldValue, string newValue, string imageId, List<ExcelComparisonStatusModel> resultChanges, List<ExcelComparisonStatusReportModel> statusReport)
        {
            try
            {
                bool hasError = false;
                string errorName = string.Empty;


                // Skip "?"
                if ((oldValue == "?" && !string.IsNullOrWhiteSpace(newValue) && newValue != "?") ||
                    (newValue == "?" && !string.IsNullOrWhiteSpace(oldValue) && oldValue != "?"))
                {
                    return;
                }

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

                // Deleted
                if (!string.IsNullOrWhiteSpace(oldValue) && string.IsNullOrWhiteSpace(newValue))
                {
                    existingStatus.Deleted++;
                    hasError = true;
                    errorName = "Deleted";
                }
                // Uncoded
                else if (string.IsNullOrWhiteSpace(oldValue) && !string.IsNullOrWhiteSpace(newValue))
                {
                    existingStatus.Uncoded++;
                    hasError = true;
                    errorName = "Uncoded";
                }
                // Miscoded
                else if (!oldValue.Equals(newValue, StringComparison.Ordinal))
                {
                    existingStatus.Miscoded++;
                    hasError = true;
                    errorName = "Miscoded";
                }

                if (hasError)
                {
                    statusReport.Add(new ExcelComparisonStatusReportModel
                    {
                        ImageNumber = imageId,
                        Fields = fieldName,
                        Coded = oldValue,
                        Correction = newValue,
                        TypeError = errorName
                    });
                }

                existingStatus.TotalErrors = existingStatus.TotalErrorsCount();
            }
            catch (Exception ex)
            {
                ErrorMessage(ex);
            }
        }

        /// <summary>
        /// For Dynamic Image_ID
        /// </summary>
        /// <param name="data"></param>
        /// <returns></returns>
        private string GetPossibleImageId(ExcelDataModel data)
        {
            string result = string.Empty;
            try
            {
                List<string> imageIdList = new List<string>
                {
                    data.Image_ID,
                    data.Image_Name,
                    data.ImageNumber,
                    data.ImageName
                };

                result = imageIdList.FirstOrDefault(e => !string.IsNullOrWhiteSpace(e)) ?? "Unknown";
            }
            catch (Exception ex)
            {
                ErrorMessage(ex);
            }
            return result;
        }


        private void SelectNewOutputExecute()
        {
            try
            {
                string filterExtensions = string.Join(";", EXCEL_EXTENSION.Select(ext => $"*.{ext}"));
                string newOutputFilePath = GetFilePath("Excel File (*.xlsx)|*.xlsx", filterExtensions, "Select New Output Excel File");

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
                string filterExtensions = string.Join(";", EXCEL_EXTENSION.Select(ext => $"*.{ext}"));
                string oldOutputFilePath = GetFilePath("Excel File (*.xlsx)|*.xlsx", filterExtensions, "Select Old Output Excel File");

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
