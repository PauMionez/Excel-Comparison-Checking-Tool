using Application_Headstones_Checking_Validation_2025.Abstract;
using Syncfusion.XlsIO;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace Application_Headstones_Checking_Validation_2025.Utilities
{
    internal class ExcelHelper : DialogFunctions
    {
        public async Task<IEnumerable<string>> GetColumnHeadersAsync(string excelFilePath, string[] excludeColumns = null)
        {

            try
            {
                if (string.IsNullOrWhiteSpace(excelFilePath)) return null;

                List<string> columnHeaders = new List<string>();

                await Task.Run(() =>
                {
                    using (ExcelEngine excelEngine = new ExcelEngine())
                    {
                        IApplication application = excelEngine.Excel;
                        application.DefaultVersion = ExcelVersion.Xlsx;
                        IWorkbook workbook = application.Workbooks.Open(excelFilePath);
                        IWorksheet worksheet = workbook.Worksheets[0];

                        for (int i = 1; i <= worksheet.UsedRange.LastColumn; i++)
                        {
                            string header = worksheet[1, i].Text.Trim();

                            if (string.IsNullOrWhiteSpace(header)) continue;

                            if (excludeColumns != null && Array.Exists(excludeColumns, element => element.Equals(header, StringComparison.OrdinalIgnoreCase)))
                                continue;

                            columnHeaders.Add(header);
                        }

                    }
                });

                return columnHeaders;
            }
            catch (Exception ex)
            {
                ErrorMessage(ex);
                return null;
            }
        }

        public async Task<IEnumerable<T>> GetIEnumerableExcelData<T>(string excelFilePath) where T : new()
        {

            try
            {
                if (string.IsNullOrWhiteSpace(excelFilePath)) return null;

                List<T> resultList = new List<T>();

                await Task.Run(() =>
                {
                    using (ExcelEngine excelEngine = new ExcelEngine())
                    {
                        IApplication application = excelEngine.Excel;
                        application.DefaultVersion = ExcelVersion.Xlsx;

                        IWorkbook workbook = application.Workbooks.Open(excelFilePath);
                        IWorksheet worksheet = workbook.Worksheets[0];

                        int lastRow = worksheet.UsedRange.LastRow;
                        int lastColumn = worksheet.UsedRange.LastColumn;

                        Dictionary<int, string> headers = new Dictionary<int, string>();

                        // Read headers
                        for (int col = 1; col <= lastColumn; col++)
                        {
                            string header = worksheet[1, col].Text.Trim();
                            if (string.IsNullOrWhiteSpace(header)) continue;

                            headers[col] = header;
                        }

                        // Read data and map to model
                        for (int row = 2; row <= lastRow; row++)
                        {
                            T item = new T();
                            foreach (KeyValuePair<int, string> header in headers)
                            {
                                // case-insensitive matching
                                PropertyInfo property = typeof(T).GetProperty(header.Value,
                                    BindingFlags.IgnoreCase | BindingFlags.Public | BindingFlags.Instance);

                                if (property != null)
                                {
                                    string cellValue = worksheet[row, header.Key].Value;
                                    object safeValue = Convert.ChangeType(cellValue, property.PropertyType);
                                    property.SetValue(item, safeValue);
                                }

                            }

                            resultList.Add(item);
                        }
                    }
                });

                return resultList;
            }
            catch (Exception ex)
            {
                ErrorMessage(ex);
            }

            return null;
        }

        public async Task ModifyFontToBoldAtFullRow(string excelFilePath, int sheetIndex = 0, params int[] rowArgs)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(excelFilePath)) return;
                if (rowArgs == null) return;

                await Task.Run(() =>
                {
                    using (ExcelEngine excelEngine = new ExcelEngine())
                    {
                        IApplication application = excelEngine.Excel;
                        application.DefaultVersion = ExcelVersion.Xlsx;
                        IWorkbook workbook = application.Workbooks.Open(excelFilePath);

                        if (sheetIndex > workbook.Worksheets.Count)
                        {
                            WarningMessage("Sheet doesn't exist.");
                            return;
                        }

                        IWorksheet worksheet = workbook.Worksheets[sheetIndex];

                        int lastColumn = worksheet.UsedRange.LastColumn;

                        foreach (int row in rowArgs)
                        {
                            for (int col = 1; col <= lastColumn; col++)
                            {
                                worksheet.Range[row, col].CellStyle.Font.Bold = true;
                            }
                        }

                        workbook.Save();
                    }
                });
            }
            catch (Exception ex)
            {
                ErrorMessage(ex);
            }
        }

        public async Task ModifyTextHAlignmentAtCell(string excelFilePath, int row, int column, ExcelHAlign excelHAlign, int sheetIndex = 0)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(excelFilePath)) return;

                await Task.Run(() =>
                {
                    using (ExcelEngine excelEngine = new ExcelEngine())
                    {
                        IApplication application = excelEngine.Excel;
                        application.DefaultVersion = ExcelVersion.Xlsx;
                        IWorkbook workbook = application.Workbooks.Open(excelFilePath);

                        if (sheetIndex > workbook.Worksheets.Count)
                        {
                            WarningMessage("Sheet doesn't exist.");
                            return;
                        }

                        IWorksheet worksheet = workbook.Worksheets[sheetIndex];

                        worksheet.Range[row, column].CellStyle.HorizontalAlignment = excelHAlign;

                        workbook.Save();
                    }
                });
            }
            catch (Exception ex)
            {
                ErrorMessage(ex);
            }
        }

        public async Task RemoveFormatting(string excelFilePath, int sheetIndex = 0)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(excelFilePath)) return;

                await Task.Run(() =>
                {
                    using (ExcelEngine excelEngine = new ExcelEngine())
                    {
                        IApplication application = excelEngine.Excel;
                        application.DefaultVersion = ExcelVersion.Xlsx;
                        IWorkbook workbook = application.Workbooks.Open(excelFilePath);

                        if (sheetIndex > workbook.Worksheets.Count)
                        {
                            WarningMessage("Sheet doesn't exist.");
                            return;
                        }

                        IWorksheet worksheet = workbook.Worksheets[sheetIndex];

                        int lastRow = worksheet.UsedRange.LastRow;
                        int lastColumn = worksheet.UsedRange.LastColumn;

                        worksheet.Range[1, 1, lastRow + 1, lastColumn + 1].Clear(ExcelClearOptions.ClearFormat);


                        workbook.Save();
                    };
                });

            }
            catch (Exception ex)
            {
                ErrorMessage(ex);
            }
        }

        public async Task AddDataAtSheetIndex<T>(string excelFilePath, int sheetIndex, IEnumerable<T> data, bool makeHeadersToBold = true) where T : new()
        {
            try
            {
                if (string.IsNullOrWhiteSpace(excelFilePath)) return;

                await Task.Run(() =>
                {
                    using (ExcelEngine excelEngine = new ExcelEngine())
                    {
                        IApplication application = excelEngine.Excel;
                        application.DefaultVersion = ExcelVersion.Xlsx;
                        IWorkbook workbook = application.Workbooks.Open(excelFilePath);

                        if (sheetIndex != application.Worksheets.Count)
                        {
                            workbook.Worksheets.Create();
                        }

                        // Remove unnecessary sheets
                        foreach (IWorksheet ws in workbook.Worksheets.ToList())
                        {
                            if (ws.Index > sheetIndex)
                            {
                                workbook.Worksheets.Remove(ws);
                            }
                        }

                        IWorksheet worksheet = workbook.Worksheets[sheetIndex];

                        worksheet.Clear();

                        #region Set Column Headers
                        T type = new T();
                        Dictionary<int, string> columnHeaders = new Dictionary<int, string>();
                        PropertyInfo[] properties = type.GetType().GetProperties();

                        for (int i = 0; i < properties.Length; i++)
                        {
                            string header = properties[i].Name.ToString().Trim();
                            if (string.IsNullOrWhiteSpace(header)) continue;

                            columnHeaders[i + 1] = header;
                        }
                        #endregion

                        // Add ColumnHeaders Dynamically to worksheet
                        foreach (KeyValuePair<int, string> columnHeader in columnHeaders.OrderBy(e => e.Key))
                        {
                            string addSpaceBeforeUpperCase = Regex.Replace(columnHeader.Value.Trim(), "(?<!^)([A-Z])", " $1");

                            worksheet.Range[1, columnHeader.Key].Text = addSpaceBeforeUpperCase;
                        }

                        // Add Data to worksheet Dynamically
                        int initialRow = 2;

                        foreach (T item in data)
                        {
                            foreach (KeyValuePair<int, string> columnHeader in columnHeaders.OrderBy(e => e.Key))
                            {
                                string columnName = columnHeader.Value.Trim();
                                PropertyInfo[] itemProperties = item.GetType().GetProperties();

                                PropertyInfo propertyItem = itemProperties.FirstOrDefault(e => e.Name.Equals(columnName, StringComparison.OrdinalIgnoreCase));

                                if (propertyItem == null) continue;

                                object rowData = propertyItem.GetValue(item);

                                worksheet.Range[initialRow, columnHeader.Key].Value2 = rowData;
                            }
                            initialRow++;
                        }

                        worksheet.UsedRange.WrapText = false;
                        worksheet.UsedRange.AutofitColumns();
                        worksheet.UsedRange.AutofitRows();

                        workbook.Save();
                    }
                });
            }
            catch (Exception ex)
            {
                ErrorMessage(ex);
            }
        }

        public async Task ActivateSheetAtIndex(string excelFilePath, int sheetIndex)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(excelFilePath)) return;
                await Task.Run(() =>
                {
                    using (ExcelEngine excelEngine = new ExcelEngine())
                    {
                        IApplication application = excelEngine.Excel;
                        application.DefaultVersion = ExcelVersion.Xlsx;

                        IWorkbook workbook = application.Workbooks.Open(excelFilePath);

                        if (sheetIndex > workbook.Worksheets.Count)
                        {
                            WarningMessage("Sheet doesn't exist.");
                            return;
                        }

                        workbook.Worksheets[sheetIndex].Activate();

                        workbook.Save();
                    }
                });
            }
            catch (Exception ex)
            {
                ErrorMessage(ex);
            }
        }
    }
}
