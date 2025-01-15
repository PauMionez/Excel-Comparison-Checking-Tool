using Application_Headstones_Checking_Validation_2025.Abstract;
using Microsoft.WindowsAPICodePack.Dialogs;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace Application_Headstones_Checking_Validation_2025.Utilities
{
    internal class DirectoryHelper : DialogFunctions
    {

        public async Task OpenFile(string filePath)
        {
            try
            {
                await Task.Run(() =>
                {
                    Process.Start(new ProcessStartInfo
                    {
                        UseShellExecute = true,
                        FileName = filePath,
                    });
                });
            }
            catch (Exception ex)
            {
                ErrorMessage(ex);
            }
        }

        /// <summary>
        /// Open the file chooser and return the selected files
        /// </summary>
        /// <returns></returns>
        public List<string> GetFilePaths(string DisplayName, string ExtensionList, string Title)
        {
            List<string> pathList = null;

            try
            {
                CommonOpenFileDialog dialog = new CommonOpenFileDialog
                {
                    Title = Title,
                    Multiselect = true
                };

                dialog.Filters.Add(new CommonFileDialogFilter(DisplayName, ExtensionList));

                if (dialog.ShowDialog() != CommonFileDialogResult.Ok)
                {
                    return null;
                }

                pathList = dialog.FileNames.ToList();
            }
            catch (Exception ex)
            {
                ErrorMessage(ex);
            }

            return pathList;
        }
        /// <summary>
        /// Open the file chooser and return the selected file or folder
        /// </summary>
        /// <returns></returns>
        public string GetFilePath(string DisplayName, string ExtensionList, string Title)
        {
            string path = null;

            try
            {
                CommonOpenFileDialog dialog = new CommonOpenFileDialog
                {
                    Title = Title,
                };

                dialog.Filters.Add(new CommonFileDialogFilter(DisplayName, ExtensionList));

                if (dialog.ShowDialog() != CommonFileDialogResult.Ok)
                {
                    return null;
                }

                path = dialog.FileName;
            }
            catch (Exception ex)
            {
                ErrorMessage(ex);
            }

            return path;
        }
        /// <summary>
        /// Open the folder chooser and return the selected paths
        /// </summary>
        /// <param name="IsMultiSelect">Default is false. True = Select 2 or more files</param>
        /// <returns></returns>
        public IEnumerable<string> GetFolderPaths(string Title)
        {
            IEnumerable<string> path = null;

            try
            {
                CommonOpenFileDialog folderSelectorDialog = new CommonOpenFileDialog
                {
                    IsFolderPicker = true,
                    AllowNonFileSystemItems = false,
                    Multiselect = true,
                    Title = Title
                };

                if (folderSelectorDialog.ShowDialog() == CommonFileDialogResult.Ok)
                {
                    path = folderSelectorDialog.FileNames;
                }
            }
            catch (Exception ex)
            {
                ErrorMessage(ex);
            }

            return path;
        }

        /// <summary>
        /// Open the folder chooser and return the path
        /// </summary>
        /// <param name="IsMultiSelect">Default is false. True = Select 2 or more files</param>
        /// <returns></returns>
        public string GetFolderPath(string Title, bool isMultiSelect = false)
        {
            string path = null;

            try
            {
                CommonOpenFileDialog folderSelectorDialog = new CommonOpenFileDialog
                {
                    IsFolderPicker = true,
                    AllowNonFileSystemItems = false,
                    Multiselect = isMultiSelect,
                    Title = Title
                };

                if (folderSelectorDialog.ShowDialog() == CommonFileDialogResult.Ok)
                {
                    path = folderSelectorDialog.FileName;
                }
            }
            catch (Exception ex)
            {
                ErrorMessage(ex);
            }

            return path;
        }
        /// <summary>
        /// Get the column
        /// </summary>
        /// <param name="columnNumber"></param>
        /// <returns></returns>
        public string GetExcelColumnName(int columnNumber)
        {
            int dividend = columnNumber;
            string columnName = string.Empty;
            int modulo;

            while (dividend > 0)
            {
                modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modulo).ToString() + columnName;
                dividend = (dividend - modulo) / 26;
            }

            return columnName;
        }


        /// <summary>
        /// Validate if the filepath is being used
        /// </summary>
        /// <param name="filePath"></param>
        /// <returns></returns>
        public bool IsFileBeingUsed(string filePath)
        {
            bool result = false;

            try
            {
                // Attempt to open the file with specific FileShare options
                using (FileStream fileStream = File.Open(filePath, FileMode.Open, FileAccess.Read, FileShare.None))
                {
                    result = false;
                }
            }
            catch (IOException)
            {
                result = true;
            }

            return result;
        }
    }
}
