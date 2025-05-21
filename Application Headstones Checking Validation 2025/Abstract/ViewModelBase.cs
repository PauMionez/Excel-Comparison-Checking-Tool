using Application_Headstones_Checking_Validation_2025.Utilities;
using System;
using System.Threading.Tasks;

namespace Application_Headstones_Checking_Validation_2025.Abstract
{
    /// <summary>
    /// Base class for all view models
    /// NotifyPropertyChanged is inherited here
    /// NotifyPropertyChanged has DialogFunctions inherited
    /// </summary>
    internal abstract class ViewModelBase : NotifyPropertyChanged
    {
        private readonly DirectoryHelper _directoryHelper;

        /// <summary>
        /// Set the application name
        /// </summary>
        public string ApplicationName { get { return "Application Headstones Checking Validation 2025"; } }

        /// <summary>
        /// Set the application version here.
        /// </summary>
        public string Title
        {
            get
            {
                return $"{ApplicationName} v1.4";
            }
        }

        protected ViewModelBase()
        {
            _directoryHelper = new DirectoryHelper();
        }


        protected string GetFilePath(string displayName, string extensionList, string title)
        {
            try
            {
                return _directoryHelper.GetFilePath(displayName, extensionList, title);
            }
            catch (Exception ex)
            {
                ErrorMessage(ex);
                return string.Empty;
            }
        }

        protected async Task OpenFile(string filePath)
        {
            try
            {
                await _directoryHelper.OpenFile(filePath);
            }
            catch (Exception ex)
            {
                ErrorMessage(ex);
            }
        }

    }
}
