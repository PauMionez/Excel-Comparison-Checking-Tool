using System;
using System.Diagnostics;
using System.Threading.Tasks;

namespace Application_Headstones_Checking_Validation_2025.Abstract
{
    internal abstract class DialogFunctions
    {
        /// <summary>
        /// Display an error message
        /// </summary>
        /// <param name="ex">Your exception</param>
        /// <param name="message">Title text. Typically, put here your method name</param>
        public static async Task ErrorMessage(Exception ex)
        {
            try
            {
                StackTrace stackTrace = new StackTrace(ex);
                System.Reflection.MethodBase method = stackTrace.GetFrame(stackTrace.FrameCount - 1).GetMethod();
                string titleText = method.Name;

                //MessageBox.Show(string.Format(ex.Message + "\n\n" + ex.StackTrace + "\n\n{0}", "Please screenshot and send procedures on how this error occured. Thank you."), titleText, MessageBoxButton.OK, MessageBoxImage.Error);
                Wpf.Ui.Controls.MessageBox uiMessageBox = new Wpf.Ui.Controls.MessageBox
                {
                    Title = titleText,
                    Content =
              string.Format(ex.Message + "\n\n" + ex.StackTrace + "\n\n{0}", "Please screenshot and send procedures on how this error occured. Thank you."),
                };

                await uiMessageBox.ShowDialogAsync();
            }
            catch (Exception ex2)
            {
                await ErrorMessage(ex2);
            }
        }

        /// <summary>
        /// Displays an information messagebox
        /// </summary>
        /// <param name="message">Text</param>
        /// <param name="title">Title text</param>
        public static async Task InformationMessage(string message, string title)
        {
            try
            {
                Wpf.Ui.Controls.MessageBox uiMessageBox = new Wpf.Ui.Controls.MessageBox
                {
                    Title = title,
                    Content = message,
                };

                await uiMessageBox.ShowDialogAsync();
            }
            catch (Exception ex2)
            {
                await ErrorMessage(ex2);
            }
        }

        /// <summary>
        /// Displays a warning messagebox
        /// </summary>
        /// <param name="message">Text</param>
        /// <param name="title">Title text</param>
        public static async Task WarningMessage(string message, string title)
        {
            try
            {
                //MessageBox.Show(message, title, MessageBoxButton.OK, MessageBoxImage.Warning);
                Wpf.Ui.Controls.MessageBox uiMessageBox = new Wpf.Ui.Controls.MessageBox
                {
                    Title = title,
                    Content = message,
                };

                await uiMessageBox.ShowDialogAsync();
            }
            catch (Exception ex2)
            {
                await ErrorMessage(ex2);
            }
        }
        /// <summary>
        /// Displays a warning messagebox
        /// </summary>
        /// <param name="message">Text</param>
        /// <param name="title">Title text</param>
        public static async Task WarningMessage(string message)
        {
            try
            {
                await WarningMessage(message, "Warning");
            }
            catch (Exception ex2)
            {
                await ErrorMessage(ex2);
            }
        }

        /// <summary>
        /// Displays a question messagebox
        /// </summary>
        /// <param name="message"></param>
        /// <param name="title"></param>
        /// <returns></returns>
        //public static MessageBoxResult YesNoDialog(string message, string title)
        //{
        //    try
        //    {
        //        return MessageBox.Show(message, title, MessageBoxButton.YesNo, MessageBoxImage.Question);
        //    }
        //    catch (Exception ex2)
        //    {
        //        ErrorMessage(ex2);
        //        return MessageBoxResult.No;
        //    }
        //}
    }
}
