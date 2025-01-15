using System;
using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace Application_Headstones_Checking_Validation_2025.Abstract
{
    internal abstract class NotifyPropertyChanged : DialogFunctions, INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;

        public void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            try
            {
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
            }
            catch (Exception ex)
            {
                ErrorMessage(ex);
            }
        }
    }
}
