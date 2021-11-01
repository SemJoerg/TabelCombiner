using System;
using System.Collections.Generic;
using System.Text;
using System.Windows;

namespace TabelCombiner
{
    public static class Log
    {
        //public static void ErrorMessage(string message)
        //{
        //    MessageBox.Show(message, "Error", MessageBoxButton.OK, MessageBoxImage.Error);
        //}

        public static void ErrorMessage(object message)
        {
            MessageBox.Show(message.ToString(), "Error", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }
}
