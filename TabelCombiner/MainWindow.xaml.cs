using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.IO;
using Forms = System.Windows.Forms;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Input;
using System.ComponentModel;

namespace TabelCombiner
{
    public partial class MainWindow : Window
    {
        ObservableCollection<FileInfo> fileList;
        bool closeApplicationAfterExcelCompleted = false;

        public MainWindow()
        {
            InitializeComponent();
            ExcelLogic.excelWorker.RunWorkerCompleted += ExcelWorker_RunWorkerCompleted;
            ExcelLogic.excelWorker.ProgressChanged += ExcelWorker_ProgressChanged;
            ListBoxFiles.MouseDown += ListBoxFiles_MouseDown; //Added EventHandler here as an error accours in xaml
            fileList = new ObservableCollection<FileInfo>();
            fileList.CollectionChanged += FileList_CollectionChanged;
            ListBoxFiles.ItemsSource = fileList;
        }

        async void Window_Closing(object sender, CancelEventArgs e)
        {
            if (ExcelLogic.excelWorker.IsBusy)
            {
                e.Cancel = true;
                closeApplicationAfterExcelCompleted = true;
                ExcelLogic.excelWorker.CancelAsync();
            }
        }

        private void ExcelWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            Mouse.OverrideCursor = null;
            pbStatus.Visibility = Visibility.Hidden;
            btnCancel.Visibility = Visibility.Hidden;
            btnZusammenfügen.Visibility = Visibility.Visible;

            if(closeApplicationAfterExcelCompleted)
            {
                this.Close();
            }
        }

        private void ExcelWorker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            pbStatus.Value = e.ProgressPercentage;
        }

        
        private void EnableDisableBtnZusammenfügen()
        {
            if (fileList.Count > 0 && (cbSaveTextFile.IsChecked == true || cbShowExcelTabel.IsChecked == true))
            {
                if (btnZusammenfügen.IsEnabled == false)
                {
                    btnZusammenfügen.IsEnabled = true;
                }
            }
            else if (btnZusammenfügen.IsEnabled == true)
            {
                btnZusammenfügen.IsEnabled = false;
            }
        }
        
        private void FileList_CollectionChanged(object sender, System.Collections.Specialized.NotifyCollectionChangedEventArgs e)
        {
            EnableDisableBtnZusammenfügen();
        }

        private void cb_Click(object sender, RoutedEventArgs e)
        {
            EnableDisableBtnZusammenfügen();
        }

        private void BtnHinzufügen_Click(object sender, RoutedEventArgs e)
        {
            Forms.OpenFileDialog openFileDialog = new Forms.OpenFileDialog();
            openFileDialog.Filter = "*.xlsx;*.xlsm|*.xlsx;*.xlsm";
            openFileDialog.Multiselect = true;
            
            try
            {
                if (openFileDialog.ShowDialog() == Forms.DialogResult.OK)
                {
                    foreach (string fileName in openFileDialog.FileNames)
                    {
                        FileInfo newFile = new FileInfo(fileName);
                        bool fileAlreadyExisting = false;
                        foreach (FileInfo file in fileList)
                        {
                            if (file.FullName == newFile.FullName)
                            {
                                fileAlreadyExisting = true;
                                break;
                            }
                        }

                        if (!fileAlreadyExisting)
                        {
                            fileList.Add(newFile);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Log.ErrorMessage(ex.Message);
            }
        }

        private void BtnLöschen_Click(object sender, RoutedEventArgs e)
        {
            FileInfo[] selectedFiles = new FileInfo[ListBoxFiles.SelectedItems.Count];
            ListBoxFiles.SelectedItems.CopyTo(selectedFiles, 0);
            foreach(FileInfo fileInfo in selectedFiles)
            {
                fileList.Remove(fileInfo);
            }
        }

        private void BtnZusammenfügen_Click(object sender, RoutedEventArgs e)
        {
            if(!ExcelLogic.excelWorker.IsBusy)
            {
                Mouse.OverrideCursor = Cursors.Wait;
                btnZusammenfügen.Visibility = Visibility.Hidden;
                pbStatus.Visibility = Visibility.Visible;
                btnCancel.Visibility = Visibility.Visible;

                pbStatus.Value = 0;
                pbStatus.Maximum = fileList.Count;
                ExcelLogic.excelWorker.RunWorkerAsync(new ExcelWorkerArgs(fileList.ToArray(), cbSaveTextFile.IsChecked, cbShowExcelTabel.IsChecked));
            }
        }

        private void btnCancel_Click(object sender, RoutedEventArgs e)
        {
            if(ExcelLogic.excelWorker.IsBusy)
            {
                ExcelLogic.excelWorker.CancelAsync();
            }
        }

        private void ListBoxFiles_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            ListBox listBox = sender as ListBox;

            if(listBox.SelectedItems.Count == 0 && btnLöschen.IsEnabled == true)
            {
                btnLöschen.IsEnabled = false;
            }
            else if (btnLöschen.IsEnabled == false)
            {
                btnLöschen.IsEnabled = true;
            }
        }

        private void ListBoxFiles_MouseDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            HitTestResult hitTestResult = VisualTreeHelper.HitTest(this, e.GetPosition(this));

            if(hitTestResult.VisualHit is not ListBoxItem)
            {
                ListBoxFiles.UnselectAll();
            }
        }
    }
}
