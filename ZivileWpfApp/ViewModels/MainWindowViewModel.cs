using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Input;
using Microsoft.Office.Interop.Word;
using System.Collections.ObjectModel;
using System.IO;
using System.ComponentModel;
using System.Data;
using System.Windows;
using DataTable = System.Data.DataTable;
using Task = System.Threading.Tasks.Task;

namespace ZivileWpfApp.ViewModels
{
    public class MainWindowViewModel : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;

        public ICommand GenerateWordDocument { get; set; }

        public ICommand FindExcelFile { get; set; }

        public ICommand WriteWordDocLocation { get; set; }
        public ICommand ReadExcelFile { get; set; }

        private string _excelFileLocation;

        public string ExcelFileLocation
        {
            get { return _excelFileLocation; }
            set
            {
                _excelFileLocation = value;
                OnPropertyChanged(nameof(ExcelFileLocation));
            }
        }

        private string _appStatus;

        public string AppStatus
        {
            get { return _appStatus; }
            set
            {
                _appStatus = value;
                OnPropertyChanged(nameof(AppStatus));
            }
        }


        private string _wordFileLocation;

        public string WordFileLocation
        {
            get { return _wordFileLocation; }
            set
            {
                _wordFileLocation = value;
                OnPropertyChanged(nameof(WordFileLocation));
            }
        }

        private DataTable _excelData;

        public DataTable ExcelData
        {
            get { return _excelData; }
            set
            {
                _excelData = value;
                OnPropertyChanged(nameof(ExcelData));
            }
        }

        public MainWindowViewModel()
        {
            GenerateWordDocument = new DelegateCommand(ExecuteGenerateWordDocument, CanExecuteGenerateWordDocument);
            WriteWordDocLocation = new DelegateCommand(ExecuteWriteWordDocLocation);
            //FindExcelFile = new RelayCommand(ExceuteFindExcelFile, CanExceuteFindExcelFile);
            FindExcelFile = new DelegateCommand(ExceuteFindExcelFile);
            ReadExcelFile = new DelegateCommand(ExecuteReadExcelFile, CanExecuteReadExcelFile);
        }

        private bool CanExceuteFindExcelFile()
        {
            return true;
        }

        private bool CanExecuteReadExcelFile()
        {
            if (string.IsNullOrEmpty(ExcelFileLocation))
            {
                return false;
            }
            return true;
        }

        private async void ExecuteReadExcelFile()
        {
            await Task.Run(() =>
             {
                 ExcelData = MSExcelFuncions.ReadExcelDOcument(ExcelFileLocation);
             });

        }

        private void ExecuteWriteWordDocLocation()
        {
            WordFileLocation = FileNavigation.OpenFolderDialog();
        }

        private void ExceuteFindExcelFile()
        {
            ExcelFileLocation = FileNavigation.OpenPathDialog();
        }

        private bool CanExecuteGenerateWordDocument()
        {
            if (string.IsNullOrEmpty(ExcelFileLocation))
            {
                return false;
            }
            if (string.IsNullOrEmpty(WordFileLocation))
            {
                return false;
            }
            return true;
        }

        private async Task GenerateDocuments()
        {
            AppStatus = "In Progress";
            foreach (DataRow row in ExcelData.Rows)
            {
                if (row[0].ToString().Length > 0)
                {
                    await Task.Factory.StartNew(() => MSWordFuncions.GenerateWordFile(row[0].ToString(), row[1].ToString(),
                        row[2].ToString(), row[3].ToString(), row[4].ToString(), row[5].ToString(), row[6].ToString()
                        , row[7].ToString(), row[8].ToString(), row[9].ToString(), row[10].ToString(), WordFileLocation));
                }
            }

            AppStatus = "Done";
        }

        private async void ExecuteGenerateWordDocument()
        {
            //void MyMethod()
            //{
            //    // Do synchronous work.
            //    Thread.Sleep(1000);
            //}
            //async Task MyMethodAsync()
            //{
            //    // Do asynchronous work.
            //    await Task.Delay(1000);
            //}

            //var row = ExcelData.Rows[0];
            //MSWordFuncions.GenerateWordFile(row[0].ToString(), row[1].ToString(), row[2].ToString(), "","38","","","","","","");

            await GenerateDocuments();
        }

        protected void OnPropertyChanged(string name)
        {
            PropertyChangedEventHandler handler = PropertyChanged;
            if (handler != null)
            {
                handler(this, new PropertyChangedEventArgs(name));
            }
        }
    }
}
