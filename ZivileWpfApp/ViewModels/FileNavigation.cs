using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ZivileWpfApp.ViewModels
{
    public static class FileNavigation
    {
        public static string OpenPathDialog()
        {
            var result = "";
            using (var fileDialog = new OpenFileDialog{CheckPathExists = false,CheckFileExists = false})
            {
                fileDialog.ShowDialog();

                result = fileDialog.FileNames.FirstOrDefault();

            }
            return result;
        }

        public static string OpenFolderDialog()
        {
            var result = "";
            using (var fbd = new FolderBrowserDialog())
            {
                DialogResult fialog = fbd.ShowDialog();

                if (fialog == DialogResult.OK && !string.IsNullOrWhiteSpace(fbd.SelectedPath))
                {
                    result = fbd.SelectedPath;
                }
            }

            return result;

        }
    }
}
