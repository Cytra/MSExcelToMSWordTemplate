using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Input;

namespace ZivileWpfApp.ViewModels
{
    public class ViewModelBase
    {

        public ViewModelBase()
        {

        }

        public void SimpleMethod()
        {
            Debug.WriteLine("Hello");
        }
    }
}
