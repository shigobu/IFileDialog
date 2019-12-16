using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using COMInterfaceWrapper;

namespace test
{
    /// <summary>
    /// MainWindow.xaml の相互作用ロジック
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            FolderSelectDialog selectDialog = new FolderSelectDialog();
            selectDialog.Path = "::{20D04FE0-3AEA-1069-A2D8-08002B30309D}";
            if (selectDialog.ShowDialog(IntPtr.Zero))
            {
                button1.Content = selectDialog.Path;
            }
            
        }
    }
}
