using System;
using System.Windows;
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
            //selectDialog.Title = "タイトルが設定できます。";
            selectDialog.Path = "::{20D04FE0-3AEA-1069-A2D8-08002B30309D}";//PCを初期画面に
            IntPtr hWnd = new System.Windows.Interop.WindowInteropHelper(this).Handle;
            try
            {
                if (selectDialog.ShowDialog(hWnd))
                {
                    button1.Content = selectDialog.Path;
                }
            }
            catch(Exception ex)
            {
                button1.Content = ex.Message;
            }
            
        }
    }
}
