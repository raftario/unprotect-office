using System;
using System.IO;
using System.Collections.Generic;
using System.IO.Compression;
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
using Microsoft.Win32;
using Path = System.IO.Path;

namespace UnprotectOffice
{
    /// <summary>
    /// MainWindow interaction logic
    /// </summary>
    public partial class MainWindow : Window
    {
        private string[] _files;

        public MainWindow()
        {
            InitializeComponent();
        }

        /// <summary>
        /// Open file picker and refresh label and button based on input
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void PickFiles(object sender, RoutedEventArgs e)
        {
            var dialog = new OpenFileDialog
            {
                Multiselect = true,
                InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
                Filter = "OOXML files|*.docx;*.docm;*.pptx;*.pptm;*.xlsx;*.xlsm"
            };

            if (dialog.ShowDialog() == true)
            {
                UnprotectButton.IsEnabled = true;
                FilesText.Text = string.Join(
                    ", ",
                    dialog.FileNames.Select(ShortPath).ToArray()
                );

                _files = dialog.FileNames;
            }
            else
            {
                UnprotectButton.IsEnabled = false;
                FilesText.Text = "File(s)";

                _files = Array.Empty<string>();
            }
        }

        private void OpenUri(object sender, RequestNavigateEventArgs e)
        {
            System.Diagnostics.Process.Start(e.Uri.ToString());
        }

        /// <summary>
        /// Remove write protection from files
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void UnprotectFiles(object sender, RoutedEventArgs e)
        {
            foreach (var file in _files)
            {
                var extractPath = ExtractFile(file);
            }
        }

        /// <summary>
        /// Extract file to temporary directory
        /// </summary>
        /// <param name="file">Path to file</param>
        /// <returns>Path to extraction directory</returns>
        private static string ExtractFile(string file)
        {
            var extractPath = Path.Combine(Path.GetTempPath(), ShortPath(file));

            ZipFile.ExtractToDirectory(file, extractPath);

            return extractPath;
        }

        /// <summary>
        /// Get the file name from a path
        /// </summary>
        /// <param name="file">Path</param>
        /// <returns>File name</returns>
        private static string ShortPath(string file)
        {
            return file.Split(Path.DirectorySeparatorChar).Last();
        }
    }
}