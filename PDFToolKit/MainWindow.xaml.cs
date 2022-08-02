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
using System.IO;
using Windows.Data.Pdf;

namespace PDFToolKit
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private PdfDocument pdf;
        private int quantity = 100;

        private void Window_Drop(object sender, DragEventArgs e)
        {
            try
            {
                Cursor = Cursors.Wait;
                string[] paths = (string[])e.Data.GetData(DataFormats.FileDrop);
                paths = paths.Where(n => n.Contains("pdf")).ToArray();

                foreach (string path in paths)
                {
                    string fileName = Path.GetFileNameWithoutExtension(path);
                    FileInfo fi = new FileInfo(path);
                    DirectoryInfo dis = fi.Directory.CreateSubdirectory(fileName);

                    OutputPDF(path, dis);
                }

                MessageBox.Show("処理が完了しました");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                Cursor = Cursors.Arrow;
            }
        }

        private async void OutputPDF(string path,DirectoryInfo dis)
        {
            try
            {
                var file = await Windows.Storage.StorageFile.GetFileFromPathAsync(path);
                pdf = await PdfDocument.LoadFromFileAsync(file);
            }
            catch
            {
            }

            if (pdf == null)
            {
                return;
            }
            for (uint i = 0; i < pdf.PageCount; i++)
            {
                using (PdfPage page = pdf.GetPage(i))
                {
                    BitmapImage bi = new BitmapImage();

                    using(var stream = new Windows.Storage.Streams.InMemoryRandomAccessStream())
                    {
                        await page.RenderToStreamAsync(stream);

                        bi.BeginInit();
                        bi.CacheOption = BitmapCacheOption.OnLoad;
                        bi.StreamSource = stream.AsStream();
                        bi.EndInit();
                    }

                    JpegBitmapEncoder encoder = new JpegBitmapEncoder();
                    encoder.QualityLevel = quantity;
                    encoder.Frames.Add(BitmapFrame.Create(bi));

                    string destFilePath = dis.FullName+"\\" + (i + 1).ToString() + ".jpg";

                    using(var fileStream = new FileStream(destFilePath,FileMode.Create,FileAccess.Write))
                    {
                        encoder.Save(fileStream);
                    }
                }
            }

            pdf = null;
        }


    }
}
