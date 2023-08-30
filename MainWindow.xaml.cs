using Aspose.Cells;
using Aspose.Cells.Rendering;
using Aspose.Slides;
using Aspose.Words;
using Aspose.Words.Saving;
using PdfiumViewer;
using System;
using System.Collections.Generic;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace DocumentProcessingSoftware
{
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : Window
    {
        string filePath = "";

        string savaPath = "C:\\Users\\Administrator\\Downloads";
        public MainWindow()
        {
            InitializeComponent();
            label2.Content = "转换后的文件将存储到："+savaPath;
            if (!Directory.Exists(savaPath))
            {
                // 创建文件夹
                Directory.CreateDirectory(savaPath);
            }
        }

    

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            // 实例化一个文件选择对象
            Microsoft.Win32.OpenFileDialog dialog = new Microsoft.Win32.OpenFileDialog();
            dialog.DefaultExt = ".png";  // 设置默认类型
                                         // 设置可选格式
            dialog.Filter = @"Office Files|*.doc;*.docx;*.xls;*.xlsx;*.ppt;*.pptx;*.pdf" +
             "|All Files|*.*";
            // 打开选择框选择
            Nullable<bool> result = dialog.ShowDialog();
            if (result == true)
            {
                filePath = dialog.FileName; // 获取选择的文件名
                string fileName = dialog.SafeFileName;
                label1.Content = fileName;
            }

        }

        private void Button1_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(filePath))
            {
                MessageBox.Show("请选择文件");
            }
            else if (filePath.IndexOf("doc") > -1 || filePath.IndexOf("docx") > -1)
            {
                try
                {
                    WordToPng();
                    MessageBox.Show("转换成功，请到" + savaPath + "目录下查看图片。");
                }
                catch (Exception ex)
                {
                    MessageBox.Show("转换失败：" + ex.Message);
                    throw;
                }
            }
            else if (filePath.IndexOf("pdf") > -1)
            {
                try
                {
                    PdfToPng();
                    MessageBox.Show("转换成功，请到" + savaPath + "目录下查看图片。");
                }
                catch (Exception ex)
                {
                    MessageBox.Show("转换失败：" + ex.Message);
                    throw;
                }
            }
            else if (filePath.IndexOf("xls") > -1 || filePath.IndexOf("xlsx") > -1)
            {
                try
                {
                    ExcelToPng();
                    MessageBox.Show("转换成功，请到" + savaPath + "目录下查看图片。");
                }
                catch (Exception ex)
                {
                    MessageBox.Show("转换失败：" + ex.Message);
                    throw;
                }
            }
            else if (filePath.IndexOf("ppt") > -1 || filePath.IndexOf("pptx") > -1)
            {
                try
                {
                    PPTToPng();
                    MessageBox.Show("转换成功，请到" + savaPath + "目录下查看图片。");
                }
                catch (Exception ex)
                {
                    MessageBox.Show("转换失败：" + ex.Message);
                    throw;
                }
            }
            else {
                MessageBox.Show("不支持的格式！");
            }
        }

        private void WordToPng() {
            Document doc = new Document(filePath);
            Aspose.Words.Saving.ImageSaveOptions iso = new Aspose.Words.Saving.ImageSaveOptions(Aspose.Words.SaveFormat.Jpeg);
            iso.Resolution = 128;
            iso.PrettyFormat = true;
            for (int i = 0; i < doc.PageCount; i++)
            {
                iso.PageIndex = i;
                doc.Save(savaPath+"/wordtopng" + i + ".jpg", iso);
            }
        }

        private void PdfToPng() {
            var pdf = PdfDocument.Load(filePath);
            var pdfpage = pdf.PageCount;
            var pagesizes = pdf.PageSizes;
            for (int i = 1; i <= pdfpage; i++)
            {
                System.Drawing.Size size = new System.Drawing.Size();
                size.Height = (int)pagesizes[(i - 1)].Height;
                size.Width = (int)pagesizes[(i - 1)].Width;
                RenderPage(filePath, i, size, savaPath+"/pdftopng" + i + @".png");
            }
        }
        /// <summary>
        /// 将PDF转换为图片
        /// </summary>
        /// <param name="pdfPath">pdf文件位置</param>
        /// <param name="pageNumber">pdf文件张数</param>
        /// <param name="size">pdf文件尺寸</param>
        /// <param name="outputPath">输出图片位置与名称</param>
        public void RenderPage(string pdfPath, int pageNumber, System.Drawing.Size size, string outputPath, int dpi = 300)
        {
            using (var document = PdfDocument.Load(pdfPath))
            using (var stream = new FileStream(outputPath, FileMode.Create))
            using (var image = GetPageImage(pageNumber, size, document, dpi))
            {
                image.Save(stream, ImageFormat.Jpeg);
            }
        }
        private static System.Drawing.Image GetPageImage(int pageNumber, System.Drawing.Size size, PdfiumViewer.PdfDocument document, int dpi)
        {
            return document.Render(pageNumber - 1, size.Width, size.Height, dpi, dpi, PdfRenderFlags.Annotations);
        }
        private void ExcelToPng() {
            var workbook = new Workbook(filePath);
            for (int i = 0; i < workbook.Worksheets.Count; i++)
            {
                Worksheet sheet = workbook.Worksheets[0];
                sheet.PageSetup.LeftMargin = 0;
                sheet.PageSetup.RightMargin = 0;
                sheet.PageSetup.BottomMargin = 0;
                sheet.PageSetup.TopMargin = 0;

                ImageOrPrintOptions imgOptions = new ImageOrPrintOptions();
                imgOptions.ImageFormat = System.Drawing.Imaging.ImageFormat.Jpeg;

                imgOptions.OnePagePerSheet = true;
                imgOptions.PrintingPage = PrintingPageType.IgnoreBlank;

                SheetRender sr = new SheetRender(sheet, imgOptions);
                sr.ToImage(0, savaPath + "/exceltopng"+i+".png");
            }
            

        }

        private void PPTToPng() {
            using (Presentation presentation = new Presentation(filePath))
            {
                // 遍历每个幻灯片，将其转换为图片
                for (int slideIndex = 0; slideIndex < presentation.Slides.Count; slideIndex++)
                {
                    // 获取幻灯片对象
                    ISlide slide = presentation.Slides[slideIndex];

                    // 转换为图片
                    using (System.Drawing.Bitmap image = slide.GetThumbnail(1f, 1f))
                    {
                        // 保存图片
                        string imagePath = savaPath+"/ppttopng"+ slideIndex + 1 + ".png";
                        image.Save(imagePath, System.Drawing.Imaging.ImageFormat.Png);
                    }
                }
            }
        }
    }
}
