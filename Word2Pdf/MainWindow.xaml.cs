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
using System.IO;
using System.Collections.ObjectModel;
using Forms = System.Windows.Forms;
using System.Collections;
using System.Windows.Controls.Primitives;
using Aspose.Words;
using Aspose.Pdf.Facades;

namespace Word2Pdf
{
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : Window
    {
        private String dir;
        private ObservableCollection<FileToChange> fileToChanges = new ObservableCollection<FileToChange>();

        public MainWindow()
        {
            InitializeComponent();
            fileList.ItemsSource = fileToChanges;
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            Forms.FolderBrowserDialog m_Dialog = new Forms.FolderBrowserDialog();

            if (m_Dialog.ShowDialog() != System.Windows.Forms.DialogResult.Cancel) {
                dir = m_Dialog.SelectedPath.Trim();
                this.wordDir.Text = dir;
            }
        }

        private void scanClick(object sender, RoutedEventArgs e)
        {
            String filter = "*.doc?";
            List<FileInfo> matchedFile = filterByExtension(dir, filter);
            foreach (FileInfo fileInfo in matchedFile) {
                fileToChanges.Add(new FileToChange() {
                    FileName = fileInfo.Name,
                    FileSize = (int)Math.Floor(fileInfo.Length / 1024.0)
                });
            }
            Console.WriteLine(String.Format("文件个数为:{0}", fileToChanges.Count));
        }

        private List<FileInfo> filterByExtension(String dir, List<String> extentions)
        {
            List<FileInfo> matchedFiles = new List<FileInfo>();
            DirectoryInfo directory = new DirectoryInfo(dir);
            FileInfo[] files = directory.GetFiles();
            foreach(FileInfo fileInfo in files) {
                foreach (String extention in extentions) {
                    if (extention.ToLower().Trim() == fileInfo.Extension) {
                        matchedFiles.Add(fileInfo);
                    }
                }
            }
            return matchedFiles;
        }

        private List<FileInfo> filterByExtension(String dir, String filter)
        {
            DirectoryInfo directory = new DirectoryInfo(dir);
            FileInfo[] files = directory.GetFiles(filter);
            return files.ToList<FileInfo>();
        }

        private void mouseMove(object sender, System.Windows.Input.MouseEventArgs e)
        {
            ListView listview = sender as ListView;
            if (e.LeftButton == MouseButtonState.Pressed)
               {
                IList list = listview.SelectedItems as IList;
                DataObject data = new DataObject(typeof(IList), list);
                if (list.Count > 0)
                {
                    DragDrop.DoDragDrop(listview, data, DragDropEffects.Move);
                }
            }
        }

        private void drop(object sender, System.Windows.DragEventArgs e)
        {
            if (e.Data.GetDataPresent(typeof(IList)))
            {
                IList toMoveFile = e.Data.GetData(typeof(IList)) as IList;
                int index = GetCurrentIndex(new GetPositionDelegate(e.GetPosition));
                if (index > -1)
                {
                    FileToChange fileToChange = toMoveFile[0] as FileToChange;
                    int OldFirstIndex = fileToChanges.IndexOf(fileToChange);
                    for (int i = 0; i < toMoveFile.Count; i++) {
                        fileToChanges.Move(OldFirstIndex, index);
                    }
                    fileList.SelectedItems.Clear();
                }
            }
        }

        private int GetCurrentIndex(GetPositionDelegate getPosition)
       {  
           int index = -1;  
           for (int i = 0; i<fileList.Items.Count; ++i)  
           {  
               ListViewItem item = GetListViewItem(i);  
               if (item != null && this.IsMouseOverTarget(item, getPosition))  
               {  
                   index = i;  
                   break;  
               }  
           }  
           return index;  
       }  
   
       private bool IsMouseOverTarget(Visual target, GetPositionDelegate getPosition)
       {  
           Rect bounds = VisualTreeHelper.GetDescendantBounds(target);  
           Point mousePos = getPosition((IInputElement)target);  
           return bounds.Contains(mousePos);  
       }  
   
       delegate Point GetPositionDelegate(IInputElement element);  
   
       ListViewItem GetListViewItem(int index)
       {  
           if (fileList.ItemContainerGenerator.Status != GeneratorStatus.ContainersGenerated)  
               return null;  
           return fileList.ItemContainerGenerator.ContainerFromIndex(index) as ListViewItem;  
       }

        private void convertClick(object sender, RoutedEventArgs e)
        {
            List<String> pdfFiles = new List<String>();
            foreach (FileToChange item in fileList.Items)
            {
                
                String filePath = dir + "/" + item.FileName;
                FileInfo wordFileInfo = new FileInfo(filePath);
                String fileBaseName = wordFileInfo.Name.Substring(0, wordFileInfo.Name.LastIndexOf("."));
                String pdfPath = dir + "/" + fileBaseName + ".pdf";
                word2Pdf(wordFileInfo.FullName, pdfPath);
                pdfFiles.Add(pdfPath);
            }
            MessageBox.Show("转换结束！", "提示");
            Aspose.Pdf.License li = new Aspose.Pdf.License();
            String appDir = System.AppDomain.CurrentDomain.SetupInformation.ApplicationBase;
            Console.WriteLine("path: {0}", appDir);
            li.SetLicense(appDir + "/" + "Aspose.Total.lic");
            PdfFileEditor pdfeditor = new PdfFileEditor();
            String destPdf = dir + "/" + "final.pdf";
            pdfeditor.Concatenate(pdfFiles.ToArray(), destPdf);
            MessageBox.Show("合并结束！", "提示");
        }

        private void word2Pdf(String wordPath, String pdfPath)
        {
            Document doc = new Document(wordPath);
            doc.Save(pdfPath, SaveFormat.Pdf);
        }
    }
}
