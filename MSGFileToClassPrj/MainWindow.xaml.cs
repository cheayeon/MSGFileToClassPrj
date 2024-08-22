using Microsoft.Web.WebView2.Core;
using MSGFileToClassPrj.Models;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Forms;
using System.Windows.Input;
using System.Windows.Markup;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace MSGFileToClassPrj
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        Dictionary<string, MSGMessageModel> msgFiles = new Dictionary<string, MSGMessageModel>();
        MSGMessageModel nowShowMSG;

        public string tempPath { get; private set; }

        List<string> OpenMSGFilesPaths = new List<string>();

        public string html;

        public MainWindow()
        {
            InitializeComponent();

            tempPath = System.IO.Path.Combine(
                                        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                                        "Temp",
                                        System.AppDomain.CurrentDomain.FriendlyName.Substring(0, System.AppDomain.CurrentDomain.FriendlyName.IndexOf(".")));

            if (Directory.Exists(tempPath))
                Directory.Delete(tempPath, true);

            Directory.CreateDirectory(tempPath);
        }

        private async void Window_Loaded(object sender, RoutedEventArgs e)
        {
            var env = await CoreWebView2Environment.CreateAsync(userDataFolder: tempPath);
            await webView.EnsureCoreWebView2Async();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openImportFile = new OpenFileDialog();
            openImportFile.Filter = "msg (*.msg) |*.msg|All Aplication (*) |*";

            if (openImportFile.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                foreach (string msgfile in openImportFile.FileNames)
                {
                    MSGMessageModel message;
                    if (msgFiles.ContainsKey(msgfile))
                    {
                        message = msgFiles[msgfile];
                        try
                        {
                            webView.CoreWebView2?.Navigate("about:blank");
                            webView.CoreWebView2.Navigate(message.TempPath);

                            result.Visibility = Visibility.Collapsed;
                            webView.Visibility = Visibility.Visible;
                        }
                        catch (Exception)
                        {
                            result.Visibility = Visibility.Visible;
                            webView.Visibility = Visibility.Collapsed;

                            result.Text += message.From + "\n";
                            result.Text += message.FromAdd + "\n";
                            result.Text += message.Recipients[0].DisplayName + "\n";
                            result.Text += message.Recipients[0].Email + "\n";
                            result.Text += "------------------------------------------\n";
                            result.Text += message.Subject + "\n";
                            result.Text += message.BodyText + "\n";
                            //result.Text += message.BodyRTF + "\n";
                            result.Text += "------------------------------------------\n";
                            result.Text += message.Attachments[0].Filename + "\n";
                        }
                    }
                    else
                    {
                        Stream messageStream = File.Open(msgfile, FileMode.Open, FileAccess.Read);
                        try
                        {
                            message = new MSGMessageModel(messageStream);
                            msgFiles.Add(msgfile, message);
                        }
                        catch (Exception) {
                            message = null;
                            // microsoft에서 제공하는 msg 파일이 아니거나, 파일이 손상되었음.
                            break;
                        }

                        messageStream.Close();
                        message.Dispose();

                        if (LoadRtfIntoRichTextBox(message))
                        {
                            result.Visibility = Visibility.Collapsed;
                            webView.Visibility = Visibility.Visible;
                        }
                        else
                        {
                            result.Visibility = Visibility.Visible;
                            webView.Visibility = Visibility.Collapsed;

                            result.Text += message.From + "\n";
                            result.Text += message.FromAdd + "\n";
                            result.Text += message.Recipients[0].DisplayName + "\n";
                            result.Text += message.Recipients[0].Email + "\n";
                            result.Text += "------------------------------------------\n";
                            result.Text += message.Subject + "\n";
                            result.Text += message.BodyText + "\n";
                            //result.Text += message.BodyRTF + "\n";
                            result.Text += "------------------------------------------\n";
                            result.Text += message.Attachments[0].Filename + "\n";

                        }
                    }
                    result.Text += message.Dates.ToString();

                    nowShowMSG = message;
                }
            }
            //string xx = msgFiles[0].Recipients[0].Type.ToString();
            //result.Text = "";
        }

        private bool LoadRtfIntoRichTextBox(MSGMessageModel message)
        {
            try
            { 
                string MSGTempPath = tempPath + "\\" + Guid.NewGuid().ToString();
                Directory.CreateDirectory(MSGTempPath);
                TextReader stringReader = new StringReader(message.BodyRTF);
                html = RtfPipe.Rtf.ToHtml(new RtfPipe.RtfSource(stringReader));

                foreach (MSGAttachmentModel setAttach in message.Attachments)
                {
                    string AttachFileName = setAttach.Filename;

                    if (String.IsNullOrEmpty(setAttach.Filename))
                        AttachFileName = "Invalid File";

                    string imgPath = MSGTempPath + "\\" + AttachFileName;

                    FileInfo fileInfo = new FileInfo(imgPath);
                    File.WriteAllBytes(fileInfo.FullName, setAttach.Data);
                    
                    if (!String.IsNullOrEmpty(html))
                    {
                        if (!string.IsNullOrEmpty(setAttach.ContentId) && html.Contains(setAttach.ContentId))
                            html = html.Replace("cid:" + setAttach.ContentId, (fileInfo.FullName));
                    }
                }

                // 한번 저장한 후에 불러야만 그림등이 오류없이 보여짐.
                message.TempPath = MSGTempPath + "\\" + message.Subject + ".html";
                FileInfo fileIn = new FileInfo(message.TempPath);
                File.WriteAllText(fileIn.FullName, html);

                webView.CoreWebView2?.Navigate("about:blank");
                webView.CoreWebView2.Navigate(message.TempPath);
            }
            catch (Exception e)
            {
                return false;
            }

            return true;
        }

        private void Btn_byteSave_Click(object sender, EventArgs e)
        {
            //첨부파일 저장용
            string extension = nowShowMSG.Attachments[0].Filename.Substring(nowShowMSG.Attachments[0].Filename.Contains(".") ? nowShowMSG.Attachments[0].Filename.LastIndexOf(".") : 0);

            SaveFileDialog openImportFile = new SaveFileDialog();
            openImportFile.Filter = $"{extension}|*{extension}";
            openImportFile.Title = nowShowMSG.Attachments[0].Filename;
            openImportFile.FileName = nowShowMSG.Attachments[0].Filename;

            if (openImportFile.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                if (openImportFile.FileName != "")
                {
                    string dir = openImportFile.FileName; //경로 + 파일명            
                    FileStream file = new FileStream(dir, FileMode.Create);
                    file.Write(nowShowMSG.Attachments[0].Data, 0, nowShowMSG.Attachments[0].Data.Length);
                    file.Close();
                }
            }
        }

        private void Window_Closed(object sender, EventArgs e)
        {
            if (Directory.Exists(tempPath))
                Directory.Delete(tempPath, true);
        }
    }
}
