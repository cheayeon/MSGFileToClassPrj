﻿using MSGFileToClassPrj.Models;
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
        List<MSGMessageModel> msgFiles = new List<MSGMessageModel>();

        public MainWindow()
        {
            InitializeComponent();

            //windowsFormsHost.Child = new System.Windows.Forms.RichTextBox();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openImportFile = new OpenFileDialog();
            openImportFile.Filter = "msg (*.msg) |*.msg|All Aplication (*) |*";

            if (openImportFile.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                foreach (string msgfile in openImportFile.FileNames)
                {
                    Stream messageStream = File.Open(msgfile, FileMode.Open, FileAccess.Read);
                    MSGMessageModel message;
                    try
                    {
                        message = new MSGMessageModel(messageStream);
                        msgFiles.Add(message);
                    }
                    catch (Exception) {
                        message = null;
                        // microsoft에서 제공하는 msg 파일이 아니거나, 파일이 손상되었음.
                        break;
                    }

                    messageStream.Close();
                    message.Dispose();

                    if (LoadRtfIntoRichTextBox(message.BodyByte))
                    {
                        result.Visibility = Visibility.Collapsed;
                        myRichTextBox.Visibility = Visibility.Visible;
                    }
                    else
                    {
                        result.Visibility = Visibility.Visible;
                        myRichTextBox.Visibility = Visibility.Collapsed;

                        result.Text += message.From + "\n";
                        result.Text += message.FromAdd + "\n";
                        result.Text += message.Recipients[0].DisplayName + "\n";
                        result.Text += message.Recipients[0].Email + "\n";
                        result.Text += "------------------------------------------\n";
                        result.Text += message.Subject + "\n";
                        result.Text += message.BodyText + "\n";
                        result.Text += message.BodyRTF + "\n";
                        result.Text += "------------------------------------------\n";
                        result.Text += message.Attachments[0].Filename + "\n";
                    }

                }
            }
            //string xx = msgFiles[0].Recipients[0].Type.ToString();
            //result.Text = "";
        }

        private bool LoadRtfIntoRichTextBox(byte[] rtfData)
        {
            try
            {
                using (MemoryStream rtfStream = new MemoryStream(rtfData))
                {
                    TextRange textRange = new TextRange(myRichTextBox.Document.ContentStart, myRichTextBox.Document.ContentEnd);
                    textRange.Load(rtfStream, System.Windows.DataFormats.Rtf);
                }
            }
            catch (Exception)
            {
                return false;
            }

            return true;
        }

        private void Btn_byteSave_Click(object sender, EventArgs e)
        {
            string extension = msgFiles[0].Attachments[0].Filename.Substring(msgFiles[0].Attachments[0].Filename.Contains(".") ?msgFiles[0].Attachments[0].Filename.LastIndexOf(".") : 0);

            SaveFileDialog openImportFile = new SaveFileDialog();
            openImportFile.Filter = $"{extension}|*{extension}";
            openImportFile.Title = msgFiles[0].Attachments[0].Filename;

            if (openImportFile.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                if (openImportFile.FileName != "")
                {
                    string dir = openImportFile.FileName; //경로 + 파일명            
                    FileStream file = new FileStream(dir, FileMode.Create);
                    file.Write(msgFiles[0].Attachments[0].Data, 0, msgFiles[0].Attachments[0].Data.Length);
                    file.Close();
                }
            }
        }


    }
}