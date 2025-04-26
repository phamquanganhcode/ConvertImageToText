using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using AutoItX3Lib;
using System.IO;

namespace ConvertToText
{
    public partial class Form1 : Form
    {
        private ChromeDriver chromeDriver = null;
        string mainPath = @"D:\NCKH\@2018\Screenshot 2025-04-26 113411.png";

        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Openwebsite("https://www.google.com.vn/?hl=vi");
            GetString(mainPath);
            WriteToNotePad("abcs.txt");
        }
        private void Openwebsite(string url)
        {
            try
            {
                if (chromeDriver == null)
                {
                    ChromeOptions options = new ChromeOptions();
                    options.AddExcludedArgument("enable-automation");
                    options.AddArgument("--disable-webrtc");
                    options.AddArgument("--disable-blink-features=AutomationControlled");
                    options.AddArgument("user-data-dir=C:\\Users\\Pham Quang Anh\\AppData\\Local\\Google\\Chrome\\User Data\\Default");

                    chromeDriver = new ChromeDriver(options);
                }

                chromeDriver.Navigate().GoToUrl(url);
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Lỗi khi mở website: {ex.Message}");
                CloseBrowser();
                throw;
            }
        }
        private void CloseBrowser()
        {
            try
            {
                if (chromeDriver != null)
                {
                    // Đóng trình duyệt một cách sạch sẽ
                    chromeDriver.Quit();
                    chromeDriver.Dispose();
                    chromeDriver = null;
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Lỗi khi đóng trình duyệt: {ex.Message}");
            }
        }

        void GetString(string str_folder)
        {
            try
            {
                Thread.Sleep(TimeSpan.FromSeconds(4));

                var upload1 = chromeDriver.FindElement(By.CssSelector("body > div.L3eUgb > div.o3j99.ikrT4e.om7nvf > form > div:nth-child(1) > div.A8SBwf > div.RNNXgb > div.SDkEP > div.fM33ce.dRYYxd > div.nDcEnd > svg > path"));
                upload1.Click();

                Thread.Sleep(TimeSpan.FromSeconds(3));

                var upload2 = chromeDriver.FindElement(By.CssSelector("#ow12 > div.RNNXgb > div.M8H8pb > c-wiz > div.NzSfif > div > div.NrdQVe > div.f6GA0 > div.BH9rn > div.ZeVBtc > span"));
                upload2.Click();
                Thread.Sleep(TimeSpan.FromSeconds(3));

                var autoIT = new AutoItX3();
                if (autoIT.WinWait("Open", "", 10) == 1) // Chờ tối đa 10s
                {
                    autoIT.WinActivate("Open");
                    autoIT.Send(str_folder);
                    Thread.Sleep(1000);
                    autoIT.Send("{ENTER}");
                }
                else
                {
                    throw new Exception("Không tìm thấy cửa sổ upload");
                }

                //cần sửa tiêu đề sau đó ấn nút là xong
                Thread.Sleep(TimeSpan.FromSeconds(5));
                
                //chon 1 kí tự trong frame
                IWebElement element1 = chromeDriver.FindElement(By.CssSelector("#tsf > div:nth-child(1) > div.A8SBwf.BgPPrc > div.kzKS5c > div.o5UAab > div > c-wiz > div > div.PPaMi > img"));
                element1.Click();
                Thread.Sleep(TimeSpan.FromSeconds(5));

                //chon tat cả
                IWebElement element2 = chromeDriver.FindElement(By.CssSelector("#tsf > div:nth-child(1) > div.A8SBwf.BgPPrc > div.kzKS5c > div.o5UAab > div > c-wiz > div > div.d7TPDe > div.BRSHvf > div > div > div:nth-child(5) > button > span.DqO7jb"));
                element2.Click();

                Thread.Sleep(TimeSpan.FromSeconds(2));
                //sao chep
                IWebElement element3= chromeDriver.FindElement(By.CssSelector("#tsf > div:nth-child(1) > div.A8SBwf.BgPPrc > div.kzKS5c > div.o5UAab > div > c-wiz > div > div.d7TPDe > div.BRSHvf > div > div > div:nth-child(1) > button > span.DqO7jb"));
                element3.Click();


            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Lỗi khi upload file: {ex.Message}");
                throw;
            }
        }

        void ChangeNameFile()
        {
            string[] extensions = new[] { "*.jpg", "*.png", "*.jpeg" };
            var imageFiles = extensions.SelectMany(ext => Directory.GetFiles(mainPath, ext)).ToArray();
            // Lấy tất cả file ảnh trong folder

            foreach (string filePath in imageFiles)
            {
                string fileName = Path.GetFileName(filePath);

                // Nếu tên file có đuôi "_done" thì bỏ qua
                if (fileName.Contains("_done"))
                {
                    Console.WriteLine($"Bỏ qua: {fileName}");
                    continue;
                }

                // Xử lý với file đó
                Console.WriteLine($"Đang xử lý: {fileName}");

                // Sau khi xử lý xong → đổi tên file thêm _done
                string newFileName = Path.GetFileNameWithoutExtension(filePath) + "_done" + Path.GetExtension(filePath);
                string newFilePath = Path.Combine(mainPath, newFileName);

                File.Move(filePath, newFilePath);
                Console.WriteLine($"Đã đổi tên thành: {newFileName}");
            }

            Console.WriteLine("Hoàn tất.");
        }

        public string GetFirstFileName(string folderPath)
        {
            try
            {
                if (!Directory.Exists(folderPath))
                {
                    throw new DirectoryNotFoundException($"Thư mục không tồn tại: {folderPath}");
                }

                var files = Directory.GetFiles(folderPath);
                if (files.Length == 0) return null;

                return Path.GetFileName(files[0]);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Lỗi GetFirstFileName: {ex.Message}");
                throw;
            }
        }

        void WriteToNotePad(string name) 
        {
            Process.Start("notepad.exe");
            Thread.Sleep(TimeSpan.FromSeconds(1));
            SendKeys.SendWait("^n");
            Thread.Sleep(TimeSpan.FromSeconds(1));
            SendKeys.SendWait("^v");
            Thread.Sleep(TimeSpan.FromSeconds(2));
            SendKeys.SendWait("^s");
            var autoIT = new AutoItX3();
            if (autoIT.WinWait("Save as", "", 10) == 1) // Chờ tối đa 10s
            {
                autoIT.WinActivate("Save as");
                autoIT.Send(name);
                Thread.Sleep(1000);
                autoIT.Send("{ENTER}");
            }
            else
            {
                throw new Exception("Không tìm thấy cửa sổ upload");
            }
        }
        public string GetClipboardText()
        {
            string result = "";
            Thread thread = new Thread(() =>
            {
                result = Clipboard.GetText();
            });
            thread.SetApartmentState(ApartmentState.STA);
            thread.Start();
            thread.Join();
            return result;
        }

    }
}
