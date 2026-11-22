using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using Microsoft.UI.Xaml.Controls.Primitives;
using Microsoft.UI.Xaml.Data;
using Microsoft.UI.Xaml.Input;
using Microsoft.UI.Xaml.Media;
using Microsoft.UI.Xaml.Media.Imaging;
using Microsoft.UI.Xaml.Navigation;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices.WindowsRuntime;
using Windows.Foundation;
using Windows.Foundation.Collections;
using Windows.System;

// To learn more about WinUI, the WinUI project structure,
// and more about our project templates, see: http://aka.ms/winui-project-info.

namespace OfficeHub
{
    /// <summary>
    /// An empty page that can be used on its own or navigated to within a Frame.
    /// </summary>
    public sealed partial class HomePage : Page
    {
        public HomePage()
        {
            InitializeComponent();

            this.Loaded += (s, e) =>
            {
                Page_Loaded();
            };

            Uri DOUri = new Uri("https://www.microsoft.com/zh-cn/download/details.aspx?id=55374");
            Uri BOUri = new Uri("https://www.microsoft.com/zh-cn/microsoft-365/buy/compare-all-microsoft-365-products?icid=mscom_marcom_M365#FamilyAnnualPaidYearly");
            DownloadOfficeBtn.Click += async (s, e) => await Windows.System.Launcher.LaunchUriAsync(DOUri);
            BuyOfficeBtn.Click += async (s, e) => await Windows.System.Launcher.LaunchUriAsync(BOUri);

            OpWord.Click += async (s, e) => LaunchOfficeApp("winword.exe");
            OpExcel.Click += async (s, e) => LaunchOfficeApp("excel.exe");
            OpPPT.Click += async (s, e) => LaunchOfficeApp("powerpnt.exe");
            OpOutlook.Click += async (s, e) => LaunchOfficeApp("outlook.exe");
            OpOneNote.Click += async (s, e) => LaunchOfficeApp("https://onenote.cloud.microsoft/");
            OpOneDrive.Click += async (s, e) => LaunchOfficeApp("https://onedrive.live.com/");
            OpToDo.Click += async (s, e) => LaunchOfficeApp("ms-todo://");

            OpBing.Click += async (s, e) => LaunchOfficeApp("https://cn.bing.com");
            OpMsn.Click += async (s, e) => LaunchOfficeApp("https://www.msn.cn");
        }

        private void LaunchOfficeApp(string exeName)
        {
            try
            {
                ProcessStartInfo startInfo = new ProcessStartInfo
                {
                    FileName = exeName,
                    UseShellExecute = true
                };
                Process.Start(startInfo);
            }
            catch
            {
                switch (exeName)
                {
                    case "winword.exe":
                        ShowDialog_ANF("Word");
                        break;
                    case "excel.exe":
                        ShowDialog_ANF("Excel");
                        break;
                    case "powerpnt.exe":
                        ShowDialog_ANF("PowerPoint");
                        break;
                    case "outlook.exe":
                        ShowDialog_ANF("Outlook");
                        break;
                    case "":
                        ShowDialog_ANF("OneDrive");
                        break;
                    default:
                        ShowDialog_ANF("此应用");
                        break;
                }
            }
        }

        private async void ShowDialog_ANF(string OfficeAppName)
        {
            ContentDialog dialog = new ContentDialog();

            dialog.XamlRoot = this.XamlRoot;
            dialog.Style = Application.Current.Resources["DefaultContentDialogStyle"] as Style;
            dialog.Title = "没有找到此Office应用";
            dialog.PrimaryButtonText = "确定";
            dialog.CloseButtonText = "关闭";
            dialog.DefaultButton = ContentDialogButton.Primary;
            dialog.Content = $"请确保您已安装{OfficeAppName}";

            var result = await dialog.ShowAsync();
        }

        private string SetGreet()
        {
            TimeSpan nowTime = DateTime.Now.TimeOfDay;

            TimeSpan morng = new TimeSpan(7, 0, 0);
            TimeSpan noon = new TimeSpan(12, 0, 0);
            TimeSpan aftnoon = new TimeSpan(14, 0, 0);
            TimeSpan eveng = new TimeSpan(18, 0, 0);
            string Greet = "";

            if (nowTime >= morng && nowTime < noon)
            {
                Greet = "早上";
            }
            else if (nowTime >= noon && nowTime < aftnoon)
            {
                Greet = "中午";
            }
            else if (nowTime >= aftnoon && nowTime < eveng)
            {
                Greet = "下午";
            }
            else if (nowTime >= eveng)
            {
                Greet = "晚上";
            }

            return Greet;
        }

        private async Task<string> GetUserDisplayNameAsync()
        {
            try
            {
                var users = await User.FindAllAsync(UserType.LocalUser);
                if (users.Count > 0)
                {
                    var props = await users[0].GetPropertiesAsync(new[] { KnownUserProperties.DisplayName });
                    string name = props[KnownUserProperties.DisplayName] as string ?? "";
                    if (!string.IsNullOrEmpty(name)) return name;
                }
            }
            catch { }

            return Environment.UserName;
        }

        private async void Page_Loaded()
        {
            string userName = await GetUserDisplayNameAsync();

            string Greet = SetGreet();
            GreetText.Text = $"{Greet}好，\n{userName}。\n有什么要完成的？";
        }
    }
}
