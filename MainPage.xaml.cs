using Microsoft.UI;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using Microsoft.UI.Xaml.Controls.Primitives;
using Microsoft.UI.Xaml.Media;
using Microsoft.UI.Xaml.Media.Imaging;
using System.Diagnostics;
using Windows.Data.Pdf;
using Windows.System;
using Windows.Storage.Streams;

// To learn more about WinUI, the WinUI project structure,
// and more about our project templates, see: http://aka.ms/winui-project-info.

namespace OfficeHub
{
    /// <summary>
    /// The main content page displayed inside the application window.
    /// Add your UI logic, event handlers, and data binding here.
    /// </summary>
    public sealed partial class MainPage : Page
    {
        private string version = "1.0.0";

        public MainPage()
        {
            InitializeComponent();

            // TODO: Add your initialization logic here.

            ContentFrame.Navigate(typeof(HomePage));

            this.Loaded += (s, e) =>
            {
                Page_Loaded();
            };

            OfficeHubVerT.Content = $"版本信息-v{version}";
        }

        private void MainNavView_MenuItemInvoked(NavigationView sender, NavigationViewItemInvokedEventArgs args)
        {
            if (args.InvokedItemContainer is NavigationViewItem item)
            {
                string tag = item.Tag?.ToString() ?? "";

                switch (tag)
                {
                    case "Home":
                        ContentFrame.Navigate(typeof(HomePage));
                        break;
                    case "Word":
                        LaunchOfficeApp("winword.exe");
                        break;
                    case "Excel":
                        LaunchOfficeApp("excel.exe");
                        break;
                    case "PowerPoint":
                        LaunchOfficeApp("powerpnt.exe");
                        break;
                    case "Outlook":
                        LaunchOfficeApp("outlook.exe");
                        break;
                    case "About":
                        ContentFrame.Navigate(typeof(AboutPage));
                        break;
                    case "OfficeHubVerT":
                        ContentFrame.Navigate(typeof(VerInfoPage), version);
                        break;
                    case "User":
                        UserTeachingTip.IsOpen = true;
                        break;
                }
            }
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

        private async Task<BitmapImage?> GetUserAvatarAsync()
        {
            try
            {
                var users = await User.FindAllAsync(UserType.LocalUser);
                if (users.Count == 0) return null;

                var user = users[0];
                var pictureRef = await user.GetPictureAsync(UserPictureSize.Size208x208);
                if (pictureRef == null) return null;

                using (var stream = await pictureRef.OpenReadAsync())
                {
                    var bitmap = new BitmapImage();
                    await bitmap.SetSourceAsync(stream);
                    return bitmap;
                }
            }
            catch
            {
                return null; 
            }
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
            UserNameText.Text = await GetUserDisplayNameAsync();
            ToolTipService.SetToolTip(UserNavItem, UserNameText.Text);
            UserTeachingTip.Title = UserNameText.Text;

            string userxname = Environment.UserName;
            UserTeachingTip.Subtitle = userxname;

            var avatar = await GetUserAvatarAsync();
            if (avatar != null)
            {
                UserAvatar.ProfilePicture = avatar;
                UserAvatar2.ProfilePicture = avatar;
            }
            else
            {
                UserAvatar.Initials = UserNameText.Text.Substring(0, 1).ToUpper();
                UserAvatar2.Initials = UserNameText.Text.Substring(0, 1).ToUpper();
            }
        }
    }
}
