using IWshRuntimeLibrary;
using Microsoft.Win32;
using Newtonsoft.Json.Linq;
using RestSharp;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.IO.Packaging;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Forms;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Animation;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace dot_Installer
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public static string InstallationPath = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86) + "\\Dot HQ\\Browser\\Application";
        public static bool createShortcut = true;
        public MainWindow()
        {
            if (DarkMode.GetWindowsTheme() == DarkMode.WindowsTheme.Dark)
            {
                this.Resources["BackgroundColor"] = new SolidColorBrush(Color.FromRgb(47, 49, 54));
                this.Resources["BackgroundBoxColor"] = new SolidColorBrush(Color.FromRgb(70, 74, 82));
                this.Resources["ForegroundColor"] = Brushes.White;
                this.Resources["Border"] = new SolidColorBrush(Color.FromRgb(74, 74, 74));
            }
            InitializeComponent();
            Splash_Screen.Visibility = Visibility.Visible;
            ScreenOne.Visibility = Visibility.Hidden;
            ScreenTwo.Visibility = Visibility.Hidden;
            CustomPathInputBox.Text = InstallationPath;
            updateCheckbox();

            splash();
        }

        public void updateCheckbox()
        {
            if (createShortcut)
            {
                if (DarkMode.GetWindowsTheme() == DarkMode.WindowsTheme.Dark)
                    DesktopLnkCheckImage.Source = new BitmapImage(new Uri("pack://application:,,,/dot Installer;component/Assets/checkbox_dark_checked.png"));
                else
                    DesktopLnkCheckImage.Source = new BitmapImage(new Uri("pack://application:,,,/dot Installer;component/Assets/checkbox_light_checked.png"));
            }
            else
            {
                if (DarkMode.GetWindowsTheme() == DarkMode.WindowsTheme.Dark)
                    DesktopLnkCheckImage.Source = new BitmapImage(new Uri("pack://application:,,,/dot Installer;component/Assets/checkbox_dark.png"));
                else
                    DesktopLnkCheckImage.Source = new BitmapImage(new Uri("pack://application:,,,/dot Installer;component/Assets/checkbox_light.png"));
            }
        }

        public void animate(UIElement element, bool toVisible, int ms = 750)
        {
            DoubleAnimation animation = new DoubleAnimation();
            animation.To = toVisible ? 1: 0;
            animation.From = toVisible ? 0 : 1;
            animation.Duration = TimeSpan.FromMilliseconds(ms);
            animation.EasingFunction = new QuadraticEase();

            Storyboard sb = new Storyboard();
            sb.Children.Add(animation);
            element.Opacity = 1;
            element.Visibility = Visibility.Visible;

            Storyboard.SetTarget(sb, element);
            Storyboard.SetTargetProperty(sb, new PropertyPath(OpacityProperty));

            sb.Begin();

            sb.Completed += delegate (object sender, EventArgs e)
            {
                element.Visibility = Visibility.Collapsed;
            };
        }
        public void animateProgress(System.Windows.Controls.ProgressBar progressBar, int newValue, int ms = 100)
        {
            DoubleAnimation animation = new DoubleAnimation();
            animation.To = newValue;
            animation.From = progressBar.Value;
            animation.Duration = TimeSpan.FromMilliseconds(ms);
            animation.EasingFunction = new QuadraticEase();

            Storyboard sb = new Storyboard();
            sb.Children.Add(animation);

            Storyboard.SetTarget(sb, progressBar);
            Storyboard.SetTargetProperty(sb, new PropertyPath(System.Windows.Controls.ProgressBar.ValueProperty));

            sb.Begin();
        }
        public async Task splash()
        {
            await Task.Delay(1250);
            animate(Splash_Screen, false);
            animate(ScreenOne, true);
        }

        private void Window_MouseDown(object sender, MouseButtonEventArgs e)
        {
            if (e.ChangedButton == MouseButton.Left)
                this.DragMove();
        }

        [STAThread]
        private void Browse_Click(object sender, RoutedEventArgs e)
        {
            FolderBrowserDialog dialog = new FolderBrowserDialog();
            dialog.ShowNewFolderButton = true;
            dialog.ShowDialog();
        }
        private void InstallDot_Click(object sender, RoutedEventArgs e)
        {
            InstallationPath = CustomPathInputBox.Text;
            if (InstallationPath.EndsWith("/") || InstallationPath.EndsWith("\\"))
            {
                System.Windows.MessageBox.Show("Installation path can't end with a slash", "DotHQ");
                return;
            }
            animate(ScreenOne, false);
            animate(ScreenTwo, true);

            doInstall();
        }
        public void updateDoing(string content)
        {
            CurrentlyDoing.Content = content;
            animate(CurrentlyDoing, true, 150);
        }

        public async Task doInstall()
        {
            try
            {
                MainGrid.Focus();
                updateDoing("Contacting GitHub");
                animateProgress(dotProgress, 33);
                RestClient client = new RestClient("https://api.github.com/repos/dothq/browser/");
                RestRequest request = new RestRequest("releases");
                var response = await client.ExecuteAsync(request);

                if (!response.Content.StartsWith("["))
                {
                    throw new Exception("Server sent an invalid response");
                }
                updateDoing("Parsing Data");
                animateProgress(dotProgress, 66);
                Debug.WriteLine(response.Content);
                JArray releases = JArray.Parse(response.Content);
                JObject latestRelease = releases[0].Value<JObject>();
                JArray assets = latestRelease["assets"].Value<JArray>();
                JObject asset = null;
                for (int i = 0; i < assets.Count; i++)
                {
                    JObject thisAsset = assets[i].Value<JObject>();
                    if (thisAsset["content_type"].Value<string>() == "application/x-zip-compressed")
                    {
                        asset = thisAsset;
                        break;
                    }
                }
                if (asset == null)
                {
                    throw new Exception("Unable to find archive");
                }
                updateDoing("Downloading archive");
                animateProgress(dotProgress, 100);
                WebClient dlClient = new WebClient();
                DateTime lastUpdate = DateTime.Now;
                dlClient.DownloadProgressChanged += (object sender, DownloadProgressChangedEventArgs e) =>
                {
                    if (lastUpdate.Second == DateTime.Now.Second) return;
                    lastUpdate = DateTime.Now;
                    Debug.WriteLine("Downloaded  " + (double)e.BytesReceived);
                    Debug.WriteLine("To Download " + (double)e.TotalBytesToReceive);
                    Debug.WriteLine("Percent     " + ((100*e.BytesReceived / e.TotalBytesToReceive)).ToString());
                    animateProgress(dotProgress, Convert.ToInt32(((100 * e.BytesReceived / e.TotalBytesToReceive) + 100)));
                };
                string tempFileName = System.IO.Path.GetTempFileName();
                await dlClient.DownloadFileTaskAsync(new Uri(asset["browser_download_url"].Value<string>()), tempFileName);
                if (Directory.Exists(InstallationPath))
                    deleteFolder(InstallationPath);
                Directory.CreateDirectory(InstallationPath);
                updateDoing("Unpacking");
                ZipFile.ExtractToDirectory(tempFileName, InstallationPath);

                string DotBinary = "";
                string[] files = Directory.GetFiles(InstallationPath);
                for (int i = 0; i < files.Length; i++)
                {
                    if (files[i].EndsWith(".exe"))
                        DotBinary = files[i];
                }

                if (createShortcut)
                {
                    updateDoing("Creating Shortcut");
                    object shDesktop = "Desktop";
                    WshShell shell = new WshShell();
                    IWshShortcut shortcut = (IWshShortcut)shell.CreateShortcut((string)shell.SpecialFolders.Item(ref shDesktop) + @"\Dot Browser.lnk");
                    shortcut.Description = "Start Dot Browser";
                    shortcut.TargetPath = DotBinary;
                    shortcut.Save();
                }

                updateDoing("Dot is now installed");
                animateProgress(dotProgress, 300, 1000);
                animate(dotProgress, false, 1000);
                await Task.Delay(1000);
                Process.Start(DotBinary);
                await Task.Delay(5000);
                Process.GetCurrentProcess().Kill();
            }
            catch (Exception e)
            {
                updateDoing(e.Message);
                Close.Visibility = Visibility.Visible;
                dotProgress.Foreground = Brushes.Red;
            }
        }

        public static void deleteFolder(string path)
        {
            foreach (var dir in Directory.GetDirectories(path))
            {
                deleteFolder(dir);
            }
            foreach (var file in Directory.GetFiles(path))
            {
                System.IO.File.Delete(file);
            }
            Directory.Delete(path);
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            if (createShortcut)
            {
                createShortcut = false;
            }
            else
            {
                createShortcut = true;
            }
            updateCheckbox();
        }

        private void ViewMore_Click(object sender, RoutedEventArgs e)
        {
            if (ViewMoreContent.Visibility == Visibility.Visible)
            {
                ViewMore.Content = "⌵ View More";
                ViewMoreContent.Visibility = Visibility.Hidden;
            }
            else
            {
                ViewMore.Content = "ߍ View Less";
                ViewMoreContent.Visibility = Visibility.Visible;
            }
        }

        private void Close_Click(object sender, RoutedEventArgs e)
        {
            Process.GetCurrentProcess().Kill();
        }
    }
}
