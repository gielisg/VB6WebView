using Microsoft.Web.WebView2.Core;
using System;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;

namespace WpfBrowserLib
{
    public partial class BrowserControl : UserControl
    {
        // Default URL
        public string DefaultUrl { get; set; } =
            "https://gielisg.github.io/ServiceDesk/#/auth/signin?SiteId=Demo3";

        public BrowserControl()
        {
            InitializeComponent();
            this.Loaded += BrowserControl_Loaded;
        }

        private async void BrowserControl_Loaded(object sender, RoutedEventArgs e)
        {
            try
            {
                await EnsureCoreWebView2Async();

                // ✅ Block right-click and common shortcuts
                string script = @"
                    document.addEventListener('contextmenu', function(e){
                        e.preventDefault();
                    }, true);

                    document.addEventListener('keydown', function(e){
                        if (
                            (e.ctrlKey && (e.key === 's' || e.key === 'S')) ||
                            (e.ctrlKey && (e.key === 'u' || e.key === 'U')) ||
                            (e.ctrlKey && (e.key === 'o' || e.key === 'O')) ||
                            (e.ctrlKey && e.shiftKey && (e.key === 'I' || e.key === 'i')) ||
                            (e.key === 'F12')
                        )
                        {
                            e.preventDefault();
                            e.stopPropagation();
                            return false;
                        }
                    }, true);
                ";

                await webView.CoreWebView2.AddScriptToExecuteOnDocumentCreatedAsync(script);

                //Block popups / external windows
                webView.CoreWebView2.NewWindowRequested += (s, ev) =>
                {
                    ev.Handled = true;
                };

                // Navigate to site
                webView.CoreWebView2.Navigate(DefaultUrl);
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    "WebView2 initialization failed:\n\n" + ex.Message,
                    "WebView2 Error",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error
                );
            }
        }

        private Task EnsureCoreWebView2Async()
        {
            var tcs = new TaskCompletionSource<bool>();

            if (webView.CoreWebView2 != null)
            {
                tcs.SetResult(true);
            }
            else
            {
                webView.CoreWebView2InitializationCompleted += (s, e) =>
                {
                    if (e.IsSuccess)
                        tcs.SetResult(true);
                    else
                        tcs.SetException(e.InitializationException);
                };

                webView.EnsureCoreWebView2Async();
            }

            return tcs.Task;
        }
    }
}
