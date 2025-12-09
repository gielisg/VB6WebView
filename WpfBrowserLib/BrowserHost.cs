using System;
using System.Runtime.InteropServices;
using System.Windows;

namespace WpfBrowserLib
{
    [ComVisible(true)]
    [Guid("8E2F0F3D-9D1D-4C86-9A88-123456789ABC")]  
    [ProgId("WpfBrowserLib.BrowserHost")]
    public class BrowserHost
    {
        public void ShowBrowser()
        {
            try
            {
                Application.Current.Dispatcher.Invoke(() =>
                {
                    BrowserWindow win = new BrowserWindow();
                    win.ShowDialog(); 
                });
            }
            catch
            {
                Application app = new Application();
                BrowserWindow win = new BrowserWindow();
                win.ShowDialog();
            }
        }
    }
}
