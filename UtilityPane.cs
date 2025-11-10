using System;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Windows.Forms;
using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.WinForms;

namespace PowerEditAddIn
{
    public class UtilityPane : UserControl
    {
        private const int COLLAPSED_WIDTH = 28;
        private const int EXPANDED_WIDTH = 380;

        private Panel strip;
        private WebView2 web;
        private Timer slideTimer;
        private bool isExpanded = false;
        private int targetWidth;

        public event EventHandler<bool> ExpandedChanged;

        public UtilityPane()
        {
            BuildUi();
        }

        private void BuildUi()
        {
            this.Dock = DockStyle.Fill;

            strip = new Panel
            {
                Dock = DockStyle.Left,
                Width = COLLAPSED_WIDTH,
                BackColor = Color.FromArgb(44, 62, 80),
                Cursor = Cursors.Hand
            };
            strip.Paint += Strip_Paint;
            strip.Click += delegate { Toggle(); };

            web = new WebView2
            {
                Dock = DockStyle.Fill,
                Visible = false
            };

            slideTimer = new Timer { Interval = 10 };
            slideTimer.Tick += SlideTimer_Tick;

            this.Controls.Add(web);
            this.Controls.Add(strip);
            this.Load += UtilityPane_Load;
        }

        private async void UtilityPane_Load(object sender, EventArgs e)
        {
            try
            {
                // 1) WebView2 cache in LocalAppData
                string dataFolder = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "PowerEditAddin", "UtilityWebView2Data");
                Directory.CreateDirectory(dataFolder);

                var env = await CoreWebView2Environment.CreateAsync(userDataFolder: dataFolder);
                await web.EnsureCoreWebView2Async(env);

                // 2) Find UI source folder (project's /ui)
                string uiSource = FindUiSourceFolder();

                // 3) Destination folder in LocalAppData
                string uiDest = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "PowerEditAddin", "ui");
                Directory.CreateDirectory(uiDest);

                // 4) Copy files (if missing/newer)
                if (Directory.Exists(uiSource))
                {
                    foreach (string src in Directory.GetFiles(uiSource, "*.*", SearchOption.TopDirectoryOnly))
                    {
                        string name = Path.GetFileName(src);
                        string dst = Path.Combine(uiDest, name);

                        if (!File.Exists(dst) ||
                            File.GetLastWriteTimeUtc(src) > File.GetLastWriteTimeUtc(dst))
                        {
                            File.Copy(src, dst, true);
                        }
                    }
                }

                // 5) Load utility.html
                string htmlPath = Path.Combine(uiDest, "utility.html");
                if (File.Exists(htmlPath))
                {
                    web.Source = new Uri(htmlPath);
                }
                else
                {
                    web.NavigateToString(
                        "<html><body style='font-family:Segoe UI;color:red;padding:10px'>" +
                        "❌ utility.html not found<br><small>Expected at:<br>" + htmlPath + "</small></body></html>");
                }
            }
            catch (Exception ex)
            {
                web.NavigateToString(
                    "<html><body style='font-family:Segoe UI;color:red;padding:10px'>" +
                    "⚠️ WebView2 init failed:<br>" + ex.Message + "</body></html>");
            }
        }

        // Search upwards for a folder named "ui" that contains utility.html
        private string FindUiSourceFolder()
        {
            // Start at the physical DLL path (not shadow copy)
            string dllPath;
            try
            {
                // CodeBase returns file:// URL – convert to local path
                dllPath = new Uri(Assembly.GetExecutingAssembly().CodeBase).LocalPath;
            }
            catch
            {
                dllPath = Assembly.GetExecutingAssembly().Location;
            }

            string dir = Path.GetDirectoryName(dllPath) ?? "";
            // Try a few strategies:
            // 1) same folder /ui
            // 2) parent /ui (repeat up to 6 levels)
            for (int i = 0; i <= 6; i++)
            {
                string candidate = Path.Combine(dir, "ui");
                if (Directory.Exists(candidate) && File.Exists(Path.Combine(candidate, "utility.html")))
                    return candidate;

                string parent = Directory.GetParent(dir) != null ? Directory.GetParent(dir).FullName : null;
                if (string.IsNullOrEmpty(parent)) break;
                dir = parent;
            }

            // Last resort: AppDomain base
            string alt = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "ui");
            if (Directory.Exists(alt) && File.Exists(Path.Combine(alt, "utility.html")))
                return alt;

            // If nothing found, return empty (caller will show helpful message)
            return "";
        }

        private void Strip_Paint(object sender, PaintEventArgs e)
        {
            e.Graphics.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
            e.Graphics.TranslateTransform(strip.Width / 2f, strip.Height / 2f);
            e.Graphics.RotateTransform(-90);

            using (Font f = new Font("Segoe UI Semibold", 10f))
            using (SolidBrush b = new SolidBrush(Color.White))
            {
                string text = "PowerEdit Utility";
                SizeF size = e.Graphics.MeasureString(text, f);
                e.Graphics.DrawString(text, f, b, -size.Width / 2f, -size.Height / 2f);
            }

            e.Graphics.ResetTransform();
        }

        private void SlideTimer_Tick(object sender, EventArgs e)
        {
            int diff = targetWidth - this.Width;
            if (Math.Abs(diff) < 5)
            {
                this.Width = targetWidth;
                slideTimer.Stop();
                web.Visible = isExpanded;
                ExpandedChanged?.Invoke(this, isExpanded);
                return;
            }

            this.Width += diff / 5;           // easing
            this.Parent?.Refresh();
            this.BringToFront();
        }

        public void Toggle() { SetExpanded(!isExpanded); }

        public void SetExpanded(bool expand)
        {
            isExpanded = expand;
            targetWidth = expand ? EXPANDED_WIDTH : COLLAPSED_WIDTH;
            slideTimer.Start();
        }

        public int DesiredWidth
        {
            get { return isExpanded ? EXPANDED_WIDTH : COLLAPSED_WIDTH; }
        }
    }
}
