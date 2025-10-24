using Microsoft.Office.Interop.Word;
using Microsoft.Office.Tools;
using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.WinForms;
using System;
using System.IO;
using System.Threading.Tasks;
using Office = Microsoft.Office.Core;
using Word = Microsoft.Office.Interop.Word;


namespace PowerEditAddIn
{
    public partial class ThisAddIn
    {
        private CustomTaskPane _pane;
        private DesignerPane _control;

        private async void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            _control = new DesignerPane();
            _pane = this.CustomTaskPanes.Add(_control, "PowerEdit");
            _pane.DockPosition = Office.MsoCTPDockPosition.msoCTPDockPositionLeft;
            _pane.Width = 420;
            _pane.Visible = false;

            // === MOVE INIT HERE ===
            var dataFolder = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "PowerEditAddin", "WebView2Data");
            Directory.CreateDirectory(dataFolder);

            var env = await CoreWebView2Environment.CreateAsync(
                userDataFolder: dataFolder);

            await _control.webView21.EnsureCoreWebView2Async(env);

            string basePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "ui");
            _control.webView21.CoreWebView2.SetVirtualHostNameToFolderMapping(
                "app.local", basePath,
                CoreWebView2HostResourceAccessKind.Allow);

            _control.webView21.CoreWebView2.Navigate("https://app.local/index.html");
            _control.webView21.CoreWebView2.WebMessageReceived += (s, args) =>
            {
                var doc = System.Text.Json.JsonDocument.Parse(args.WebMessageAsJson);
                var root = doc.RootElement;
                var type = root.GetProperty("type").GetString();

                if (type == "browse")
                    BrowseAndOpenDocx();

                else if (type == "runAction")
                {
                    var action = root.GetProperty("action").GetString();
                    RunSelectedAction(action);
                }
            };

        }

        public void TogglePane()
        {
            if (_pane == null)
            {
                System.Windows.Forms.MessageBox.Show("Pane is NULL!");
                return;
            }

            _pane.Visible = !_pane.Visible;
        }
        public void BrowseAndOpenDocx()
        {
            try
            {
                using (var dlg = new System.Windows.Forms.OpenFileDialog())
                {
                    dlg.Filter = "Word Documents|*.docx";
                    dlg.Title = "Select a Word Document";

                    if (dlg.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                    {
                        // Get current open document
                        var activeDoc = this.Application.ActiveDocument;

                        // Clear existing content if needed
                        activeDoc.Content.Delete();

                        // Insert file INTO current document
                        activeDoc.Range(0, 0).InsertFile(dlg.FileName);

                        // Keep pane visible
                        _pane.Visible = true;

                        NotifyClient("Loaded into same document: " + Path.GetFileName(dlg.FileName));
                    }
                }
            }
            catch (Exception ex)
            {
                NotifyClient("Open failed: " + ex.Message);
            }
        }



        public void RunSelectedAction(string key)
        {
            // Abhi ke liye sirf testing ke liye
            NotifyClient("Action selected: " + key);
        }


        public void InsertTextIntoWord(string text)
        {
            try
            {
                var sel = this.Application.Selection;
                sel.TypeText(text ?? "");
            }
            catch (System.Exception ex)
            {
                System.Windows.Forms.MessageBox.Show("Insert failed: " + ex.Message);
            }
        }

        public void NotifyClient(string msg)
        {
            _control?.SendToClient(new { type = "notify", text = msg });
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e) { }

        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
    }
}
