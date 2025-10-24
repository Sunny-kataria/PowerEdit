using System;
using System.IO;
using System.Text.Json;
using System.Windows.Forms;
using Microsoft.Web.WebView2.WinForms;
using Microsoft.Web.WebView2.Core;

namespace PowerEditAddIn
{
    public partial class DesignerPane : UserControl
    {
        public DesignerPane()
        {
            InitializeComponent();
            this.Load += DesignerPane_Load;
        }

        private void DesignerPane_Load(object sender, EventArgs e)
        {
            // EMPTY — no webview init here now
        }


        public void SendToClient(object payload)
        {
            webView21?.CoreWebView2?.PostWebMessageAsJson(JsonSerializer.Serialize(payload));
        }
    }
}
