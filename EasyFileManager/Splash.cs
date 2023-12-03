using System.Reflection;

using WinFormsLib;

namespace EasyFileManager
{
    public partial class Splash : Form
    {
        public Splash(TimeSpan? timeSpan = null)
        {
            InitializeComponent();
            Load += Splash_Load;
            Click += Splash_Click;
            Disposed += Splash_Disposed;
            if (timeSpan != null) { Utils.Delay(Close, (int)timeSpan.Value.TotalMilliseconds); }
        }

        private void Splash_Load(object? sender, EventArgs e)
        {
            BringToFront();
            Assembly assembly = Assembly.GetExecutingAssembly();
            TitleLabel.Text = assembly.GetTitle();
            VersionLabel.Text = $"Version {assembly.GetFileVersion()}";
            string k = $"© {assembly.GetCompany()}";
            string v = assembly.GetCopyright();
            CopyrightLinkLabel.Text = k;
            CopyrightLinkLabel.UpdateLinks(Utils.GetLinkLabelLinks(k, new() { { k, v } }));
        }

        private void Splash_Disposed(object? sender, EventArgs e)
        {
            Load -= Splash_Load;
            Click -= Splash_Click;
            Disposed -= Splash_Disposed;
        }

        private void Splash_Click(object? sender, EventArgs e) => Close();
    }
}
