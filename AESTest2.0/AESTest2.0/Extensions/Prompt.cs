using System.Windows.Forms;

namespace AESTest2._0.Extensions
{
    public static class Prompt
    {
        public static string ShowDialog(string text, string Caption)
        {
            Form prompt = new Form()
            {
                Width = 500,
                Height = 150,
                FormBorderStyle = FormBorderStyle.FixedDialog,
                Text = Caption,
                StartPosition = FormStartPosition.CenterScreen,
            };
            Label textLabel = new Label() { Left = 50, Top = 20, Text = text };
            TextBox tbox = new TextBox() { Left = 50, Top = 50, Width = 400 };
            Button btn = new Button() { Text = "Ок", Left = 350, Width = 100, Top = 70, DialogResult = DialogResult.OK };
            btn.Click += (sender, e) => { prompt.Close(); };
            prompt.Controls.Add(textLabel);
            prompt.Controls.Add(tbox);
            prompt.Controls.Add(btn);
            prompt.AcceptButton = btn;
            return prompt.ShowDialog() == DialogResult.OK ? tbox.Text : "";
        }
    }

}
