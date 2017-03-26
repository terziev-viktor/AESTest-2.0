using System;
using System.IO;
using System.Windows.Forms;

namespace AESTest2._0
{
    public partial class PasswordForm : Form
    {
        private string path = @"C:\Windows\pass.txt";
        public PasswordForm()
        {
            InitializeComponent();
        }

        private void btnExit_Click(object sender, EventArgs e)
        {
            using(StreamReader reader = new StreamReader(this.path))
            {
                string decrypted = reader.ReadLine();
                string encrypted = Protection.Crypt(inp.Text);
                if(decrypted == encrypted)
                {
                    ExplorerManager.Start();
                    Application.Exit();
                    //ExplorerManager.Start(); // starting explorer.exe
                }
                else
                {
                    MessageBox.Show("Грешна парола!");
                    this.Close();
                }
            }
        }

        private void btnChangePass_Click(object sender, EventArgs e)
        {
            string pass;
            using(StreamReader reader = new StreamReader(this.path))
            {
                pass = reader.ReadLine();
            }

            string encrypted = Protection.Crypt(inp_1.Text);

            if (inp_2.Text == inp_3.Text && pass == encrypted)
            {
                string newPassEncrypted = Protection.Crypt(inp_2.Text);
                using(StreamWriter w = new StreamWriter(this.path))
                {
                    w.WriteLine(newPassEncrypted);
                }
                MessageBox.Show("Паролата е сменена успешно!");
            }
            else 
            {
                MessageBox.Show("Грешна парола!");
            }
            inp_1.Text = "";
            inp_2.Text = "";
            inp_3.Text = "";
        }
    }
}
