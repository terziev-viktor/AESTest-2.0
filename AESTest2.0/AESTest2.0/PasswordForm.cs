using AESTest2._0.Extensions;
using System;
using System.IO;
using System.Security.AccessControl;
using System.Windows.Forms;

namespace AESTest2._0
{
    public partial class PasswordForm : Form
    {
        private const string PATH = @"C:\Windows\iluesjkdgbk.txt";
        private const string DEFAULTPASS = "123";
        public PasswordForm()
        {
            InitializeComponent();
            if (!File.Exists(PATH))
            {
                MessageBox.Show("Паролата ви е по подразбиране. Силно ви съветваме да я смените с нова!");

                using (StreamWriter writer = File.CreateText(PATH))
                {
                    string encrypted = Protection.Crypt(DEFAULTPASS);
                    writer.WriteLine(encrypted);
                }
            }
        }

        private void btnExit_Click(object sender, EventArgs e)
        {
            using(StreamReader reader = new StreamReader(PATH))
            {
                string decrypted = reader.ReadLine();
                string encrypted = Protection.Crypt(inp.Text);
                if(decrypted == encrypted)
                {
                    ExplorerManager.Start();
                    Application.Exit();
                }
                else
                {
                    MessageBox.Show("Грешна парола!");
                    this.Close();
                }
            }
        }

        public static bool GetPassword()
        {
            string inp = Prompt.ShowDialog("Поле за парола:", "За да влезете в мениджърския прозорец въведете парола!");
            bool rightPass = false;
            using (var r = new StreamReader(PATH))
            {
                string cryptedPass = r.ReadLine();
                string cryptedInp = Protection.Crypt(inp);
                rightPass = cryptedPass == cryptedInp;
            }
            return rightPass;
        }

        private void btnChangePass_Click(object sender, EventArgs e)
        {
            string pass;
            using(StreamReader reader = new StreamReader(PATH))
            {
                pass = reader.ReadLine();
            }

            string encrypted = Protection.Crypt(inp_1.Text);

            if (inp_2.Text == inp_3.Text && pass == encrypted)
            {
                string newPassEncrypted = Protection.Crypt(inp_2.Text);
                using(StreamWriter w = new StreamWriter(PATH))
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
