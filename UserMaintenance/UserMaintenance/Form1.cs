using System.ComponentModel;
using System.IO;
using System.Windows.Forms;
using UserMaintenance.Entities;

namespace UserMaintenance
{
    public partial class Form1 : Form
    {
        BindingList<User> users = new BindingList<User>();

        public Form1()
        {
            InitializeComponent();
            label1.Text = Resource1.FullName;
            button1.Text = Resource1.Add;
            button2.Text = Resource1.WriteToFile;

            listBox1.DataSource = users;
            listBox1.ValueMember = "ID";
            listBox1.DisplayMember = "FullName";
        }

        private void button1_Click(object sender, System.EventArgs e)
        {
            var u = new User()
            {
                FullName = textBox1.Text
            };
            users.Add(u);
        }

        private void button2_Click(object sender, System.EventArgs e)
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "Csv files (*.csv)|*.csv";

            var path = saveFileDialog.ShowDialog();
            var fileName = saveFileDialog.FileName;

            if (!string.IsNullOrEmpty(fileName))
                UsersToFile(saveFileDialog.FileName);
        }

        private void UsersToFile(string path)
        {
            using (StreamWriter sw = new StreamWriter(path, true))
            {
                foreach (var item in users)
                {
                    sw.WriteLine(item.FullName);
                }
            }
        }
    }
}
