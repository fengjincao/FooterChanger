using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;


namespace FooterChanger
{
    public partial class Form1 : Form
    {
       private String filePath = "";
        private string dirPath = "";
        private DirectoryInfo input_dir;
        private bool hasPath = false;
        private bool hasNewFooter = false;

        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            

            openFileDialog1.Filter = "Word files(2007-2016)|*.docx";
            openFileDialog1.Title = "Select new file";

            // Show the Dialog.  
            // If the user clicked OK in the dialog and  
            // a .CUR file was selected, open it.  
            if (openFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                // Assign the cursor in the Stream to the Form's Cursor property.  

                filePath = openFileDialog1.FileName;
                FileInfo input_file = new FileInfo(filePath);
                input_dir = new DirectoryInfo(input_file.DirectoryName);
                dirPath = input_dir.FullName;
                textBox1.Text = dirPath;
                hasPath = true;
            }

        }

        private void button2_Click(object sender, EventArgs e)
        {
            
            if (hasPath && textBox2.Text.Count()>0)
            {
                foreach (var item in input_dir.GetFiles())
                {
                    if (!item.Name.Contains("~") && item.Name.EndsWith(".docx"))
                    {
                        WordOperator.AlertWordFooter(item.FullName, textBox2.Text);
                    }
                }
            }
        }
    }
}
