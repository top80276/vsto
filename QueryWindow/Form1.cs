using DB;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace QueryWindow
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
            this.Close();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            DB.sp_classes_list.Row[] row1 = DB.sp_classes_list.ExecuteArr();
            foreach(var r in row1)
            {
                string item_name = r.class_id + "-" + r.class_name;
                comboBox1.Items.Add(item_name);
                
            }
            var properties = typeof(sp_student_list.Row).GetProperties();
            string[] title_list = properties.Select(p => p.Name).ToArray();
            foreach (var title in title_list) 
            {
                listBox1.Items.Add(title);
            }
        }

        public int SelectedClassId { get; private set; } = -1;
        public string[] column { get; private set; }


        private void btnQuery_Click(object sender, EventArgs e)
        {
            if (comboBox1.SelectedIndex < 0)
            {
                MessageBox.Show("一定要選班級");
                comboBox1.Focus();
            }
            else
            {
                if (listBox2.Items.Count == 0)
                {
                    MessageBox.Show("至少要一個欄位");
                }
                else
                {
                    string item = comboBox1.SelectedItem.ToString();
                    string[] arr = item.Split('-');
                    if (arr.Length == 2)
                    {
                        SelectedClassId = int.Parse(arr[0]);
                        column = listBox2.Items.Cast<string>().ToArray();
                    }
                    else
                    {
                        throw new InvalidProgramException("???");
                    }
                    this.DialogResult = DialogResult.OK;
                    this.Close();
                }
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (listBox1.SelectedItem != null)
            {
                listBox2.Items.Add(listBox1.SelectedItem.ToString());
                listBox1.Items.Remove(listBox1.SelectedItem.ToString());
            }
            
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if(listBox2.SelectedItem != null)
            {
                listBox1.Items.Add(listBox2.SelectedItem.ToString());
                listBox2.Items.Remove(listBox2.SelectedItem.ToString());
            }
            
        }
    }
}
