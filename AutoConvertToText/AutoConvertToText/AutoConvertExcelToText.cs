using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace AutoConvertToText
{
    public partial class AutoConvertExcelToText : Form
    {
        public AutoConvertExcelToText()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog dialog = new OpenFileDialog())
            {
                dialog.Multiselect = true;
                if(dialog.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        this.textBox1.Text = dialog.FileName;
                    }
                    catch(Exception ex)
                    {
                        throw (ex);
                    }
                }
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            string[] path = { textBox1.Text ,"ConvertResult","/v"};
            string[] file = textBox1.Text.Split('.');
            Console.WriteLine("当前截取的片段--->{0}", file[file.Length - 1]);
            if(textBox1.Text == null || file[file.Length-1] != "xlsx")
            {
                DialogResult dr = MessageBox.Show("请选择Excel文件", "错误！", MessageBoxButtons.OK, MessageBoxIcon.Question);
                if (dr == DialogResult.OK)
                {
                    textBox1.Text = string.Empty;
                    return;
                }
                return;
            }
            ExcelToTxt.RunConvertTool(path);
            DialogResult succeed = MessageBox.Show("Excel转换为.txt成功！" + textBox1.Text, "提示", MessageBoxButtons.OK, MessageBoxIcon.Question);
            if(succeed == DialogResult.OK)
            {
                textBox1.Text = string.Empty;
                return;
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog path = new FolderBrowserDialog();
            path.ShowDialog();
            textBox2.Text = path.SelectedPath;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            
            string[] path = { textBox2.Text, textBox2.Text, "/v" };

            if (textBox2.Text == string.Empty)
            {
                DialogResult dr = MessageBox.Show("请选择文件夹路径!", "错误！", MessageBoxButtons.OK, MessageBoxIcon.Question);
                if (dr == DialogResult.OK)
                {
                    textBox2.Text = string.Empty;
                    return;
                }
                return;
            }
            ExcelToTxt.RunConvertTool(path);
            DialogResult succeed = MessageBox.Show("Excel转换为.txt成功！" + textBox1.Text, "提示", MessageBoxButtons.OK, MessageBoxIcon.Question);
            if (succeed == DialogResult.OK)
            {
                textBox1.Text = string.Empty;
                return;
            }
        }
        
        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            string currentSelection = comboBox1.SelectedItem.ToString();
            
        }
        public static void SetComboBox(ComboBox comboBox,String []param)
        {
            if (param.Length < 1)
            {
                DialogResult dialog = MessageBox.Show("当前没有任何Office版本数据！请检查！", "错误", MessageBoxButtons.OK, MessageBoxIcon.Question);
                if (dialog == DialogResult.OK)
                {
                    return;
                }
            }           
            for (int i = 0; i < param.Length -1; i++)
            {
                comboBox.Items.Add(param[i]);
                if(param[i] == "Office2016")
                {
                    comboBox.SelectedItem = param[i];
                }
            }
        }
    }
}
