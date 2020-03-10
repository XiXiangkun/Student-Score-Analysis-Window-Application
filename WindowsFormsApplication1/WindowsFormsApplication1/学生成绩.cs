using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WindowsFormsApplication1
{
    public partial class 学生成绩 : Form
    {
        int[] total;
        public 学生成绩()
        {
            InitializeComponent();
        }
        public 学生成绩(object[,] total)
        {
            InitializeComponent();
            listView1.Clear();
            listView1.GridLines = true;//表格是否显示网格线
            listView1.FullRowSelect = true;//是否选中整行
            listView1.View = View.Details;//设置显示方式
            listView1.Scrollable = true;//是否自动显示滚动条
            listView1.MultiSelect = false;//是否可以选择多行

            listView1.Columns.Add("学号", 50, System.Windows.Forms.HorizontalAlignment.Center);
            listView1.Columns.Add("班级", 50, System.Windows.Forms.HorizontalAlignment.Center);
            listView1.Columns.Add("总成绩", 50, System.Windows.Forms.HorizontalAlignment.Center);
            listView1.Columns.Add("排名", 50, System.Windows.Forms.HorizontalAlignment.Center);

            for (int i = 0; i < 10; i++)
            {
                ListViewItem item = new ListViewItem();
                item.SubItems.Clear();
                //将名字添加进去
                item.SubItems[0].Text = total[i,0].ToString();
                item.SubItems.Add(total[i, 1].ToString());
                item.SubItems.Add(total[i, 2].ToString());
                item.SubItems.Add((i+1).ToString());
                listView1.Items.Add(item);
            }
        }
    }
}
