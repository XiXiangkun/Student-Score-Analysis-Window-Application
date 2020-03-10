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
    public partial class 班级平均分 : Form
    {
        public 班级平均分()
        {
            InitializeComponent();
        }
        public 班级平均分(object[,] gra_ave)
        {
            InitializeComponent();
            listView1.Clear();
            listView1.GridLines = true;//表格是否显示网格线
            listView1.FullRowSelect = true;//是否选中整行
            listView1.View = View.Details;//设置显示方式
            listView1.Scrollable = true;//是否自动显示滚动条
            listView1.MultiSelect = false;//是否可以选择多行

            listView1.Columns.Add("班级", 50, System.Windows.Forms.HorizontalAlignment.Center);
            listView1.Columns.Add("总平均成绩", 150, System.Windows.Forms.HorizontalAlignment.Center);

            for (int i = 0; i < 3; i++)
            {
                ListViewItem item = new ListViewItem();
                item.SubItems.Clear();
                //将名字添加进去
                item.SubItems[0].Text = gra_ave[i, 0].ToString();
                item.SubItems.Add(gra_ave[i, 1].ToString());
                listView1.Items.Add(item);
            }
        }
    }
}
