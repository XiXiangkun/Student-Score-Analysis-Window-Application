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
    public partial class 各班各科及格率 : Form
    {
        public 各班各科及格率()
        {
            InitializeComponent();
        }
        public 各班各科及格率(object[,] y_pass_per, object[,] s_pass_per, object[,] e_pass_per)
        {
            InitializeComponent();
            listView1.Clear();
            listView1.GridLines = true;//表格是否显示网格线
            listView1.FullRowSelect = true;//是否选中整行
            listView1.View = View.Details;//设置显示方式
            listView1.Scrollable = true;//是否自动显示滚动条
            listView1.MultiSelect = false;//是否可以选择多行
            listView1.Columns.Add("科目", 50, System.Windows.Forms.HorizontalAlignment.Center);
            for (int i = 0; i < 8; i++)
            {
                listView1.Columns.Add((i + 1).ToString() + "班", 60, System.Windows.Forms.HorizontalAlignment.Center);
            }
            ListViewItem item = new ListViewItem();
            item.SubItems.Clear();
            //将名字添加进去
            item.SubItems[0].Text = "语文";
            for(int i = 0; i < 8; i++)
            {
                item.SubItems.Add(y_pass_per[i, 1].ToString());
            }
            listView1.Items.Add(item);
            ListViewItem item_2 = new ListViewItem();
            item_2.SubItems.Clear();
            //将名字添加进去
            item_2.SubItems[0].Text = "数学";
            for (int i = 0; i < 8; i++)
            {
                item_2.SubItems.Add(s_pass_per[i, 1].ToString());
            }
            listView1.Items.Add(item_2);
            ListViewItem item_3 = new ListViewItem();
            item_3.SubItems.Clear();
            //将名字添加进去
            item_3.SubItems[0].Text = "英语";
            for (int i = 0; i < 8; i++)
            {
                item_3.SubItems.Add(e_pass_per[i, 1].ToString());
            }
            listView1.Items.Add(item_3);

        }
    }
}
