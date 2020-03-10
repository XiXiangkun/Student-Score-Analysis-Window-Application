using System;
using System.IO;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Vevisoft.Excel.Core;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.Collections;

namespace WindowsFormsApplication1
{
    public partial class Form1 : Form
    {
        public string[] names;
        public string[] subjects;
        public object[,] tables;
        public object[,] total;
        public int[] stu_num = new int[8];
        public int[] gra_total = new int[8];
        public object[,] gra_ave;
        public int[] y_pass_num = new int[8];
        public int[] s_pass_num = new int[8];
        public int[] e_pass_num = new int[8];
        public object[,] y_pass_per;
        public object[,] s_pass_per;
        public object[,] e_pass_per;

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (total != null)
            {
                //显示学生成绩
                Order.Orderby(total, new int[] { 2 }, 1);
                学生成绩 newForm = new 学生成绩(total);
                newForm.Show();
            }
            else
            {
                MessageBox.Show("请完成步骤①②，导入Excel数据并计算！");
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            IWorkbook workbook = null;  //新建IWorkbook对象
            string fileName = "..\\..\\..\\..\\grades.xlsx";
            FileStream fileStream = new FileStream(@"..\\..\\..\\..\\grades.xlsx", FileMode.Open, FileAccess.Read);
            if (fileName.IndexOf(".xlsx") > 0) // 2007版本
            {
                workbook = new XSSFWorkbook(fileStream);  //xlsx数据读入workbook
            }
            else if (fileName.IndexOf(".xls") > 0) // 2003版本
            {
                workbook = new HSSFWorkbook(fileStream);  //xls数据读入workbook
            }
            ISheet sheet = workbook.GetSheetAt(0);  //获取第一个工作表
            IRow row;// = sheet.GetRow(0);            //新建当前工作表行数据

            //保存成绩
            row = sheet.GetRow(0);
            names = new string[sheet.LastRowNum + 1];
            subjects = new string[row.LastCellNum];
            tables=new object[sheet.LastRowNum, row.LastCellNum-1];

            for (int i = 0; i < sheet.LastRowNum + 1; i++)  //对工作表每一行
            {
                row = sheet.GetRow(i);   //row读入第i行数据
                if (row != null)
                {
                    for (int j = 0; j < row.LastCellNum; j++)  //对工作表每一列
                    {
                        string cellValue = row.GetCell(j).ToString(); //获取i行j列数据
                        if (i == 0)
                        {
                            subjects[j] = cellValue;
                        }
                        else
                        {
                            if (j == 0)
                            {
                                names[i] = cellValue;
                            }
                            else
                            {
                                tables[i-1, j-1] = int.Parse(cellValue);
                            }
                        }
                    }
                }
            }
            fileStream.Close();
            workbook.Close();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (gra_ave != null)
            {
                //显示班级总平均成绩
                Order.Orderby(gra_ave, new int[] { 1 }, 1);
                班级平均分 newForm = new 班级平均分(gra_ave);
                newForm.Show();
            }
            else
            {
                MessageBox.Show("请完成步骤①②，导入Excel数据并计算！");
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {

        }

        private void button5_Click(object sender, EventArgs e)
        {
            if (gra_ave != null)
            {
                //显示班级总平均成绩
                各班各科及格率 newForm = new 各班各科及格率(y_pass_per,s_pass_per,e_pass_per);
                newForm.Show();
            }
            else
            {
                MessageBox.Show("请完成步骤①②，导入Excel数据并计算！");
            }
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void button6_Click(object sender, EventArgs e)
        {
            listView1.Clear();
            listView1.GridLines = true;//表格是否显示网格线
            listView1.FullRowSelect = true;//是否选中整行
            listView1.View = View.Details;//设置显示方式
            listView1.Scrollable = true;//是否自动显示滚动条
            listView1.MultiSelect = false;//是否可以选择多行
            
            //没有导入数据的情况
            if (subjects == null)
            {
                listView1.Columns.Add("没有导入数据", 150, System.Windows.Forms.HorizontalAlignment.Center);
                return;
            }

            //正常导入后
            //Order.Orderby(tables, new int[] { 1 }, 1);
            total = new object[names.Length - 1, 3];
            gra_ave = new object[8,2];
            y_pass_per = new object[8, 2];
            s_pass_per = new object[8, 2];
            e_pass_per = new object[8, 2];

            for (int i = 0; i < 8; i++)
            {
                stu_num[i] = 0;
                y_pass_num[i] = 0;
                s_pass_num[i] = 0;
                e_pass_num[i] = 0;
                gra_total[i] = 0;
                gra_ave[i, 0] = i + 1;
                y_pass_per[i,0] = i+1;
                y_pass_per[i, 1] = 0;
                s_pass_per[i, 0] = i+1;
                s_pass_per[i, 1] = 0;
                e_pass_per[i, 0] = i+1;
                e_pass_per[i, 1] = 0;
            }
            
            for (int i = 0; i < subjects.Length; i++)
            {
                listView1.Columns.Add(subjects[i], 60, System.Windows.Forms.HorizontalAlignment.Center);
            }
            for (int i = 1; i < names.Length; i++)
            {
                ListViewItem item = new ListViewItem();
                item.SubItems.Clear();
                //将名字添加进去
                item.SubItems[0].Text = names[i];
                //统计总成绩
                total[i-1,0] = names[i];
                total[i - 1, 1] = 0;
                total[i - 1, 2] = 0;


                //计算班级人数
                stu_num[int.Parse(tables[i - 1, 0].ToString()) - 1]++;
                //textBox1.Text = "?";
                //textBox1.Show();
                for (int k = 0; k < subjects.Length - 1; k++)
                {
                    item.SubItems.Add(tables[i-1,k].ToString());
                    if (k != 0)
                    {
                        if (k == 1)
                        {
                            //及格人数统计
                            if (int.Parse(tables[i - 1, k].ToString()) >= 90) y_pass_num[int.Parse(tables[i - 1, 0].ToString()) - 1]++;
                        }
                        else if (k == 2)
                        {
                            if (int.Parse(tables[i - 1, k].ToString()) >= 90) s_pass_num[int.Parse(tables[i - 1, 0].ToString()) - 1]++;
                        }
                        else if (k == 3)
                        {
                            if (int.Parse(tables[i - 1, k].ToString()) >= 90) e_pass_num[int.Parse(tables[i - 1, 0].ToString()) - 1]++;
                        }
                        
                        total[i - 1, 2] = int.Parse(total[i - 1, 2].ToString()) + int.Parse(tables[i - 1, k].ToString());
                    }
                    total[i - 1, 1] = int.Parse(tables[i - 1, 0].ToString());
                }
                listView1.Items.Add(item);

            }

            //统计班级平均分 
            for (int i=0;i< names.Length - 1; i++)
            {
                gra_total[int.Parse(total[i, 1].ToString())-1] += int.Parse(total[i, 2].ToString());
            }
            for (int i = 0; i < 8; i++)
            {
                gra_ave[i,1] = gra_total[i]/stu_num[i];
                y_pass_per[i, 1] = Math.Round((float)y_pass_num[i] / stu_num[i], 2);
                s_pass_per[i, 1] = Math.Round((float)s_pass_num[i] / stu_num[i], 2);
                e_pass_per[i, 1] = Math.Round((float)e_pass_num[i] / stu_num[i], 2);
            }
        }

        private void listView1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }
    }

    public class Order
    {
        /// <summary>
        /// 对二维数组排序
        /// </summary>
        /// <param name="values">排序的二维数组</param>
        /// <param name="orderColumnsIndexs">排序根据的列的索引号数组</param>
        /// <param name="type">排序的类型，1代表降序，0代表升序</param>
        /// <returns>返回排序后的二维数组</returns>
        public static object[,] Orderby(object[,] values, int[] orderColumnsIndexs, int type)
        {
            object[] temp = new object[values.GetLength(1)];
            int k;
            int compareResult;
            for (int i = 0; i < values.GetLength(0); i++)
            {
                for (k = i + 1; k < values.GetLength(0); k++)
                {
                    if (type.Equals(1))
                    {
                        for (int h = 0; h < orderColumnsIndexs.Length; h++)
                        {
                            compareResult = Comparer.Default.Compare(GetRowByID(values, k).GetValue(orderColumnsIndexs[h]), GetRowByID(values, i).GetValue(orderColumnsIndexs[h]));
                            if (compareResult.Equals(1))
                            {
                                temp = GetRowByID(values, i);
                                Array.Copy(values, k * values.GetLength(1), values, i * values.GetLength(1), values.GetLength(1));
                                CopyToRow(values, k, temp);
                            }
                            if (compareResult != 0)
                                break;
                        }
                    }
                    else
                    {
                        for (int h = 0; h < orderColumnsIndexs.Length; h++)
                        {
                            compareResult = Comparer.Default.Compare(GetRowByID(values, k).GetValue(orderColumnsIndexs[h]), GetRowByID(values, i).GetValue(orderColumnsIndexs[h]));
                            if (compareResult.Equals(-1))
                            {
                                temp = GetRowByID(values, i);
                                Array.Copy(values, k * values.GetLength(1), values, i * values.GetLength(1), values.GetLength(1));
                                CopyToRow(values, k, temp);
                            }
                            if (compareResult != 0)
                                break;
                        }
                    }
                }
            }
            return values;

        }
        /// <summary>
        /// 获取二维数组中一行的数据
        /// </summary>
        /// <param name="values">二维数据</param>
        /// <param name="rowID">行ID</param>
        /// <returns>返回一行的数据</returns>
        static object[] GetRowByID(object[,] values, int rowID)
        {
            if (rowID > (values.GetLength(0) - 1))
                throw new Exception("rowID超出最大的行索引号!");

            object[] row = new object[values.GetLength(1)];
            for (int i = 0; i < values.GetLength(1); i++)
            {
                row[i] = values[rowID, i];

            }
            return row;

        }
        /// <summary>
        /// 复制一行数据到二维数组指定的行上
        /// </summary>
        /// <param name="values"></param>
        /// <param name="rowID"></param>
        /// <param name="row"></param>
        static void CopyToRow(object[,] values, int rowID, object[] row)
        {
            if (rowID > (values.GetLength(0) - 1))
                throw new Exception("rowID超出最大的行索引号!");
            if (row.Length > (values.GetLength(1)))
                throw new Exception("row行数据列数超过二维数组的列数!");
            for (int i = 0; i < row.Length; i++)
            {
                values[rowID, i] = row[i];
            }
        }
    }
}
