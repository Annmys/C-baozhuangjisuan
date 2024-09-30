using System;
using System.Collections.Generic;
using System.Collections.Concurrent;
using System.Linq;
using System.IO;
using System.Threading;
using Sunny.UI;
using System.Windows.Forms;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml; // EPPlus的命名空间

namespace 包装计算
{
    public partial class Form1 : UIForm
    {
        private string 订单excel地址;
        private string 附件excel地址;
        private List<string> 订单型号列表 = new List<string>();
        List<double> 灯带_米数列表 = new List<double>();
        private List<string> 公司型号列表 = new List<string> 
        { "F10","F15","F16","F21","F22","F23","F1212","F2008","F2010","F2012","F2019","F2222" };

        private int 查找组合_基数 = 1900;  // 查找组合的基数


        private List<灯带尺寸> 灯带尺寸列表 = new List<灯带尺寸>   // 灯带尺寸列表
        {
          new 灯带尺寸("F10",9,18),
          new 灯带尺寸("F15",11.5,21),
          new 灯带尺寸("F16",15.5,6),
          new 灯带尺寸("F21",11.5,29),
          new 灯带尺寸("F22",16,17),
          new 灯带尺寸("F23",10,10),
          new 灯带尺寸("F1212",12,12),
          new 灯带尺寸("F2008",20,8),
          new 灯带尺寸("F2010",20,10),
          new 灯带尺寸("F2012",20,12),
          new 灯带尺寸("F2219",22,19),
          new 灯带尺寸("F2222",22,22)

        };
        // 灯带尺寸类
        public class 灯带尺寸
        {
            public string 型号 { get; set; }
            public double 宽度 { get; set; }
            public double 高度 { get; set; }
            public double 每米面积 { get; set; }

            public 灯带尺寸(string 型号,double 宽度,double 高度)
            {
                this.型号 = 型号;
                this.宽度 = 宽度;
                this.高度 = 高度;
                this.每米面积= 宽度 * 10; //单位是CM

            }

            // 重写 ToString 方法以便打印
            public override string ToString()
            {
                return $"{型号} - 宽度:{宽度} - 高度:{高度} - 面积:{每米面积}";
            }
        }  

        public Form1()
        {
            InitializeComponent();

            //读取excel文件必须增加声明才能运行正常
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
        }

        private void button_订单导入_Click(object sender,EventArgs e)
        {
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;

            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Multiselect = false;
            dialog.Title = "请选择数据库文件";
            dialog.Filter = "excel文件(*.xlsx)|*.xlsx|All files (*.*)|*.*";
            dialog.InitialDirectory = Application.StartupPath + @"\数据库";

            if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                订单excel地址 = dialog.FileName;
                uiTextBox_订单地址.Text = dialog.FileName;
                uiTextBox_订单地址.BackColor = System.Drawing.Color.LightGreen;
                uiTextBox_状态.AppendText("订单导入" + Environment.NewLine);

                try
                {
                    int materialColumn = 寻找EXCEL表格_特定内容位置(订单excel地址,"规格型号","Sheet1");
                    if (materialColumn != -1)
                    {
                        //MessageBox.Show($"找到在第 {materialColumn} 列");
                        uiTextBox_状态.AppendText($"规格型号在第 {materialColumn} 列" + Environment.NewLine);

                        using (var package = new ExcelPackage(new FileInfo(订单excel地址)))
                        {
                            // 读取工作表
                            var worksheet = package.Workbook.Worksheets[0];

                            // 从第二行开始遍历，假设第一行是标题行
                            for (int row = 2; row <= worksheet.Dimension.End.Row; row++)
                            {
                                var cell = worksheet.Cells[row,materialColumn];
                                if (cell.Value != null)
                                {
                                    string value = cell.Text;

                                    // 按照符号“-”进行分割
                                    string[] parts = value.Split('-');
                                    // 确保有足够的分割部分
                                    if (parts.Length >= 3)
                                    {
                                        string partBetweenSecondAndThirdDashes = parts[2]; // 获取第二个“-”和第三个“-”之间的数据

                                        // 排除包含“/”的数据
                                        if (!partBetweenSecondAndThirdDashes.Contains("/"))
                                        {
                                            // 检查是否包含公司型号列表中的型号
                                            if (公司型号列表.Exists(model => partBetweenSecondAndThirdDashes.Contains(model)))
                                            {
                                                // 如果包含，则添加到订单型号列表中
                                                if (!订单型号列表.Contains(partBetweenSecondAndThirdDashes))
                                                {
                                                    //MessageBox.Show(partBetweenSecondAndThirdDashes,row.ToString());
                                                    订单型号列表.Add(partBetweenSecondAndThirdDashes);
                                                }
                                            }
                                        }
                                    }
                                }
                            }

                            // 显示订单型号列表
                            uiTextBox_状态.AppendText("订单型号列表：" + Environment.NewLine);
                            foreach (var model in 订单型号列表)
                            {
                                uiTextBox_状态.AppendText(model + Environment.NewLine);
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"发生错误：{ex.Message}","错误",MessageBoxButtons.OK,MessageBoxIcon.Error);
                }
            }
        }

        private int 寻找EXCEL表格_特定内容位置(string excel地址,string 寻找字符,string sheet名字)
        {
            int materialColumn = -1;
            string shuchu = "";

            using (var package = new ExcelPackage(new FileInfo(excel地址)))
            {
                var worksheet = package.Workbook.Worksheets[sheet名字];

                for (int col = 1; col <= worksheet.Dimension.End.Column; col++)
                {
                    var cell = worksheet.Cells[1,col];
                    if (cell.Value != null && cell.Value.ToString().Contains(寻找字符))
                    {
                        materialColumn = col;
                        shuchu = "A" + materialColumn;
                        break;
                    }
                }
            }

            return materialColumn;
        }

        private void button_附件导入_Click_1(object sender,EventArgs e)
        {
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;

            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Multiselect = false;
            dialog.Title = "请选择数据库文件";
            dialog.Filter = "excel文件(*.xlsx)|*.xlsx|All files (*.*)|*.*";
            dialog.InitialDirectory = Application.StartupPath + @"\数据库";
            uiTextBox_状态.AppendText("附件导入" + Environment.NewLine);

            if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                附件excel地址 = dialog.FileName;
                uiTextBox_附件地址.Text = dialog.FileName;
                uiTextBox_附件地址.BackColor = System.Drawing.Color.LightGreen;
            }
        }

        private void button_开始处理_Click(object sender,EventArgs e)
        {
            开始组合();

            foreach (var 尺寸 in 灯带尺寸列表)
            {
                if (尺寸.型号 == "F22")
                {
                    //MessageBox.Show($"型号: {尺寸.型号}, 宽度: {尺寸.宽度}, 高度: {尺寸.高度}, 面积: {尺寸.每米面积}");
                    //MessageBox.Show($"型号: {尺寸.型号},  面积: {尺寸.每米面积}");


                    


                    break; // 如果找到型号为"F22"的尺寸，输出宽度并退出循环
                }
            }

        }

        private void 开始组合()
        {
            查找组合_基数 = 8;

            //灯带_米数列表.Add(31.3);
            //灯带_米数列表.Add(85.4);
            //灯带_米数列表.Add(31.3);
            //灯带_米数列表.Add(85.4);
            //灯带_米数列表.Add(31.1);
            //灯带_米数列表.Add(85.2);
            //灯带_米数列表.Add(31.1);
            //灯带_米数列表.Add(85.2);
            //灯带_米数列表.Add(91.3);
            //灯带_米数列表.Add(131.4);
            //灯带_米数列表.Add(260.5);
            灯带_米数列表.Add(29.6);

            灯带_米数列表.Add(6.6);

            灯带_米数列表.Add(1.1);
            灯带_米数列表.Add(2.2);
            灯带_米数列表.Add(3.3);
            灯带_米数列表.Add(4.4);
            灯带_米数列表.Add(5.5);


            HashSet<string> 结果_列表 = new HashSet<string>(); // 使用HashSet来自动去重
            int 包装有效面积 = 查找组合_基数;

            // 遍历每个数字，从每个数字开始寻找组合
            foreach (double num in 灯带_米数列表)
            {
                List<double> 临时_列表 = new List<double> { num };  // 初始化临时列表
                查找组合(灯带_米数列表,临时_列表,包装有效面积 - num,0,结果_列表);
            }

            StringBuilder message = new StringBuilder();
            foreach (var result in 结果_列表)
            {
                message.AppendLine(result + " < " + 查找组合_基数.ToString());
            }

            if (message.Length > 0)
            {
                MessageBox.Show(message.ToString(),"组合结果",MessageBoxButtons.OK,MessageBoxIcon.Information);
            }
            else
            {
                MessageBox.Show("没有找到任何组合。","无结果",MessageBoxButtons.OK,MessageBoxIcon.Warning);
            }
        }

        private void 查找组合(List<double> 灯带_米数列表,List<double> 单前_列表,double 剩下的,int start,HashSet<string> 结果_列表)
        {
            // 排序并去重
            单前_列表 = 单前_列表.OrderBy(x => x).ToList();
            if (单前_列表.Sum() < 查找组合_基数 && 单前_列表.Count > 1)
            {
                string result = string.Join(" + ",单前_列表);
                结果_列表.Add(result);
            }

            for (int i = start; i < 灯带_米数列表.Count; i++)
            {
                double num = 灯带_米数列表[i];
                if (!单前_列表.Contains(num) && 单前_列表.Sum() + num < 查找组合_基数)
                {
                    单前_列表.Add(num);
                    查找组合(灯带_米数列表,单前_列表,剩下的 - num,i + 1,结果_列表);
                    单前_列表.RemoveAt(单前_列表.Count - 1);
                }
            }
        }

    }
}