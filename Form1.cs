using System;
using System.Collections.Generic;
using System.Collections.Concurrent;
using System.Drawing;
using System.Linq;
using System.IO;
using System.Threading;
using Sunny.UI;
using System.Windows.Forms;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml; // EPPlus的命名空间.

namespace 包装计算
{
    public partial class Form1 : UIForm
    {
        private 变量 变量 = new 变量();

        public Form1()
        {
            InitializeComponent();

            //读取excel文件必须增加声明才能运行正常
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
        }

        private void button_订单导入_Click(object sender, EventArgs e)
        {
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;

            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Multiselect = false;
            dialog.Title = "请选择数据库文件";
            dialog.Filter = "excel文件(*.xlsx)|*.xlsx|All files (*.*)|*.*";
            dialog.InitialDirectory = Application.StartupPath + @"\数据库";

            if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                变量.订单excel地址 = dialog.FileName;
                uiTextBox_订单地址.Text = dialog.FileName;
                uiTextBox_订单地址.BackColor = System.Drawing.Color.LightGreen;

                EXCEL订单数据_转列表();

                uiTextBox_状态.AppendText("订单导入" + Environment.NewLine);
            }
        }

        private void EXCEL订单数据_转列表()
        {
            try
            {
                int materialColumn = 寻找EXCEL表格_特定内容位置(变量.订单excel地址, "规格型号", "Sheet1");
                if (materialColumn != -1)
                {
                    //MessageBox.Show($"找到在第 {materialColumn} 列");
                    uiTextBox_状态.AppendText($"规格型号在第 {materialColumn} 列" + Environment.NewLine);

                    using (var package = new ExcelPackage(new FileInfo(变量.订单excel地址)))
                    {
                        // 读取工作表
                        var worksheet = package.Workbook.Worksheets[0];

                        // 从第二行开始遍历，假设第一行是标题行
                        for (int row = 2; row <= worksheet.Dimension.End.Row; row++)
                        {
                            var cell = worksheet.Cells[row, materialColumn];
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
                                        if (变量.公司型号列表.Exists(model => partBetweenSecondAndThirdDashes.Contains(model)))
                                        {
                                            // 如果包含，则添加到订单型号列表中
                                            if (!变量.订单型号列表.Contains(partBetweenSecondAndThirdDashes))
                                            {
                                                //MessageBox.Show(partBetweenSecondAndThirdDashes,row.ToString());
                                                变量.订单型号列表.Add(partBetweenSecondAndThirdDashes);
                                            }
                                        }
                                    }
                                }
                            }
                        }

                        // 显示订单型号列表
                        uiTextBox_状态.AppendText("订单型号列表：" + Environment.NewLine);
                        foreach (var model in 变量.订单型号列表)
                        {
                            uiTextBox_状态.AppendText(model + Environment.NewLine);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"发生错误：{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private int 寻找EXCEL表格_特定内容位置(string excel地址, string 寻找字符, string sheet名字)
        {
            int materialColumn = -1;
            string shuchu = "";

            using (var package = new ExcelPackage(new FileInfo(excel地址)))
            {
                var worksheet = package.Workbook.Worksheets[sheet名字];

                for (int col = 1; col <= worksheet.Dimension.End.Column; col++)
                {
                    var cell = worksheet.Cells[1, col];
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

        private void button_附件导入_Click_1(object sender, EventArgs e)
        {
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;

            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Multiselect = false;
            dialog.Title = "请选择数据库文件";
            dialog.Filter = "Excel 文件(*.xlsx)|*.xlsx|所有文件 (*.*)|*.*";
            dialog.InitialDirectory = Application.StartupPath + @"\数据库";

            if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                变量.附件excel地址 = dialog.FileName;
                uiTextBox_附件地址.Text = dialog.FileName;
                uiTextBox_附件地址.BackColor = System.Drawing.Color.LightGreen;

                EXCEL附件数据_转列表();

                uiTextBox_状态.AppendText("附件导入" + Environment.NewLine);

                //寻找检查字典中对应型号SHEET数据_测试代码();
            }
        }

        private void 寻找检查字典中对应型号SHEET数据_测试代码()
        {
            // 获取用户指定的型号SHEET
            string 型号SHEET = "TLX8 naked"; // 这里替换为实际的型号SHEET名称
            string 型号SHEET1 = "TLX8 SC HB"; // 这里替换为实际的型号SHEET名称

            // 检查字典中是否有对应的型号SHEET数据
            if (变量.附件表数据.ContainsKey(型号SHEET))
            {
                string 提示信息 = $"工作表: {型号SHEET}\r\n";
                bool foundC1 = false;
                foreach (var item in 变量.附件表数据[型号SHEET])
                {
                    string[] 数据 = item.Split(',');

                    // 假设内容A在第一个位置，内容R在第三个位置
                    if (数据.Length >= 3 && 数据[0].Trim() == "C1")
                    {
                        foundC1 = true;
                        string 内容R = 数据[2].Trim(); // 提取内容R的值
                        提示信息 += $"内容A为C1的项，内容R的值为: {内容R}\r\n";
                    }
                    else
                    {
                        提示信息 += item + "\r\n";
                    }
                }

                if (!foundC1)
                {
                    提示信息 += "没有找到内容A为C1的项。";
                }

                MessageBox.Show(提示信息, "总长度数据", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                MessageBox.Show("没有找到指定型号的工作表数据。", "总长度数据", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }

            // 检查字典中是否有对应的型号SHEET1数据
            if (变量.附件表数据.ContainsKey(型号SHEET1))
            {
                string 提示信息1 = $"工作表: {型号SHEET1}\r\n";
                foreach (var item1 in 变量.附件表数据[型号SHEET1])
                {
                    提示信息1 += item1 + "\r\n";
                }
                MessageBox.Show(提示信息1, "总长度数据", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                MessageBox.Show("没有找到指定型号的工作表数据。", "总长度数据", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void EXCEL附件数据_转列表()
        {
            try
            {
                using (ExcelPackage package = new ExcelPackage(new FileInfo(变量.附件excel地址)))
                {
                    foreach (ExcelWorksheet worksheet in package.Workbook.Worksheets)
                    {
                        if (!变量.附件表数据.ContainsKey(worksheet.Name))
                        {
                            变量.附件表数据[worksheet.Name] = new List<string>();
                        }

                        //List<string> 当前工作表数据 = new List<string>();
                        int 总长度米行号 = -1;

                        for (int i = 1; i <= worksheet.Dimension.End.Row; i++)
                        {
                            var cellValue = worksheet.Cells[i, 18].Value;
                            if (cellValue != null && cellValue.ToString().Contains("总长度 (米)"))
                            {
                                总长度米行号 = i;
                                break;
                            }
                        }

                        if (总长度米行号 >= 0)
                        {
                            for (int row = 总长度米行号 + 1; row <= worksheet.Dimension.End.Row; row++)
                            {
                                var cellA = worksheet.Cells[row, 1];
                                var cellO = worksheet.Cells[row, 15];
                                var cellR = worksheet.Cells[row, 18];

                                if (cellA.Value != null && cellO.Value != null && cellR.Value != null)
                                {
                                    string 内容A = cellA.Value.ToString();
                                    int 内容O = Convert.ToInt32(cellO.Value);
                                    double 内容R = Convert.ToDouble(cellR.Value);

                                    if (内容O > 1)
                                    {
                                        double 分割后的R值 = 内容R / 内容O;
                                        for (int i = 0; i < 内容O; i++)
                                        {
                                            //当前工作表数据.Add($"{内容A}, {1}, {分割后的R值}");
                                            // 添加到新的列表中
                                            List<string> 新数据列表 = new List<string> { $"{内容A}, {1}, {分割后的R值}" };
                                            变量.附件表数据[worksheet.Name].AddRange(新数据列表);
                                        }
                                    }
                                    else
                                    {
                                        变量.附件表数据[worksheet.Name].Add($"{内容A}, {内容O}, {内容R}");
                                    }
                                }
                            }
                        }

                        //if (当前工作表数据.Count > 0)
                        //{
                        //    变量.附件表数据[worksheet.Name] = 当前工作表数据;
                        //}
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("读取Excel文件时出错：" + ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void button_开始处理_Click(object sender, EventArgs e)
        {
            string 型号SHEET = "TLX8 naked";
            开始组合(型号SHEET);

            foreach (var 灯带 in 变量.灯带尺寸列表)
            {
                if (灯带.型号 == "F22")
                {
                    MessageBox.Show($"型号: {灯带.型号}, 宽度: {灯带.宽度}, 高度: {灯带.高度}, 面积: {灯带.每米面积}");
                    //MessageBox.Show($"型号: {灯带.型号},  面积: {灯带.每米面积}");

                    break; // 如果找到型号为"F22"的灯带，输出宽度并退出循环
                }
            }
        }

        private void 开始组合(string 型号SHEET)
        {
            变量.查找组合_基数 = 10;

            变量.测试.Clear();  // 清空之前的列表，如果有的话

            List<数据项> 数据项列表 = new List<数据项>();

            //string 型号SHEET = "TLX8 naked";

            if (变量.附件表数据.ContainsKey(型号SHEET))
            {
                string 提示信息 = $"工作表: {型号SHEET}\r\n";
                foreach (var item in 变量.附件表数据[型号SHEET])
                {
                    string[] 数据 = item.Split(',');
                    double 内容R;
                    double.TryParse(数据[2].Trim(), out 内容R);
                    string 内容A = 数据[0].Trim(); // 内容A在数组的第1个位置
                    string 内容O = 数据[1].Trim(); // 内容O在数组的第2个位置
                    string 标志 = Guid.NewGuid().ToString(); // 生成唯一标志
                    数据项列表.Add(new 数据项(内容A, 内容O, 内容R, 标志));
                    变量.测试.Add(内容R);

                    //MessageBox.Show(数据[0].Trim()+数据[2].Trim(), "总长度数据", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }

            // 调用算法
            Solution s = new Solution();
            var ans = s.CalculateCombinations(变量.测试, 变量.查找组合_基数);


            保存组合结果到Excel(ans, 数据项列表, 型号SHEET);

            string 输出信息 = "找到的组合：" + Environment.NewLine;
            foreach (var combination in ans)
            {
                List<数据项> 临时数据项列表 = new List<数据项>(数据项列表); // 每次处理前重建数据项列表
                输出信息 += string.Join(" + ", combination.Select(r =>
                {
                    var 项 = 临时数据项列表.FirstOrDefault(d => d.内容R == r);
                    if (项 != null)
                    {
                        临时数据项列表.Remove(项); // 从临时列表中移除这一项
                        return $"{r} ({项.内容A}, {项.内容O})";
                    }
                    return $"{r}";
                })) + " < " + 变量.查找组合_基数 + Environment.NewLine;
            }

            MessageBox.Show(输出信息, "组合结果", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void 保存组合结果到Excel(List<List<double>> 组合结果, List<数据项> 数据项原列表, string 文件名)
        {
            string 文件路径 = "输出结果\\" + 文件名 + ".xlsx";
            FileInfo 文件信息 = new FileInfo(文件路径);

            using (ExcelPackage excelPackage = new ExcelPackage(文件信息))
            {
                for (int i = 0; i < 组合结果.Count; i++)
                {
                    var combination = 组合结果[i];
                    List<数据项> 数据项列表 = new List<数据项>(数据项原列表); // 每次处理前重建数据项列表
                    ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add($"第 {i + 1}盒");
                    worksheet.Cells[1, 1].Value = "内容R";
                    worksheet.Cells[1, 2].Value = "内容A";
                    worksheet.Cells[1, 3].Value = "内容O";

                    for (int j = 0; j < combination.Count; j++)
                    {
                        double 内容R = combination[j];
                        var 项 = 数据项列表.FirstOrDefault(d => d.内容R == 内容R);
                        worksheet.Cells[j + 2, 1].Value = 内容R;
                        worksheet.Cells[j + 2, 2].Value = 项 != null ? 项.内容A : "";
                        worksheet.Cells[j + 2, 3].Value = 项 != null ? 项.内容O : "";
                        if (项 != null)
                        {
                            数据项列表.Remove(项); // 从列表中移除这一项
                        }
                    }
                }

                excelPackage.Save();
            }

            MessageBox.Show($"结果已保存到 {文件路径}", "保存成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void uiButton1_Click(object sender, EventArgs e)
        {
            新包装 新包装实例 = new 新包装();

            // 示例查找
            var 包装资料 = 新包装实例.查找包装资料("F10", "双层注塑底部出线");

            if (包装资料 != null)
            {
                string 输出信息 = $"使用包装: {包装资料.半成品BOM物料码}\n总有效容积 = {包装资料.总有效容积}";
                MessageBox.Show(输出信息, "包装资料", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                MessageBox.Show("未找到匹配的包装资料", "查找结果", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }
    }
}