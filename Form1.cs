using OfficeOpenXml; // EPPlus的命名空间.
using Sunny.UI;
using System.Text;

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

                EXCEL订单数据_转列表(变量.订单excel地址);

                uiTextBox_状态.AppendText("订单导入" + Environment.NewLine);
            }
        }

        public void EXCEL订单数据_转列表(string excel文件路径)
        {

            try
            {
                using (var package = new ExcelPackage(new FileInfo(excel文件路径)))
                {
                    var worksheet = package.Workbook.Worksheets[0];
                    变量.订单出线字典 = new Dictionary<string, List<(string 型号, HashSet<string> 出线方式, string F列内容)>>();

                    string 订单编号 = worksheet.Cells["A2"].Text;
                    变量.订单出线字典[订单编号] = new List<(string, HashSet<string>, string)>();

                    string 当前型号 = "";
                    var 当前出线方式 = new HashSet<string>();
                    string 当前F列内容 = "";

                    for (int row = 2; row <= worksheet.Dimension.End.Row; row++)
                    {
                        var specCell = worksheet.Cells[row, 3]; // 规格型号列
                        var fCell = worksheet.Cells[row, 6];    // F列

                        if (specCell.Value != null)
                        {
                            string 规格型号 = specCell.Text;

                            if (规格型号.StartsWith("C-"))
                            {
                                // 保存之前的型号信息
                                if (!string.IsNullOrEmpty(当前型号))
                                {
                                    变量.订单出线字典[订单编号].Add((当前型号, new HashSet<string>(当前出线方式), 当前F列内容));
                                    当前出线方式.Clear();
                                }

                                // 提取新的型号和出线方式
                                var parts = 规格型号.Split('-');
                                当前型号 = parts.Length >= 3 ? parts[2] : "";
                                当前出线方式 = new HashSet<string>(Extract出线方式(规格型号, fCell.Text));
                                
                                当前F列内容 = fCell.Text;
                            }
                            else
                            {
                                // 为非 C- 开头的行添加出线方式
                                var 额外出线方式 = Extract出线方式(规格型号);
                                foreach (var 方式 in 额外出线方式)
                                {
                                    当前出线方式.Add(方式);
                                    
                                }
                            }
                        }
                    }

                    // 保存最后一个型号的信息
                    if (!string.IsNullOrEmpty(当前型号))
                    {
                        变量.订单出线字典[订单编号].Add((当前型号, new HashSet<string>(当前出线方式), 当前F列内容));
                    }

                    // 输出订单信息
                    输出订单信息(订单编号, 变量.订单出线字典[订单编号]);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"发生错误：{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void 输出订单信息(string 订单编号, List<(string 型号, HashSet<string> 出线方式, string F列内容)> 型号出线方式列表)
        {
            uiTextBox_状态.AppendText($"订单编号: {订单编号}" + Environment.NewLine);
            foreach (var (型号, 出线方式, F列内容) in 型号出线方式列表)
            {
                string 出线方式字符串 = 出线方式.Count > 0 ? string.Join("，", 出线方式) : "无";
                uiTextBox_状态.AppendText($"当前型号: {型号}" + Environment.NewLine);
                uiTextBox_状态.AppendText($"当前出线方式: {出线方式字符串}" + Environment.NewLine);
                uiTextBox_状态.AppendText($"F列内容: {F列内容}" + Environment.NewLine);
                uiTextBox_状态.AppendText(Environment.NewLine); // 添加一个空行，使输出更易读
            }
        }



        private List<string> Extract出线方式(string 规格型号, string F列内容 = "")  // 添加F列内容参数
        {
            var 出线方式列表 = new List<string>();
            var 需要判断弯型的型号 = new List<string> { "F22", "F23", "F2222", "F2219" };

            var 出线方式对照表 = new Dictionary<string, List<string>>
    {
        { "端部出线", new List<string> { "端部出线", "端部" } },
        { "侧面出线", new List<string> { "侧面出线", "侧部出线", "侧部", "侧面" } },
        { "底部出线", new List<string> { "底部出线", "底部" } }
    };

            // 提取型号名称
            string 基础型号 = "";
            if (规格型号.StartsWith("C-"))
            {
                var parts = 规格型号.Split('-');
                if (parts.Length >= 3)
                {
                    基础型号 = parts[2];
                }
            }

            // 判断是否是特殊型号
            bool 是特殊型号 = 需要判断弯型的型号.Any(型号 => 基础型号.StartsWith(型号));

            if (规格型号.StartsWith("C-") && 是特殊型号)
            {
                bool 是正弯 = 规格型号.Contains("正弯");
                string 弯型前缀 = 是正弯 ? "正弯" : "侧弯";

                // 对于F22型号，默认添加所有出线方式
                if (基础型号.StartsWith("F22"))
                {
                    出线方式列表.Add($"{弯型前缀}底部出线");
                    出线方式列表.Add($"{弯型前缀}端部出线");
                    出线方式列表.Add($"{弯型前缀}侧面出线");
                }
                // 对于F23B型号，根据F列内容判断
                else if (基础型号.StartsWith("F23"))
                {
                    if (!F列内容.Contains("naked"))
                    {
                        出线方式列表.Add($"{弯型前缀}端部出线");
                        出线方式列表.Add($"{弯型前缀}侧面出线");
                    }
                }
                // 其他特殊型号的处理
                else
                {
                    foreach (var 出线方式 in 出线方式对照表)
                    {
                        if (出线方式.Value.Any(变体 => 规格型号.Contains(变体)))
                        {
                            出线方式列表.Add($"{弯型前缀}{出线方式.Key}");
                        }
                    }
                }

                //MessageBox.Show($"规格型号: {规格型号}\n基础型号: {基础型号}\n是正弯: {是正弯}\n弯型前缀: {弯型前缀}\nF列内容: {F列内容}\n出线方式: {string.Join(", ", 出线方式列表)}");
            }
            else if (规格型号.StartsWith("C-"))
            {
                foreach (var 出线方式 in 出线方式对照表)
                {
                    if (出线方式.Value.Any(变体 => 规格型号.Contains(变体)))
                    {
                        出线方式列表.Add(出线方式.Key);
                    }
                }
            }

            return 出线方式列表;
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
            string 产品型号 = "F23";  // 根据实际情况设置对应的产品型号

            // 查找对应型号的灯带尺寸
            var 灯带尺寸 = 变量.灯带尺寸列表.FirstOrDefault(x => x.型号 == 产品型号);
            if (灯带尺寸 == null)
            {
                MessageBox.Show($"未找到型号 {产品型号} 的尺寸数据", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // 设置默认的包装容积（如果需要的话）
            变量.查找组合_基数 = 1420*0.6;  // 这里可以根据实际情况设置默认值

            // 调用修改后的开始组合方法
            开始组合(型号SHEET, 产品型号);

            // 显示选中灯带的信息
            MessageBox.Show(
                $"型号: {灯带尺寸.型号}\n" +
                $"宽度: {灯带尺寸.横截面宽度}\n" +
                $"高度: {灯带尺寸.横截面高度}\n" +
                $"每厘米面积: {灯带尺寸.每厘米面积}",
                "灯带信息",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information
            );
        }

        private void 开始组合(string 型号SHEET, string 产品型号)
        {
            // 查找对应型号的灯带尺寸
            var 灯带尺寸 = 变量.灯带尺寸列表.FirstOrDefault(x => x.型号 == 产品型号.Replace("B", ""));
            if (灯带尺寸 == null)
            {
                MessageBox.Show($"未找到型号 {产品型号} 的尺寸数据", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            变量.测试.Clear();  // 清空之前的列表
            List<数据项> 数据项列表 = new List<数据项>();

            if (变量.附件表数据.ContainsKey(型号SHEET))
            {
                string 提示信息 = $"工作表: {型号SHEET}\r\n";
                foreach (var item in 变量.附件表数据[型号SHEET])
                {
                    string[] 数据 = item.Split(',');
                    double 内容R;
                    if (double.TryParse(数据[2].Trim(), out 内容R))
                    {
                        // 将米转换为厘米并计算实际面积
                        double 实际面积 = 灯带尺寸.每厘米面积 * (内容R * 100);  // 内容R转换为厘米

                        string 内容A = 数据[0].Trim(); // 内容A在数组的第1个位置
                        string 内容O = 数据[1].Trim(); // 内容O在数组的第2个位置
                        string 标志 = Guid.NewGuid().ToString(); // 生成唯一标志

                        数据项列表.Add(new 数据项(内容A, 内容O, 实际面积, 标志));
                        变量.测试.Add(实际面积);
                    }
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

            if (变量.订单出线字典 == null || !变量.订单出线字典.Any())
            {
                MessageBox.Show("请先导入订单数据", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // 处理每个订单
            foreach (var 订单 in 变量.订单出线字典)
            {
                string 订单编号 = 订单.Key;
                var 订单明细列表 = 订单.Value;

                uiTextBox_状态.AppendText($"开始处理订单：{订单编号}" + Environment.NewLine);

                // 处理订单中的每个型号
                foreach (var (型号, 出线方式, F列内容) in 订单明细列表)
                {
                    处理订单包装(型号, 出线方式.ToList(), F列内容);
                }

                uiTextBox_状态.AppendText("订单处理完成" + Environment.NewLine);
                uiTextBox_状态.AppendText("------------------------" + Environment.NewLine);
            }


        }



        private void 处理订单包装(string 型号, List<string> 出线方式列表, string F列内容)
        {
            新包装 新包装实例 = new 新包装();
            StringBuilder 结果信息 = new StringBuilder();
            结果信息.AppendLine($"型号: {型号}");
            结果信息.AppendLine($"F列内容: {F列内容}");

            // 如果是F23B，需要移除B后缀进行查询
            string 查询型号 = 型号.Replace("B", "");

            // 处理每个出线方式
            if (出线方式列表.Count > 0)
            {
                foreach (var 出线方式 in 出线方式列表)
                {
                    string 查询类型 = 转换出线方式格式(型号, 出线方式, F列内容);
                    var 包装资料 = 新包装实例.查找包装资料(查询型号, 查询类型);

                    if (包装资料 != null)
                    {
                        结果信息.AppendLine($"出线方式: {出线方式}");
                        结果信息.AppendLine($"使用包装: {包装资料.半成品BOM物料码}");
                        结果信息.AppendLine($"总有效容积: {包装资料.总有效容积}");
                        结果信息.AppendLine("-------------------");
                    }
                    else
                    {
                        结果信息.AppendLine($"出线方式: {出线方式} - 未找到匹配的包装资料");
                        结果信息.AppendLine("-------------------");
                    }
                }
            }
            else
            {
                // 处理无出线方式的情况（如 TLX8 naked）
                string 查询类型 = "多条或短条包装";  // 根据实际情况调整
                var 包装资料 = 新包装实例.查找包装资料(查询型号, 查询类型);

                if (包装资料 != null)
                {
                    结果信息.AppendLine("多条或短条包装");
                    结果信息.AppendLine($"使用包装: {包装资料.半成品BOM物料码}");
                    结果信息.AppendLine($"总有效容积: {包装资料.总有效容积}");
                }
                else
                {
                    结果信息.AppendLine("未找到匹配的包装资料");
                }
            }

            MessageBox.Show(结果信息.ToString(), "包装查询结果", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private string 转换出线方式格式(string 型号, string 出线方式, string F列内容)
        {
            // 移除可能的"正弯"或"侧弯"前缀
            string 处理后出线方式 = 出线方式
                .Replace("正弯", "")
                .Replace("侧弯", "");

            // 根据F列内容判断是单层还是双层
            string 层类型 = F列内容.Contains("SC") ? "双层" : "单层";

            // 组合最终的查询类型
            return $"{层类型}注塑{处理后出线方式}";
        }









    }



}