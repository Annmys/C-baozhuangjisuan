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

            // 调用 uiButton1_Click 方法
            uiButton1_Click(this, EventArgs.Empty);
        }

        public void EXCEL订单数据_转列表(string excel文件路径)
        {
            try
            {
                using (var package = new ExcelPackage(new FileInfo(excel文件路径)))
                {
                    var worksheet = package.Workbook.Worksheets[0];

                    // 初始化列索引
                    int 订单编号列 = -1;
                    int 规格型号列 = -1;
                    int 销售数量列 = -1;
                    int F列 = -1;

                    // 在第1行查找列标题
                    for (int col = 1; col <= worksheet.Dimension.End.Column; col++)
                    {
                        string? 列标题 = worksheet.Cells[1, col].Text?.Trim();
                        if (string.IsNullOrEmpty(列标题)) continue;

                        switch (列标题)
                        {
                            case "单据编号":
                                订单编号列 = col;
                                break;
                            case "规格型号":
                                规格型号列 = col;
                                break;
                            case "销售数量":
                                销售数量列 = col;
                                break;
                            case "物料编码":
                                F列 = col;
                                break;
                        }
                    }

                    // 验证必要的列是否都找到
                    if (订单编号列 == -1 || 规格型号列 == -1 || 销售数量列 == -1)
                    {
                        throw new Exception("未找到必要的列标题（单据编号、规格型号或销售数量）");
                    }

                    变量.订单出线字典 = new Dictionary<string, List<(string 型号, HashSet<string> 出线方式, string F列内容, double 销售数量)>>();
                    变量.订单编号 = worksheet.Cells[2, 订单编号列].Text;
                    变量.订单出线字典[变量.订单编号] = new List<(string, HashSet<string>, string, double)>();

                    int startRow = -1;
                    string 当前型号 = "";
                    var 当前出线方式 = new HashSet<string>();
                    string 当前F列内容 = "";
                    double 当前销售数量 = 0;

                    // 从第2行开始处理数据
                    for (int row = 2; row <= worksheet.Dimension.End.Row; row++)
                    {
                        var materialCell = worksheet.Cells[row, F列];
                        var specCell = worksheet.Cells[row, 规格型号列];
                        var salesCell = worksheet.Cells[row, 销售数量列];

                        if (materialCell.Value != null && materialCell.Text.StartsWith("80."))
                        {
                            // 如果已经有一个开始行，说明找到了下一个80.开头的行，需要处理之前的数据
                            if (startRow != -1)
                            {
                                // 处理之前的区间
                                ProcessSection(worksheet, startRow, row - 1, 规格型号列, F列, 销售数量列, ref 当前型号, ref 当前出线方式, ref 当前F列内容, ref 当前销售数量);
                                变量.订单出线字典[变量.订单编号].Add((当前型号, new HashSet<string>(当前出线方式), 当前F列内容, 当前销售数量));
                            }

                            // 更新新的开始行
                            startRow = row;
                            当前出线方式 = new HashSet<string>();
                        }
                    }

                    // 处理最后一个区间
                    if (startRow != -1)
                    {
                        ProcessSection(worksheet, startRow, worksheet.Dimension.End.Row, 规格型号列, F列, 销售数量列, ref 当前型号, ref 当前出线方式, ref 当前F列内容, ref 当前销售数量);
                        变量.订单出线字典[变量.订单编号].Add((当前型号, new HashSet<string>(当前出线方式), 当前F列内容, 当前销售数量));
                    }

                    // 输出订单信息
                    输出订单信息(变量.订单编号, 变量.订单出线字典[变量.订单编号]);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"发生错误：{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ProcessSection(ExcelWorksheet worksheet, int startRow, int endRow, int 规格型号列, int F列, int 销售数量列,
    ref string 当前型号, ref HashSet<string> 当前出线方式, ref string 当前F列内容, ref double 当前销售数量)
        {
            // 获取主行信息
            var mainSpecCell = worksheet.Cells[startRow, 规格型号列];
            var mainMaterialCell = worksheet.Cells[startRow, F列];
            var mainSalesCell = worksheet.Cells[startRow, 销售数量列];

            string 规格型号 = mainSpecCell.Text;
            当前F列内容 = mainMaterialCell.Text;
            double.TryParse(mainSalesCell.Text, out 当前销售数量);

            

            // 从规格型号中提取基础型号
            if (规格型号.Contains("C-SFR-") || 规格型号.Contains("C-FR-"))
            {
                var parts = 规格型号.Split(new[] { "C-SFR-", "C-FR-" }, StringSplitOptions.RemoveEmptyEntries);
                if (parts.Length >= 1)
                {
                    string 型号部分 = parts[0];
                    var match = System.Text.RegularExpressions.Regex.Match(型号部分, @"F\d+");
                    if (match.Success)
                    {
                        当前型号 = match.Value;
                    }
                }
            }


            // 确定弯型
            var 需要判断弯型的型号 = new List<string> { "F22", "F23", "F2222", "F2219" };
            string 弯型前缀 = "";
            if (规格型号.Contains("【正弯】"))
            {
                弯型前缀 = "正弯";
            }
            else if (规格型号.Contains("【侧弯】"))
            {
                弯型前缀 = "侧弯";
            }

            // 清空并初始化基础出线方式集合
            var 基础出线方式 = new HashSet<string>();

            // 处理该区间内的所有行以获取出线方式
            for (int row = startRow; row <= endRow; row++)
            {
                var specCell = worksheet.Cells[row, 规格型号列];
                if (specCell.Value == null) continue;

                string 当前行规格型号 = specCell.Text;

                // 检查是否包含配件行标识
                if (当前行规格型号.Contains("B-硅胶注塑式") || 当前行规格型号.Contains("B-双层注塑式"))
                {
                    // 检查常见的出线方式关键词
                    if (当前行规格型号.Contains("端部出线"))
                    {
                        基础出线方式.Add("端部出线");
                    }
                    if (当前行规格型号.Contains("侧面出线"))
                    {
                        基础出线方式.Add("侧面出线");
                    }
                    if (当前行规格型号.Contains("底部出线"))
                    {
                        基础出线方式.Add("底部出线");
                    }
                }
            }

           

            // 最后添加弯型前缀（如果有）
            当前出线方式.Clear();
            bool 是特殊型号 = false;
            foreach (var 特殊型号 in 需要判断弯型的型号)
            {
                if (当前型号.StartsWith(特殊型号))
                {
                    是特殊型号 = true;
                    break;
                }
            }

            foreach (var 出线方式 in 基础出线方式)
            {
                //当前出线方式.Add(string.IsNullOrEmpty(弯型前缀) ? 出线方式 : $"{弯型前缀}{出线方式}");
                // 只有特殊型号且有弯型前缀时才添加前缀
                if (是特殊型号 && !string.IsNullOrEmpty(弯型前缀))
                {
                    当前出线方式.Add($"{弯型前缀}{出线方式}");
                }
                else
                {
                    当前出线方式.Add(出线方式);
                }
            }

            // Debug输出
            string debugInfo = $"处理区间：{startRow}-{endRow}\n" +
                              $"型号：{当前型号}\n" +
                              $"规格型号：{规格型号}\n" +
                              $"F列内容：{当前F列内容}\n" +
                              $"出线方式：{string.Join("，", 当前出线方式)}\n" +
                              $"销售数量：{当前销售数量}";

            MessageBox.Show(debugInfo);
        }

        private void 输出订单信息(string 订单编号, List<(string 型号, HashSet<string> 出线方式, string F列内容, double 销售数量)> 型号出线方式列表)
        {
            uiTextBox_状态.AppendText($"订单编号: {订单编号}" + Environment.NewLine);
            foreach (var (型号, 出线方式, F列内容, 销售数量) in 型号出线方式列表)
            {
                string 出线方式字符串 = 出线方式.Count > 0 ? string.Join("，", 出线方式) : "无";
                uiTextBox_状态.AppendText($"型号: {型号}" + Environment.NewLine);
                uiTextBox_状态.AppendText($"出线方式: {出线方式字符串}" + Environment.NewLine);
                uiTextBox_状态.AppendText($"F列内容: {F列内容}" + Environment.NewLine);
                uiTextBox_状态.AppendText($"销售数量: {销售数量}" + Environment.NewLine);
                uiTextBox_状态.AppendText(Environment.NewLine);
            }
        }


        private async void button_附件导入_Click_1(object sender, EventArgs e)
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
                                    // 检查是否为错误值
                                    if (cellA.Value is ExcelErrorValue || cellO.Value is ExcelErrorValue || cellR.Value is ExcelErrorValue)
                                    {
                                        continue; // 跳过错误值
                                    }

                                    string 内容A = cellA.Value.ToString();
                                    int 内容O = Convert.ToInt32(cellO.Value);
                                    double 内容R = Convert.ToDouble(cellR.Value);

                                    if (内容O > 1)
                                    {
                                        double 分割后的R值 = 内容R / 内容O;
                                        for (int i = 0; i < 内容O; i++)
                                        {
                                            变量.附件表数据[worksheet.Name].Add($"{内容A}, {1}, {分割后的R值}");
                                        }
                                    }
                                    else
                                    {
                                        变量.附件表数据[worksheet.Name].Add($"{内容A}, {内容O}, {内容R}");
                                    }
                                }
                            }
                        }
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
            foreach (var (型号, 出线方式, F列内容, 销售数量) in 变量.订单出线字典[变量.订单编号])
            {        
                string 出线方式字符串 = 出线方式.Count > 0 ? string.Join("，", 出线方式) : "无";
                MessageBox.Show($"型号：{型号}\n出线方式：{出线方式字符串}\nF列内容：{F列内容}\n销售数量：{销售数量}");
                //uiTextBox_状态.AppendText($"当前出线方式: {出线方式字符串}" + Environment.NewLine);

                //string 型号SHEET = "TLX8 naked";
                string 型号SHEET = F列内容;
                string 产品型号 = "F23";  // 根据实际情况设置对应的产品型号

                // 查找对应型号的灯带尺寸
                var 灯带尺寸 = 变量.灯带尺寸列表.FirstOrDefault(x => x.型号 == 产品型号);
                if (灯带尺寸 == null)
                {
                    MessageBox.Show($"未找到型号 {产品型号} 的尺寸数据", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                // 设置默认的包装容积（如果需要的话）
                变量.查找组合_基数 = 1420 * 0.7;  // 这里可以根据实际情况设置默认值

                // 调用修改后的开始组合方法
                开始组合(型号SHEET, 产品型号, 变量.订单编号);

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

            
        }

        private void 开始组合(string 型号SHEET, string 产品型号, string 订单编号)
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

            // 保存组合结果到Excel
            保存组合结果到Excel(ans, 数据项列表, 订单编号, 产品型号);

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

        private void 保存组合结果到Excel(List<List<double>> 组合结果, List<数据项> 数据项原列表, string 订单编号, string 型号)
        {
            // 创建订单编号的文件夹
            string 文件夹路径 = Path.Combine("输出结果", 订单编号);
            if (!Directory.Exists(文件夹路径))
            {
                Directory.CreateDirectory(文件夹路径);
            }

            // 为每个型号创建单独的Excel文件
            string 文件路径 = Path.Combine(文件夹路径, $"{型号}.xlsx");
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

                // 使用Invoke在UI线程上更新UI控件
                uiTextBox_状态.Invoke((MethodInvoker)(() =>
                {
                    uiTextBox_状态.AppendText($"开始处理订单：{订单编号}" + Environment.NewLine);
                }));

                // 处理订单中的每个型号
                foreach (var (型号, 出线方式, F列内容, 销售数量) in 订单明细列表)
                {
                    处理订单包装(型号, 出线方式.ToList(), F列内容, 销售数量);
                }

                uiTextBox_状态.Invoke((MethodInvoker)(() =>
        {
            uiTextBox_状态.AppendText("订单处理完成" + Environment.NewLine);
            uiTextBox_状态.AppendText("------------------------" + Environment.NewLine);
        }));
            }
        }

        private void 处理订单包装(string 型号, List<string> 出线方式列表, string F列内容, double 销售数量)
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

            //MessageBox.Show(结果信息.ToString(), "包装查询结果", MessageBoxButtons.OK, MessageBoxIcon.Information);
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