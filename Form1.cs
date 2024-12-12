using OfficeOpenXml; // EPPlus的命名空间.
using Sunny.UI;
using System.Text;
using OfficeOpenXml.Style;
using System.Drawing;

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

        //开始处理订单-------------------------------------------------------------------------------------
        private void button_订单导入_Click(object sender, EventArgs e)
        {
            // 在导入新订单前清理之前的数据
            清理之前数据();

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


        private void 清理之前数据()
        {
            // 清理变量中的数据
            变量.订单excel地址 = string.Empty;
            变量.订单编号 = string.Empty;
            if (变量.订单出线字典 != null)
            {
                变量.订单出线字典.Clear();
            }

            // 清理界面显示
            uiTextBox_订单地址.Text = string.Empty;
            uiTextBox_订单地址.BackColor = System.Drawing.SystemColors.Window;
            uiTextBox_状态.Clear();
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
                    if (当前行规格型号.Contains("侧部出线"))
                    {
                        基础出线方式.Add("侧部出线");
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

            //MessageBox.Show(debugInfo);
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
                    // 从匹配信息列表中查找对应型号的工作表名称
                    var 匹配信息 = 变量.订单附件匹配列表.FirstOrDefault(x =>
                        x.订单编号 == 订单编号 &&
                        x.产品型号 == 型号);

                    if (匹配信息 != null)
                    {
                        处理订单包装(型号, 出线方式.ToList(), F列内容, 销售数量, 匹配信息.工作表名称);
                    }
                    else
                    {
                        // 如果找不到匹配信息，可以使用一个默认值或者记录错误
                        uiTextBox_状态.Invoke((MethodInvoker)(() =>
                        {
                            uiTextBox_状态.AppendText($"警告：未找到型号 {型号} 的匹配工作表信息" + Environment.NewLine);
                        }));
                        处理订单包装(型号, 出线方式.ToList(), F列内容, 销售数量, "");
                    }
                }

                uiTextBox_状态.Invoke((MethodInvoker)(() =>
        {
            uiTextBox_状态.AppendText("------------------------" + Environment.NewLine);
            uiTextBox_状态.AppendText("订单导入完成" + Environment.NewLine);
            uiTextBox_状态.AppendText("------------------------" + Environment.NewLine);
        }));
            }
        }

        private void 处理订单包装(string 型号, List<string> 出线方式列表, string F列内容, double 销售数量, string 工作表名称)  // 添加工作表名称参数
        {
            新包装 新包装实例 = new 新包装();
            StringBuilder 结果信息 = new StringBuilder();
            结果信息.AppendLine($"型号: {型号}");
            结果信息.AppendLine($"F列内容: {F列内容}");

            // 获取当前工作表的序号前缀
            string 序号前缀 = "";
            if (变量.附件表数据.ContainsKey(工作表名称) && 变量.附件表数据[工作表名称].Count > 0)
            {
                string 首行数据 = 变量.附件表数据[工作表名称][0];
                string 序号 = 首行数据.Split(',')[0].Trim();
                序号前缀 = new string(序号.TakeWhile(c => !char.IsDigit(c)).ToArray());
            }

            // 如果是F23B，需要移除B后缀进行查询
            string 查询型号 = 型号.Replace("B", "");

            // 判断米数是否小于6米
            bool 使用600mm包装 = 销售数量 < 6;

            // 处理每个出线方式
            if (出线方式列表.Count > 0)
            {
                foreach (var 出线方式 in 出线方式列表)
                {
                    string 查询类型 = 转换出线方式格式(型号, 出线方式, F列内容);
                    var 包装资料 = 使用600mm包装
                        ? 新包装实例.查找600mm包装资料(查询型号, 查询类型)
                        : 新包装实例.查找包装资料(查询型号, 查询类型);

                    if (包装资料 != null)
                    {
                        结果信息.AppendLine($"序号前缀: {序号前缀}");  // 添加序号前缀信息
                        结果信息.AppendLine($"出线方式: {出线方式}");
                        结果信息.AppendLine($"使用包装: {包装资料.半成品BOM物料码}");
                        结果信息.AppendLine($"总有效容积: {包装资料.总有效容积}");
                        结果信息.AppendLine("-------------------");

                        // 存储序号前缀到匹配信息中
                        var 匹配信息 = 变量.订单附件匹配列表.FirstOrDefault(x => x.工作表名称 == 工作表名称);
                        if (匹配信息 != null)
                        {
                            匹配信息.设置Sheet序号前缀(变量.附件表数据[工作表名称]);
                        }
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
                // 处理无出线方式的情况
                string 查询类型 = "多条或短条包装";
                var 包装资料 = 使用600mm包装
                    ? 新包装实例.查找600mm包装资料(查询型号, 查询类型)
                    : 新包装实例.查找包装资料(查询型号, 查询类型);

                if (包装资料 != null)
                {
                    结果信息.AppendLine($"序号前缀: {序号前缀}");  // 添加序号前缀信息
                    结果信息.AppendLine("多条或短条包装");
                    结果信息.AppendLine($"使用包装: {包装资料.半成品BOM物料码}");
                    结果信息.AppendLine($"总有效容积: {包装资料.总有效容积}");

                    // 存储序号前缀到匹配信息中
                    var 匹配信息 = 变量.订单附件匹配列表.FirstOrDefault(x => x.工作表名称 == 工作表名称);
                    if (匹配信息 != null)
                    {
                        匹配信息.设置Sheet序号前缀(变量.附件表数据[工作表名称]);
                    }
                }
                else
                {
                    结果信息.AppendLine("未找到匹配的包装资料");
                }
            }

            uiTextBox_状态.AppendText(结果信息.ToString());
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


        //开始处理附件-------------------------------------------------------------------------------------
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

                //uiTextBox_状态.AppendText("附件导入完成" + Environment.NewLine);
                //uiTextBox_状态.AppendText("------------------------" + Environment.NewLine);
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
                        // 检查工作表是否为空
                        if (worksheet.Dimension == null)
                        {
                            continue; // 跳过空工作表
                        }

                        if (!变量.附件表数据.ContainsKey(worksheet.Name))
                        {
                            变量.附件表数据[worksheet.Name] = new List<string>();
                        }

                        // 查找包含"序号"的行（扩大搜索范围）
                        int 序号行号 = -1;
                        int 序号列号 = -1;
                        int 条数列号 = -1;
                        int 总长度列号 = -1;
                        int 标签码列号 = -1;
                        int 标签码1列号 = -1;
                        int 标签码2列号 = -1;
                        int 实际剪切长度毫米列号 = -1;
                        int 实际剪切长度米列号 = -1;

                        // 扩大搜索范围到前20行，但不超过工作表的实际行数
                        int maxSearchRow = Math.Min(20, worksheet.Dimension.End.Row);
                        for (int row = 1; row <= maxSearchRow; row++)
                        {
                            for (int col = 1; col <= worksheet.Dimension.End.Column; col++)
                            {
                                var cellValue = worksheet.Cells[row, col].Text?.Trim() ?? "";
                                if (cellValue.Contains("序号"))
                                {
                                    序号行号 = row;
                                    序号列号 = col;
                                    break;
                                }
                            }
                            if (序号行号 != -1) break;
                        }

                        // 如果没找到序号行，跳过这个工作表
                        if (序号行号 == -1)
                        {
                            continue;
                        }

                        // 在标题行查找其他列
                        for (int col = 1; col <= worksheet.Dimension.End.Column; col++)
                        {
                            var cellValue = worksheet.Cells[序号行号, col].Text?.Trim() ?? "";
                            if (cellValue.Contains("条数"))
                            {
                                条数列号 = col;
                            }
                            if (cellValue.Contains("总长度"))
                            {
                                总长度列号 = col;
                            }
                            if (cellValue.Contains("标签码") && !cellValue.Contains("标签码1") && !cellValue.Contains("标签码2"))
                            {
                                标签码列号 = col;
                            }
                            if (cellValue.Contains("标签码1"))
                            {
                                标签码1列号 = col;
                            }
                            if (cellValue.Contains("标签码2"))
                            {
                                标签码2列号 = col;
                            }
                            if (cellValue.Contains("实际剪切长度(毫米)"))
                            {
                                实际剪切长度毫米列号 = col;
                            }
                            if (cellValue.Contains("实际剪切长度(米)") )
                            {
                                实际剪切长度米列号 = col;
                            }
                        }

                        // 验证必要的列是否都找到
                        if (序号列号 == -1 || 条数列号 == -1 || 总长度列号 == -1)
                        {
                            MessageBox.Show($"工作表 {worksheet.Name} 中未找到必要的列标题", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            continue;
                        }

                        // 处理数据（从标题行的下一行开始，确保不超出工作表范围）
                        for (int row = 序号行号 + 1; row <= worksheet.Dimension.End.Row; row++)
                        {
                            try
                            {
                                var cell序号 = worksheet.Cells[row, 序号列号];
                                var cell条数 = worksheet.Cells[row, 条数列号];
                                var cell总米数 = worksheet.Cells[row, 总长度列号];

                                string 标签码 = "";
                                string 标签码1 = "";
                                string 标签码2 = "";

                                if (标签码列号 != -1)
                                {
                                    标签码 = worksheet.Cells[row, 标签码列号].Text?.Trim() ?? "";
                                }
                                else if (标签码1列号 != -1 || 标签码2列号 != -1)
                                {
                                    标签码1 = 标签码1列号 != -1 ? worksheet.Cells[row, 标签码1列号].Text?.Trim() ?? "" : "";
                                    标签码2 = 标签码2列号 != -1 ? worksheet.Cells[row, 标签码2列号].Text?.Trim() ?? "" : "";
                                }

                                if (cell序号.Value != null && cell条数.Value != null && cell总米数.Value != null)
                                {
                                    if (cell序号.Value is ExcelErrorValue || cell条数.Value is ExcelErrorValue || cell总米数.Value is ExcelErrorValue)
                                    {
                                        continue;
                                    }

                                    string 序号 = cell序号.Value.ToString();
                                    int 条数;
                                    double 总米数;

                                    if (int.TryParse(cell条数.Value.ToString(), out 条数) &&
                                        double.TryParse(cell总米数.Value.ToString(), out 总米数))
                                    {
                                        if (条数 > 1)
                                        {
                                            double 分割后的米数 = 总米数 / 条数;
                                            for (int i = 0; i < 条数; i++)
                                            {
                                                string 实际剪切长度毫米 = 实际剪切长度毫米列号 != -1 ? worksheet.Cells[row, 实际剪切长度毫米列号].Text?.Trim() ?? "" : "";
                                                string 实际剪切长度米 = 实际剪切长度米列号 != -1 ? worksheet.Cells[row, 实际剪切长度米列号].Text?.Trim() ?? "" : "";
                                                变量.附件表数据[worksheet.Name].Add($"{序号}, {1}, {分割后的米数}, {标签码}, {标签码1}, {标签码2}, {实际剪切长度毫米}, {实际剪切长度米}");
                                            }
                                        }
                                        else
                                        {
                                            string 实际剪切长度毫米 = 实际剪切长度毫米列号 != -1 ? worksheet.Cells[row, 实际剪切长度毫米列号].Text?.Trim() ?? "" : "";
                                            string 实际剪切长度米 = 实际剪切长度米列号 != -1 ? worksheet.Cells[row, 实际剪切长度米列号].Text?.Trim() ?? "" : "";
                                            变量.附件表数据[worksheet.Name].Add($"{序号}, {条数}, {总米数}, {标签码}, {标签码1}, {标签码2}, {实际剪切长度毫米}, {实际剪切长度米}");
                                        }
                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show($"处理第 {row} 行时出错：{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                continue;
                            }
                        }
                        // 处理完所有数据后，计算总米数和
                        if (变量.附件表数据[worksheet.Name].Count > 0)
                        {
                            double 总米数和 = 0;
                            foreach (string 记录 in 变量.附件表数据[worksheet.Name])
                            {
                                string[] 数据项 = 记录.Split(',');
                                if (数据项.Length >= 3 && double.TryParse(数据项[2].Trim(), out double 总米数))
                                {
                                    总米数和 += 总米数;
                                }
                            }

                            // 找到对应的匹配信息并设置序号前缀
                            foreach (var 匹配信息 in 变量.订单附件匹配列表.Where(x => x.工作表名称 == worksheet.Name))
                            {
                                匹配信息.设置Sheet序号前缀(变量.附件表数据[worksheet.Name]);
                            }

                            // 四舍五入到小数点后三位
                            总米数和 = Math.Round(总米数和, 3);

                            // 添加总计记录
                            变量.附件表数据[worksheet.Name].Add($"总计, , {总米数和:F3}, , , ");
                            //MessageBox.Show($"工作表 {worksheet.Name} 的总米数和: {总米数和:F3}");

                            // 如果订单已导入，进行数量比对
                            if (!string.IsNullOrEmpty(uiTextBox_订单地址.Text) && 变量.订单出线字典 != null && 变量.订单出线字典.Any())
                            {
                                foreach (var 订单 in 变量.订单出线字典)
                                {
                                    foreach (var (型号, 出线方式, F列内容, 销售数量) in 订单.Value)
                                    {
                                        // 将销售数量四舍五入到三位小数进行比较
                                        double 订单数量 = Math.Round(销售数量, 3);
                                        if (Math.Abs(订单数量 - 总米数和) < 0.001) // 使用小于0.001的差值来判断相等
                                        {
                                            var 匹配 = new 匹配信息(订单.Key, 型号, 出线方式, 订单数量, worksheet.Name, 总米数和);  // 修改这行，添加出线方式参数
                                            变量.订单附件匹配列表.Add(匹配);
                                            //MessageBox.Show(匹配.ToString(), "找到匹配", MessageBoxButtons.OK, MessageBoxIcon.Information);
                                            // 使用AppendText替代MessageBox
                                            //uiTextBox_状态.AppendText(匹配.ToString() + Environment.NewLine);
                                            //uiTextBox_状态.AppendText("------------------------" + Environment.NewLine);
                                        }
                                    }
                                }
                                // 检查是否找到匹配
                                var 当前工作表匹配 = 变量.订单附件匹配列表.FirstOrDefault(x => x.工作表名称 == worksheet.Name);
                                if (当前工作表匹配 != null)
                                {
                                    // 找到匹配，显示匹配信息
                                    uiTextBox_状态.AppendText(当前工作表匹配.ToString() + Environment.NewLine);
                                    uiTextBox_状态.AppendText("------------------------" + Environment.NewLine);
                                }
                                else
                                {
                                    // 未找到匹配，显示提示信息
                                    uiTextBox_状态.AppendText($"工作表 {worksheet.Name} (总米数: {总米数和:F3}) 未找到匹配的订单" + Environment.NewLine);
                                    uiTextBox_状态.AppendText("------------------------" + Environment.NewLine);
                                }
                            }

                        }
                    }

                    uiTextBox_状态.AppendText("附件导入完成" + Environment.NewLine);
                    uiTextBox_状态.AppendText("------------------------" + Environment.NewLine);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("读取Excel文件时出错：" + ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }


        private void button_开始处理_Click(object sender, EventArgs e)
        {
            if (变量.订单附件匹配列表.Count == 0)
            {
                MessageBox.Show("没有找到匹配的订单和附件信息", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            新包装 包装查询 = new 新包装();

            foreach (var 匹配信息 in 变量.订单附件匹配列表)
            {
                string 型号SHEET = 匹配信息.工作表名称;
                string 产品型号 = 匹配信息.产品型号;

                // 查找对应型号的灯带尺寸
                var 灯带尺寸 = 变量.灯带尺寸列表.FirstOrDefault(x => x.型号 == 产品型号.Replace("B", ""));
                if (灯带尺寸 == null)
                {
                    MessageBox.Show($"未找到型号 {产品型号} 的尺寸数据", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    continue;
                }

                // 获取订单中该型号的出线方式
                var 订单出线信息 = 变量.订单出线字典[匹配信息.订单编号]
                    .FirstOrDefault(x => x.Item1 == 产品型号);

                if (订单出线信息 == default)
                {
                    MessageBox.Show($"未找到型号 {产品型号} 的出线方式信息", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    continue;
                }

                // 确定包装类型
                string 包装类型;
                if (订单出线信息.Item2.Count == 0)
                {
                    包装类型 = "多条或短条包装";
                }
                else
                {
                    包装类型 = 订单出线信息.Item2.First(); // 使用第一个出线方式
                }

                // 获取候选包装列表
                var 候选包装列表 = 包装查询.获取候选包装列表(产品型号, 包装类型);

                //// 检查是否有候选包装
                //if (候选包装列表 != null && 候选包装列表.Any())
                //{
                //    // 构建显示字符串
                //    StringBuilder sb = new StringBuilder();
                //    sb.AppendLine("候选包装列表:");

                //    foreach (var 包装 in 候选包装列表)
                //    {
                //        sb.AppendLine($"----------------------------------------");
                //        sb.AppendLine($"包装名称: {包装.包装名称}");
                //        sb.AppendLine($"系统编码: {包装.半成品BOM物料码}");

                //    }

                //    // 显示信息框
                //    MessageBox.Show(sb.ToString(), "候选包装列表", MessageBoxButtons.OK, MessageBoxIcon.Information);
                //}


                if (候选包装列表.Any())
                {
                    int 数据条数 = 变量.附件表数据[型号SHEET].Count - 1;

                    //MessageBox.Show(数据条数.ToString(), "数据条数", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    // 选择最佳包装   测试V0.9
                    var 包装资料 = 包装查询.选择最佳包装(候选包装列表, 数据条数);

                    //MessageBox.Show(包装资料.半成品BOM物料码, "数据条数", MessageBoxButtons.OK, MessageBoxIcon.Information);

                    // 保存完整的包装资料到匹配信息中
                    匹配信息.选中包装资料 = 包装资料;  // 这一行是关键

                    if (包装资料 != null)
                    {
                        // 计算实际可用面积
                        double 实际面积 = (包装资料.总有效面积 - 包装资料.内撑纸卡面积) * 0.8;

                        // 设置实际可用容积
                        变量.查找组合_基数 = 实际面积;

                        uiTextBox_状态.AppendText($"型号 {产品型号} 使用包装: {包装资料.包装名称}" + Environment.NewLine);
                        uiTextBox_状态.AppendText($"总面积: {包装资料.总有效面积}" + Environment.NewLine);
                        uiTextBox_状态.AppendText($"实际可用面积: {实际面积}" + Environment.NewLine);
                        uiTextBox_状态.AppendText($"灯带数量: {数据条数}条" + Environment.NewLine);

                        // 添加包装类型信息
                        string 包装类型说明 = 包装资料.是多条短条专用包装 ? "多条专用包装" :
                                          包装资料.允许多条包装 ? "多条包装" : "普通包装";
                        uiTextBox_状态.AppendText($"包装类型: {包装类型说明}" + Environment.NewLine);
                        uiTextBox_状态.AppendText("------------------------" + Environment.NewLine);
                    }
                    else
                    {
                        MessageBox.Show($"未找到型号 {产品型号} 出线方式 {包装类型} 的匹配包装", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        // 使用默认容积
                        变量.查找组合_基数 = 1420 - 706;

                        uiTextBox_状态.AppendText($"型号 {产品型号} 使用默认包装容积: {变量.查找组合_基数}" + Environment.NewLine);
                        uiTextBox_状态.AppendText("------------------------" + Environment.NewLine);
                    }
                }
                else
                {
                    MessageBox.Show($"未找到型号 {产品型号} 出线方式 {包装类型} 的候选包装", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    // 使用默认容积
                    变量.查找组合_基数 = 1420 - 706;

                    uiTextBox_状态.AppendText($"型号 {产品型号} 使用默认包装容积: {变量.查找组合_基数}" + Environment.NewLine);
                    uiTextBox_状态.AppendText("------------------------" + Environment.NewLine);
                }

                //// 查找匹配的包装资料
                //var 包装资料 = 包装查询.查找包装资料(产品型号, 包装类型);

                //if (包装资料 != null)
                //{
                //    // 计算实际可用面积
                //    double 实际面积 = (包装资料.总有效面积-包装资料.内撑纸卡面积)*0.8 ;
                //    //double 实际面积 = (包装资料.总有效面积-包装资料.圆形平卡面积)*0.7 ;


                //    // 设置实际可用容积
                //    变量.查找组合_基数 = 实际面积;

                //    uiTextBox_状态.AppendText($"型号 {产品型号} 使用包装: {包装资料.包装名称}" + Environment.NewLine);
                //    uiTextBox_状态.AppendText($"总面积: {包装资料.总有效面积}" + Environment.NewLine);
                //    uiTextBox_状态.AppendText($"实际可用面积: {实际面积}" + Environment.NewLine);
                //    uiTextBox_状态.AppendText("------------------------" + Environment.NewLine);
                //}
                //else
                //{
                //    MessageBox.Show($"未找到型号 {产品型号} 出线方式 {包装类型} 的匹配包装", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                //    // 使用默认容积
                //    变量.查找组合_基数 = 1420-706;

                //    uiTextBox_状态.AppendText($"型号 {产品型号} 使用默认包装容积: {变量.查找组合_基数}" + Environment.NewLine);
                //    uiTextBox_状态.AppendText("------------------------" + Environment.NewLine);
                //}

                // 调用修改后的开始组合方法
                开始组合(型号SHEET, 产品型号, 匹配信息.订单编号);
            }
        }

        private void 开始组合(string 型号SHEET, string 产品型号, string 订单编号)
        {
            var 灯带尺寸 = 变量.灯带尺寸列表.FirstOrDefault(x => x.型号 == 产品型号.Replace("B", ""));
            if (灯带尺寸 == null)
            {
                MessageBox.Show($"未找到型号 {产品型号} 的尺寸数据", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            变量.测试.Clear();
            List<数据项> 数据项列表 = new List<数据项>();
            HashSet<string> 已使用序号 = new HashSet<string>();  // 跟踪已使用的序号

            // 添加调试信息
            StringBuilder 调试信息 = new StringBuilder();
            调试信息.AppendLine($"工作表 {型号SHEET} 的所有实际面积值：");
            调试信息.AppendLine($"灯带型号: {产品型号}, 每厘米面积: {灯带尺寸.每厘米面积}");
            调试信息.AppendLine("------------------------");

            if (变量.附件表数据.ContainsKey(型号SHEET))
            {
                var 有效数据 = 变量.附件表数据[型号SHEET].Take(变量.附件表数据[型号SHEET].Count - 1);
                int 数据计数 = 0;

                foreach (var item in 有效数据)
                {
                    string[] 数据 = item.Split(',');
                    if (double.TryParse(数据[2].Trim(), out double 内容R))
                    {
                        double 实际面积 = 灯带尺寸.每厘米面积 * (内容R * 100);
                        数据计数++;

                        string 内容A = 数据[0].Trim();
                        string 内容O = 数据[1].Trim();

                        调试信息.AppendLine($"数据 {数据计数}:");
                        调试信息.AppendLine($"序号: {内容A}");
                        调试信息.AppendLine($"条数: {内容O}");
                        调试信息.AppendLine($"米数: {内容R:F3}");
                        调试信息.AppendLine($"实际面积: {实际面积:F3}");
                        调试信息.AppendLine("------------------------");

                        string 标志 = Guid.NewGuid().ToString();
                        数据项列表.Add(new 数据项(内容A, 内容O, 实际面积, 标志));
                        变量.测试.Add(实际面积);

                        uiTextBox_状态.AppendText($"处理数据 - 序号: {内容A}, 条数: {内容O}, 米数: {内容R:F3}" + Environment.NewLine);
                    }
                }

                // 添加汇总信息
                调试信息.AppendLine("\n汇总信息：");
                调试信息.AppendLine($"总数据条数: {数据计数}");
                调试信息.AppendLine($"添加到测试列表的面积值：");
                foreach (var 面积 in 变量.测试)
                {
                    调试信息.AppendLine($"面积: {面积:F3}");
                }
            }

            // 显示调试信息
            //MessageBox.Show(调试信息.ToString(), "数据导入调试信息", MessageBoxButtons.OK, MessageBoxIcon.Information);

            Solution s = new Solution();
            var ans = s.CalculateCombinations(变量.测试, 变量.查找组合_基数);

            // 显示组合结果
            StringBuilder 组合信息 = new StringBuilder();
            组合信息.AppendLine($"共找到 {ans.Count} 种组合方案");
            组合信息.AppendLine($"输入数据总数: {变量.测试.Count}");
            组合信息.AppendLine("------------------------");

            // 为每个组合创建一个可用数据项列表的副本
            List<List<数据项>> 有效组合列表 = new List<List<数据项>>();

            for (int i = 0; i < ans.Count; i++)
            {
                组合信息.AppendLine($"方案 {i + 1}:");
                var combination = ans[i];
                List<数据项> 当前组合数据项 = new List<数据项>();
                List<数据项> 可用数据项 = 数据项列表.Where(x => !已使用序号.Contains(x.内容A)).ToList();
                bool 组合有效 = true;

                foreach (double 面积值 in combination)
                {
                    var 匹配项 = 可用数据项.FirstOrDefault(d => Math.Abs(d.内容R - 面积值) < 0.001);
                    if (匹配项 != null)
                    {
                        当前组合数据项.Add(匹配项);
                        可用数据项.Remove(匹配项);
                        组合信息.AppendLine($"序号: {匹配项.内容A}, 条数: {匹配项.内容O}, 米数: {(面积值 / 灯带尺寸.每厘米面积 / 100):F3}");
                    }
                    else
                    {
                        组合有效 = false;
                        break;
                    }
                }

                if (组合有效)
                {
                    有效组合列表.Add(当前组合数据项);
                    foreach (var 项 in 当前组合数据项)
                    {
                        已使用序号.Add(项.内容A);
                    }
                    组合信息.AppendLine($"组合总面积: {combination.Sum():F3}");
                }
                组合信息.AppendLine("------------------------");
            }

            //MessageBox.Show(组合信息.ToString(), "组合计算结果", MessageBoxButtons.OK, MessageBoxIcon.Information);
            uiTextBox_状态.AppendText(组合信息.ToString());

            uiTextBox_状态.AppendText($"共找到 {有效组合列表.Count} 种有效组合方案" + Environment.NewLine);
            uiTextBox_状态.AppendText("------------------------" + Environment.NewLine);

            // 输出未使用的序号
            var 未使用序号 = 数据项列表.Where(x => !已使用序号.Contains(x.内容A)).Select(x => x.内容A);
            if (未使用序号.Any())
            {
                uiTextBox_状态.AppendText("未使用的序号：" + Environment.NewLine);
                foreach (var 序号 in 未使用序号)
                {
                    uiTextBox_状态.AppendText($"{序号}" + Environment.NewLine);
                }
            }

            // 只保存有效的组合到Excel
            保存组合结果到Excel(ans.Take(有效组合列表.Count).ToList(), 数据项列表, 订单编号, 型号SHEET, 灯带尺寸);
        }


        private void 保存组合结果到Excel(List<List<double>> 组合结果, List<数据项> 数据项原列表, string 订单编号, string 工作表名称, 灯带尺寸 灯带尺寸对象)
        {
            string 文件夹路径 = Path.Combine("输出结果", 订单编号);

            if (!Directory.Exists(文件夹路径))
            {
                Directory.CreateDirectory(文件夹路径);
            }

            string 文件路径 = Path.Combine(文件夹路径, $"{工作表名称}（{灯带尺寸对象.型号}）.xlsx");
            FileInfo 文件信息 = new FileInfo(文件路径);

            // 创建一个全局的已使用序号集合
            HashSet<string> 全局已使用序号 = new HashSet<string>();

            using (ExcelPackage excelPackage = new ExcelPackage(文件信息))
            {
                for (int i = 0; i < 组合结果.Count; i++)
                {
                    var combination = 组合结果[i];
                    List<数据项> 可用数据项 = 数据项原列表
                        .Where(x => !全局已使用序号.Contains(x.内容A))
                        .ToList();

                    // 如果可用数据项不足，跳过此组合
                    if (可用数据项.Count < combination.Count)
                    {
                        continue;
                    }

                    string 工作表名 = $"第 {i + 1}盒";
                    if (excelPackage.Workbook.Worksheets.Any(ws => ws.Name == 工作表名))
                    {
                        工作表名 = $"第 {i + 1}盒_{Guid.NewGuid().ToString().Substring(0, 4)}";
                    }

                    ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add(工作表名);

                    worksheet.Cells[1, 1].Value = "序号";
                    worksheet.Cells[1, 2].Value = "条数";
                    worksheet.Cells[1, 3].Value = "米数";

                    // 检查是否有标签码数据
                    var 原始数据 = 变量.附件表数据[工作表名称];
                    var 第一行数据 = 原始数据[0].Split(',');
                    bool 有标签码1 = false;
                    bool 有标签码2 = false;

                    if (第一行数据.Length > 3)  // 有标签码
                    {
                        if (第一行数据.Length > 4)  // 有两个标签码
                        {
                            worksheet.Cells[1, 4].Value = "标签码1";
                            worksheet.Cells[1, 5].Value = "标签码2";
                            有标签码1 = true;
                            有标签码2 = true;
                        }
                        else  // 只有一个标签码
                        {
                            worksheet.Cells[1, 4].Value = "标签码";
                            worksheet.Cells[1, 5].Value = " ";
                            有标签码1 = true;
                        }
                    }

                    worksheet.Cells[1, 6].Value = "实际剪切长度(毫米)";
                    worksheet.Cells[1, 7].Value = "实际剪切长度(米)";

                    uiTextBox_状态.AppendText($"第 {i + 1} 种组合方案：" + Environment.NewLine);

                    bool 组合有效 = true;
                    List<数据项> 当前组合已选项 = new List<数据项>();

                    for (int j = 0; j < combination.Count; j++)
                    {
                        double 面积值 = combination[j];
                        var 项 = 可用数据项.FirstOrDefault(d => Math.Abs(d.内容R - 面积值) < 0.001);

                        if (项 != null)
                        {
                            worksheet.Cells[j + 2, 1].Value = 项.内容A;
                            worksheet.Cells[j + 2, 2].Value = 项.内容O;

                            // 使用传入的灯带尺寸参数
                            double 米数 = 面积值 / (灯带尺寸对象.每厘米面积 * 100);
                            worksheet.Cells[j + 2, 3].Value = Math.Round(米数, 3);

                            // 从原始数据中获取标签码
                            if (变量.附件表数据.ContainsKey(工作表名称))
                            {
                                var 原始数据列表 = 变量.附件表数据[工作表名称];
                                // 根据序号查找对应的原始数据
                                var 匹配数据 = 原始数据列表.FirstOrDefault(x => x.Split(',')[0].Trim() == 项.内容A.Trim());
                                if (匹配数据 != null)
                                {
                                    var 数据项 = 匹配数据.Split(',');
                                    if (数据项.Length > 3)
                                    {
                                        var 标签码 = 数据项[3].Trim();

                                        if (!string.IsNullOrEmpty(标签码))
                                        {
                                            // 如果标签码不为空，则将其放入D列，原标签码放入D1
                                            worksheet.Cells[j + 2, 4].Value = 标签码;  // D列
                                            worksheet.Cells[1, 4].Value = "标签码";    // D1
                                        }
                                        else
                                        {
                                            // 如果标签码为空，则按原逻辑处理
                                            if (数据项.Length > 4)
                                            {
                                                worksheet.Cells[j + 2, 4].Value = 数据项[4].Trim();  // 标签码1

                                                if (数据项.Length > 5)
                                                {
                                                    worksheet.Cells[j + 2, 5].Value = 数据项[5].Trim();  // 标签码2
                                                }
                                            }
                                        }
                                    }
                                    if (数据项.Length > 7)  // 确保有实际剪切长度(毫米)
                                    {
                                        worksheet.Cells[j + 2, 6].Value = 数据项[6].Trim(); // 实际剪切长度(毫米)
                                    }
                                    if (数据项.Length > 8)  // 确保有实际剪切长度(米)
                                    {
                                        worksheet.Cells[j + 2, 7].Value = 数据项[7].Trim(); // 实际剪切长度(米)
                                    }
                                }
                            }

                            uiTextBox_状态.AppendText($"序号: {项.内容A}, 条数: {项.内容O}, 米数: {Math.Round(米数, 3)}" + Environment.NewLine);

                            当前组合已选项.Add(项);
                            可用数据项.Remove(项);
                        }
                        else
                        {
                            组合有效 = false;
                            break;
                        }
                    }

                    if (组合有效)
                    {
                        // 只有当组合完全有效时，才更新全局已使用序号
                        foreach (var 项 in 当前组合已选项)
                        {
                            全局已使用序号.Add(项.内容A);
                        }

                        // 添加组合总面积
                        int lastRow = worksheet.Dimension?.End.Row ?? 2;
                        worksheet.Cells[lastRow + 1, 1].Value = "总面积";
                        worksheet.Cells[lastRow + 1, 3].Value = combination.Sum();

                        uiTextBox_状态.AppendText($"组合总面积: {combination.Sum():F3}" + Environment.NewLine);
                        uiTextBox_状态.AppendText("------------------------" + Environment.NewLine);

                        worksheet.Column(1).AutoFit();
                        worksheet.Column(2).AutoFit();
                        worksheet.Column(3).AutoFit();
                    }
                    else
                    {
                        // 如果组合无效，删除当前工作表
                        excelPackage.Workbook.Worksheets.Delete(worksheet);
                    }
                }

                excelPackage.Save();
            }

            // 输出未使用的序号
            var 未使用序号 = 数据项原列表
                .Where(x => !全局已使用序号.Contains(x.内容A))
                .Select(x => x.内容A);

            if (未使用序号.Any())
            {
                uiTextBox_状态.AppendText("未使用的序号：" + Environment.NewLine);
                foreach (var 序号 in 未使用序号)
                {
                    uiTextBox_状态.AppendText($"{序号}" + Environment.NewLine);
                }
            }

            uiTextBox_状态.AppendText($"结果已保存到 {文件路径}" + Environment.NewLine);
            uiTextBox_状态.AppendText("------------------------" + Environment.NewLine);

            // 获取BOM物料码 - 直接使用之前保存的包装信息
            var 匹配信息 = 变量.订单附件匹配列表.FirstOrDefault(x =>
                x.工作表名称 == 工作表名称);

            dynamic? 包装资料 = null;
            string? BOM物料码 = null;

            if (匹配信息?.选中包装资料 != null)
            {
                包装资料 = 匹配信息.选中包装资料;
                BOM物料码 = 包装资料.半成品BOM物料码;
                uiTextBox_状态.AppendText($"使用已选择的包装 - BOM物料码: {BOM物料码}" + Environment.NewLine);
            }
            else
            {
                uiTextBox_状态.AppendText($"警告：未找到之前选择的包装信息 - 工作表: {工作表名称}" + Environment.NewLine);
                return; // 如果没有包装信息，直接返回
            }

            // 创建或更新汇总Excel
            string 汇总文件路径 = Path.Combine("输出结果", 订单编号, "包装材料需求流转单.xlsx");
            FileInfo 汇总文件信息 = new FileInfo(汇总文件路径);

            

            // 创建一个字典来统计每个BOM物料码的使用数量
            Dictionary<string, int> BOM物料码统计 = new Dictionary<string, int>();
            if (包装资料?.半成品BOM物料码 != null)
            {
                string key = 包装资料.半成品BOM物料码;
                if (!BOM物料码统计.ContainsKey(key))
                {
                    BOM物料码统计[key] = 0;
                }
                BOM物料码统计[key] += 组合结果.Count;
            }

            using (ExcelPackage 汇总包 = new ExcelPackage(汇总文件信息))
            {
                ExcelWorksheet 汇总表;
                if (汇总包.Workbook.Worksheets.Any(ws => ws.Name == "包装材料需求流转单"))
                {
                    汇总表 = 汇总包.Workbook.Worksheets["包装材料需求流转单"];

                    // 获取已有数据的范围
                    int lastRow = 汇总表.Dimension?.End.Row ?? 6;

                    // 先删除所有匹配的纸箱行
                    for (int row = 7; row <= lastRow; row++)
                    {
                        string 现有文件名 = 汇总表.Cells[row, 11].Text;
                        string 物料类型 = 汇总表.Cells[row, 2].Text;

                        // 如果是纸箱且文件名匹配
                        if (物料类型 == "纸箱" && 现有文件名 == Path.GetFileName(文件路径))
                        {
                            汇总表.DeleteRow(row);
                            row--; // 调整行索引
                            lastRow--; // 调整总行数
                        }
                    }

                    // 删除匹配的POF热缩袋行
                    for (int row = 7; row <= lastRow; row++)
                    {
                        string 现有文件名 = 汇总表.Cells[row, 11].Text;
                        string 物料类型 = 汇总表.Cells[row, 2].Text;

                        if (物料类型 == "POF热缩袋" && 现有文件名 == Path.GetFileName(文件路径))
                        {
                            汇总表.DeleteRow(row);
                            row--; // 调整行索引
                            lastRow--; // 调整总行数
                        }
                    }

                    // 删除匹配的全棉织带行
                    for (int row = 7; row <= lastRow; row++)
                    {
                        string 现有文件名 = 汇总表.Cells[row, 11].Text;
                        string 物料类型 = 汇总表.Cells[row, 2].Text;

                        if (物料类型 == "全棉织带" && 现有文件名 == Path.GetFileName(文件路径))
                        {
                            汇总表.DeleteRow(row);
                            row--; // 调整行索引
                            lastRow--; // 调整总行数
                        }
                    }

                    // 再删除匹配的半成品BOM行
                    for (int row = 7; row <= lastRow; row++)
                    {
                        string 现有文件名 = 汇总表.Cells[row, 11].Text;
                        string 现有物料码 = 汇总表.Cells[row, 3].Text;
                        string 物料类型 = 汇总表.Cells[row, 2].Text;

                        // 如果是半成品BOM且匹配
                        if (物料类型 == "半成品BOM物料码" && 现有文件名 == Path.GetFileName(文件路径) && 现有物料码 == 包装资料.半成品BOM物料码)
                        {
                            汇总表.DeleteRow(row);
                            row--; // 调整行索引
                            lastRow--; // 调整总行数
                        }
                    }
                }
                else
                {
                    // 修改表头设置部分的代码
                    汇总表 = 汇总包.Workbook.Worksheets.Add("包装材料需求流转单");

                    // 设置标题（合并A2-I2）
                    汇总表.Cells["A2:I2"].Merge = true;
                    汇总表.Cells["A2"].Value = "包装材料需求流转单";
                    汇总表.Cells["A2"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    汇总表.Cells["A2"].Style.Font.Bold = true;

                    // 第3行设置
                    汇总表.Cells["A3"].Value = "订单号:";
                    汇总表.Cells["B3"].Value = 订单编号;
                    汇总表.Cells["C3"].Value = "客户代码:";
                    汇总表.Cells["E3:I3"].Merge = true;
                    汇总表.Cells["E3"].Value = "完成时间:";
                    // D3-I3保持空白

                    // 第4行设置
                    汇总表.Cells["A4:B4"].Merge = true;
                    汇总表.Cells["A4"].Value = "制单日期:";
                    汇总表.Cells["C4:D4"].Merge = true;
                    汇总表.Cells["C4"].Value = "制单人:";
                    汇总表.Cells["E4:G4"].Merge = true;
                    汇总表.Cells["E4"].Value = "业务员:";
                    汇总表.Cells["H4:I4"].Merge = true;
                    汇总表.Cells["H4"].Value = "TO: 仓库、品质、包装、配件";

                    // 第5行（表头）设置
                    汇总表.Cells["A5:A6"].Merge = true;
                    汇总表.Cells["A5"].Value = "产品型号";
                    汇总表.Cells["B5:B6"].Merge = true;
                    汇总表.Cells["B5"].Value = "物料";
                    汇总表.Cells["C5:C6"].Merge = true;
                    汇总表.Cells["C5"].Value = "物料编码";
                    汇总表.Cells["D5:E5"].Merge = true;
                    汇总表.Cells["D5"].Value = "包装要求及需求数量";
                    汇总表.Cells["D6"].Value = "规格";
                    汇总表.Cells["E6"].Value = "需求数量";
                    汇总表.Cells["F6"].Value = "仓位";
                    汇总表.Cells["G5:H5"].Merge = true;
                    汇总表.Cells["G5"].Value = "仓库";
                    汇总表.Cells["G6"].Value = "是否缺料";
                    汇总表.Cells["H6"].Value = "欠料订单号/时间";
                    汇总表.Cells["I5:I6"].Merge = true;
                    汇总表.Cells["I5"].Value = "备注";

                    // 设置整个表格的边框
                    var tableRange = 汇总表.Cells["A2:I6"];
                    tableRange.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    tableRange.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    tableRange.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    tableRange.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

                    // 设置所有单元格的内部边框
                    for (int row = 2; row <= 5; row++)
                    {
                        for (int col = 1; col <= 9; col++)
                        {
                            var cell = 汇总表.Cells[row, col];
                            cell.Style.Border.BorderAround(ExcelBorderStyle.Thin);
                        }
                    }

                    // 设置字体和对齐方式
                    tableRange.Style.Font.Name = "微软雅黑";
                    tableRange.Style.Font.Size = 10;
                    tableRange.Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                    // 设置行高
                    汇总表.Row(2).Height = 33; // 标题行
                    汇总表.Row(3).Height = 24.7;
                    汇总表.Row(4).Height = 24.7;
                    汇总表.Row(5).Height = 24.7;
                    汇总表.Row(6).Height = 34.5;

                    // 设置列宽
                    for (int col = 1; col <= 9; col++)
                    {
                        汇总表.Column(col).AutoFit();
                    }

                    // 设置标题字体
                    汇总表.Cells["A2"].Style.Font.Size = 14;
                    汇总表.Cells["A2"].Style.Font.Bold = true;

                    // 设置特定单元格的字体加粗
                    var boldCells = new[] { "A3", "C3","E3","E4", "A4", "H4","B4", "C4", "F4" };
                    foreach (var cell in boldCells)
                    {
                        汇总表.Cells[cell].Style.Font.Bold = true;
                    }

                    // 设置第5行表头的字体加粗
                    汇总表.Cells["A5:I5"].Style.Font.Bold = true;
                    汇总表.Cells["A6:I6"].Style.Font.Bold = true;
                }



                // 获取当前数据的最后一行
                int currentRow = 8;
                if (汇总表.Dimension != null)
                {
                    currentRow = 汇总表.Dimension.End.Row + 1;
                }

                // 创建一个HashSet来跟踪已处理的组合
                HashSet<string> 已处理组合 = new HashSet<string>();

                if (包装资料 != null)
                {
                    foreach (var kvp in BOM物料码统计)
                    {
                        string bom物料码 = kvp.Key;
                        int 需求数量 = kvp.Value;
                        string 当前文件名 = Path.GetFileName(文件路径);

                        // 创建更详细的唯一标识（包含BOM物料码、文件名和需求数量）
                        string 组合标识 = $"{bom物料码}_{当前文件名}_{需求数量}";

                        // 检查是否已经处理过这个组合
                        if (已处理组合.Contains(组合标识))
                        {
                            continue; // 跳过已处理的组合
                        }

                        // 添加到已处理集合
                        已处理组合.Add(组合标识);

                        // 获取当前工作表的序号前缀
                        string 序号前缀 = "";
                        if (变量.附件表数据.ContainsKey(工作表名称) && 变量.附件表数据[工作表名称].Count > 0)
                        {
                            string 首行数据 = 变量.附件表数据[工作表名称][0];
                            string 序号 = 首行数据.Split(',')[0].Trim();
                            序号前缀 = new string(序号.TakeWhile(c => !char.IsDigit(c)).ToArray());
                        }

                        // 添加半成品BOM信息
                        汇总表.Cells[currentRow, 1].Value =序号前缀;
                        汇总表.Cells[currentRow, 2].Value = "半成品BOM物料码";
                        汇总表.Cells[currentRow, 3].Value = bom物料码;
                        汇总表.Cells[currentRow, 4].Value = " ";
                        汇总表.Cells[currentRow, 5].Value = 需求数量;
                        汇总表.Cells[currentRow, 6].Value = "#N/A";
                        汇总表.Cells[currentRow, 11].Value = 当前文件名;  // J列添加对应的Excel文件名称

                        var 半成品BOM = 汇总表.Cells[currentRow, 1, currentRow, 9];
                        半成品BOM.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                        半成品BOM.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                        半成品BOM.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        半成品BOM.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

                        currentRow++;

                        // 添加POF热缩袋信息
                        if (!string.IsNullOrEmpty(包装资料.POF热缩袋))
                        {
                            
                            汇总表.Cells[currentRow, 2].Value = "POF热缩袋";
                            汇总表.Cells[currentRow, 3].Value = 包装资料.POF热缩袋;
                            汇总表.Cells[currentRow, 4].Value = 包装资料.POF热缩袋_系统尺寸;
                            汇总表.Cells[currentRow, 5].Value = 需求数量;
                            汇总表.Cells[currentRow, 6].Value = "#N/A";
                            汇总表.Cells[currentRow, 11].Value = 当前文件名; 

                            // 设置POF热缩袋行的样式
                            var pof范围 = 汇总表.Cells[currentRow, 1, currentRow, 9];
                            pof范围.Style.Fill.PatternType = ExcelFillStyle.Solid;
                            pof范围.Style.Fill.BackgroundColor.SetColor(Color.LightGreen);  // 使用不同的颜色区分
                            pof范围.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                            pof范围.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                            pof范围.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                            pof范围.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

                            currentRow++;
                        }

                        // 添加全棉织带信息
                        if (!string.IsNullOrEmpty(包装资料.全棉织带))
                        {
                            
                            汇总表.Cells[currentRow, 2].Value = "全棉织带";
                            汇总表.Cells[currentRow, 3].Value = 包装资料.全棉织带;
                            汇总表.Cells[currentRow, 4].Value = 包装资料.全棉织带_系统尺寸;
                            汇总表.Cells[currentRow, 5].Value = 需求数量;
                            汇总表.Cells[currentRow, 6].Value = "#N/A";
                            汇总表.Cells[currentRow, 11].Value = 当前文件名; 

                            // 设置全棉织带行的样式
                            var 织带范围 = 汇总表.Cells[currentRow, 1, currentRow, 9];
                            织带范围.Style.Fill.PatternType = ExcelFillStyle.Solid;
                            织带范围.Style.Fill.BackgroundColor.SetColor(Color.LightGreen);  // 使用不同的颜色区分
                            织带范围.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                            织带范围.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                            织带范围.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                            织带范围.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

                            currentRow++;
                        }

                        // 处理纸箱信息
                        var 纸箱组合列表 = 获取最佳纸箱组合(需求数量, 包装资料);
                        foreach (var 纸箱信息 in 纸箱组合列表)
                        {
                            汇总表.Cells[currentRow, 2].Value = "纸箱";
                            汇总表.Cells[currentRow, 3].Value = 纸箱信息.Item1;  // 使用Item1获取编号
                            汇总表.Cells[currentRow, 4].Value = 获取系统尺寸(纸箱信息.Item3, 包装资料);  // 使用Item3获取盒数
                            汇总表.Cells[currentRow, 5].Value = 纸箱信息.Item2;  // 使用Item2获取数量
                            汇总表.Cells[currentRow, 6].Value = "#N/A";
                            汇总表.Cells[currentRow, 9].Value = $"{纸箱信息.Item3}盒装标准";
                            汇总表.Cells[currentRow, 11].Value = 当前文件名; 

                            // 设置纸箱行的样式
                            var range = 汇总表.Cells[currentRow, 1, currentRow, 9];
                            range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                            range.Style.Fill.BackgroundColor.SetColor(Color.LightGreen);
                            range.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                            range.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                            range.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                            range.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

                            currentRow++;
                        }
                    }

                    // 自动调整列宽
                    for (int col = 1; col <= 9; col++)
                    {
                        汇总表.Column(col).AutoFit();
                    }

                    汇总包.Save();

                }


                
            }

            uiTextBox_状态.AppendText($"包装汇总已更新到 {汇总文件路径}" + Environment.NewLine);
            uiTextBox_状态.AppendText("------------------------" + Environment.NewLine);


        }

        private string 获取系统尺寸(int 盒数, dynamic 包装资料)
        {
            switch (盒数)
            {
                case 1:
                    return 包装资料.单盒装选用_系统尺寸 ?? "";
                case 2:
                    return 包装资料.二盒装选用_系统尺寸 ?? "";
                case 3:
                    return 包装资料.三盒装选用_系统尺寸 ?? "";
                case 5:
                    return 包装资料.五盒装选用_系统尺寸 ?? "";
                default:
                    return "";
            }
        }

        // 添加辅助方法来获取最佳纸箱组合
        private List<(string 编号, int 数量, int 盒数)> 获取最佳纸箱组合(int 总盒数, 新包装资料 包装资料)
        {
            var 结果 = new List<(string 编号, int 数量, int 盒数)>();
            int 剩余盒数 = 总盒数;

            // 优先使用大容量的纸箱
            if (!string.IsNullOrEmpty(包装资料.五盒装选用) && 剩余盒数 >= 5)
            {
                int 五盒装数量 = 剩余盒数 / 5;
                结果.Add((包装资料.五盒装选用, 五盒装数量, 5));
                剩余盒数 %= 5;
            }

            if (!string.IsNullOrEmpty(包装资料.三盒装选用) && 剩余盒数 >= 3)
            {
                int 三盒装数量 = 剩余盒数 / 3;
                结果.Add((包装资料.三盒装选用, 三盒装数量, 3));
                剩余盒数 %= 3;
            }

            if (!string.IsNullOrEmpty(包装资料.二盒装选用) && 剩余盒数 >= 2)
            {
                int 二盒装数量 = 剩余盒数 / 2;
                结果.Add((包装资料.二盒装选用, 二盒装数量, 2));
                剩余盒数 %= 2;
            }

            if (!string.IsNullOrEmpty(包装资料.单盒装选用) && 剩余盒数 > 0)
            {
                结果.Add((包装资料.单盒装选用, 剩余盒数, 1));
            }

            return 结果;
        }

    }


}