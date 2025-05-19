using OfficeOpenXml; // EPPlus的命名空间.
using OfficeOpenXml.Style;
using Sunny.UI;
using System.Text; // 添加这行用于StringBuilder
using System.Text.RegularExpressions;

namespace 包装计算
{
    public partial class Form1
    {
        private string 带销售数量型号 = "";
        private string 当前处理文件名 = "";
        private Dictionary<string, 包装数据> 包装数据字典;
        private string 当前选择算法 = "常规三选一最优方案"; // 默认值
        private Dictionary<string, List<string>> 标签码1分组字典 = new Dictionary<string, List<string>>();

        public class 包装数据
        {
            // 基本信息
            public string 包装名称 { get; set; }

            public string 物料码 { get; set; }
            public string 包装类型 { get; set; }
            public string 包装种类 { get; set; }

            // 包装内容物
            public string POF热缩袋 { get; set; }

            public string POF热缩袋_系统尺寸 { get; set; }
            public string 全棉织带 { get; set; }
            public string 全棉织带_系统尺寸 { get; set; }

            // 可用外箱信息
            public string 单盒装外箱 { get; set; }

            public string 单盒装尺寸 { get; set; }
            public string 二盒装外箱 { get; set; }
            public string 二盒装尺寸 { get; set; }
            public string 三盒装外箱 { get; set; }
            public string 三盒装尺寸 { get; set; }
            public string 五盒装外箱 { get; set; }
            public string 五盒装尺寸 { get; set; }

            // 添加通口箱相关属性
            public string 通口箱 { get; set; }

            public string 通口箱尺寸 { get; set; }

            // 产品型号信息列表
            public List<产品型号信息> 支持产品型号列表 { get; set; } = new List<产品型号信息>();
        }

        public class 产品型号信息
        {
            public string 产品型号 { get; set; }
            public string 支持的盒装外箱 { get; set; }
            public double 灯带最大长度 { get; set; }
            public double 线材最大长度 { get; set; }
            public double 最低阈值 { get; set; }
            public string 装箱产品型号规格 { get; set; }
        }

        public class 包装方案
        {
            public List<(string 包装名称, string 物料码, List<(string 序号, double 灯带长度, double 线材长度, string 标签码1, string 标签码2, string 标签码3, string 标签码4, string 客户型号, string 标签显示长度)> 包含灯带)> 多条包装列表 { get; set; }
                = new List<(string 包装名称, string 物料码, List<(string 序号, double 灯带长度, double 线材长度, string 标签码1, string 标签码2, string 标签码3, string 标签码4, string 客户型号, string 标签显示长度)>)>();

            public List<(string 包装名称, string 物料码, string 序号, double 灯带长度, double 线材长度, string 标签码1, string 标签码2, string 标签码3, string 标签码4, string 客户型号, string 标签显示长度)> 单条包装列表 { get; set; }
                = new List<(string 包装名称, string 物料码, string 序号, double 灯带长度, double 线材长度, string 标签码1, string 标签码2, string 标签码3, string 标签码4, string 客户型号, string 标签显示长度)>();

            // 添加新属性
            public string 方案名称 { get; set; } = "默认方案";

            public double 空间利用率 { get; set; }
            public double 评分 { get; set; }

            // 计算总条数
            public int 总条数
            {
                get
                {
                    int 多条总数 = 多条包装列表.Sum(x => x.包含灯带.Count);
                    int 单条总数 = 单条包装列表.Count;
                    return 多条总数 + 单条总数;
                }
            }

            // 计算盒子数量
            public int 盒子数量
            {
                get
                {
                    return 多条包装列表.Count + 单条包装列表.Count;
                }
            }

            public override string ToString()
            {
                return $"{方案名称} - {盒子数量}盒 - 利用率:{空间利用率:P1}";
            }
        }

        /// <summary>
        /// 加载包装数据
        /// </summary>
        private Dictionary<string, 包装数据> 加载包装数据()
        {
            Dictionary<string, 包装数据> 包装数据字典 = new Dictionary<string, 包装数据>();

            try
            {
                string 包装清单路径 = Path.Combine(Application.StartupPath, "新包装资料", "包装清单数据.xlsx");

                if (!File.Exists(包装清单路径))
                {
                    MessageBox.Show($"包装清单文件不存在: {包装清单路径}", "错误");
                    return 包装数据字典;
                }

                using (ExcelPackage package = new ExcelPackage(new FileInfo(包装清单路径)))
                {
                    // 遍历每个工作表（每个工作表代表一种包装）
                    foreach (var worksheet in package.Workbook.Worksheets)
                    {
                        if (worksheet.Dimension == null) continue;

                        // 创建新的包装数据对象
                        包装数据 包装 = new 包装数据
                        {
                            // 读取基本信息（第2行）
                            包装名称 = worksheet.Cells[2, 1].Text?.Trim(),
                            物料码 = worksheet.Cells[2, 2].Text?.Trim(),
                            包装类型 = worksheet.Cells[2, 3].Text?.Trim(),
                            包装种类 = worksheet.Cells[2, 4].Text?.Trim(),

                            // 读取包装内容物信息
                            POF热缩袋 = worksheet.Cells[2, 5].Text?.Trim(),
                            POF热缩袋_系统尺寸 = worksheet.Cells[2, 6].Text?.Trim(),
                            全棉织带 = worksheet.Cells[2, 7].Text?.Trim(),
                            全棉织带_系统尺寸 = worksheet.Cells[2, 8].Text?.Trim(),

                            // 读取外箱信息
                            单盒装外箱 = worksheet.Cells[2, 9].Text?.Trim(),
                            单盒装尺寸 = worksheet.Cells[2, 10].Text?.Trim(),
                            二盒装外箱 = worksheet.Cells[2, 11].Text?.Trim(),
                            二盒装尺寸 = worksheet.Cells[2, 12].Text?.Trim(),
                            三盒装外箱 = worksheet.Cells[2, 13].Text?.Trim(),
                            三盒装尺寸 = worksheet.Cells[2, 14].Text?.Trim(),

                            // 添加通口箱信息读取
                            通口箱 = worksheet.Cells[2, 17].Text?.Trim(),  // 假设通口箱信息在第15列
                            通口箱尺寸 = worksheet.Cells[2, 18].Text?.Trim()  // 假设通口箱尺寸在第16列
                        };

                        // 查找产品型号信息起始行
                        int 产品型号起始行 = 0;
                        for (int row = 1; row <= worksheet.Dimension.End.Row; row++)
                        {
                            if (worksheet.Cells[row, 1].Text?.Trim() == "装箱产品型号")
                            {
                                产品型号起始行 = row;
                                break;
                            }
                        }

                        // 读取产品型号信息
                        if (产品型号起始行 > 0)
                        {
                            for (int row = 产品型号起始行 + 1; row <= worksheet.Dimension.End.Row; row++)
                            {
                                string 型号 = worksheet.Cells[row, 1].Text?.Trim();
                                if (string.IsNullOrEmpty(型号)) continue;

                                产品型号信息 产品信息 = new 产品型号信息
                                {
                                    产品型号 = 型号,
                                    支持的盒装外箱 = worksheet.Cells[row, 2].Text?.Trim(),
                                    灯带最大长度 = double.TryParse(worksheet.Cells[row, 3].Text?.Trim(), out double 灯带长度) ? 灯带长度 : 0,
                                    线材最大长度 = double.TryParse(worksheet.Cells[row, 4].Text?.Trim(), out double 线材长度) ? 线材长度 : 0,
                                    最低阈值 = double.TryParse(worksheet.Cells[row, 5].Text?.Trim(), out double 阈值) ? 阈值 : 0,
                                    装箱产品型号规格 = worksheet.Cells[row, 6].Text?.Trim()
                                };

                                包装.支持产品型号列表.Add(产品信息);
                            }
                        }

                        // 将包装数据添加到字典
                        if (!string.IsNullOrEmpty(包装.物料码))
                        {
                            包装数据字典[包装.物料码] = 包装;
                        }
                    }

                    //MessageBox.Show($"已加载 {包装数据字典.Count} 条包装数据", "加载成功");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"加载包装数据出错: {ex.Message}\n\n{ex.StackTrace}", "错误");
            }

            return 包装数据字典;
        }

        private void 获取订单数据(ExcelWorksheet worksheet, int startRow, int endRow, int 规格型号列, int 物料编码列, int 销售数量列, int 剪切长度列,
                                     ref string 当前型号, ref HashSet<string> 当前出线方式, ref string 物料编码, ref double 当前销售数量, ref string 当前客户型号, ref string 标签要求)
        {
            //MessageBox.Show(当前客户型号, "客户型号");
            string 线长 = "";
            // 获取主行信息
            var mainSpecCell = worksheet.Cells[startRow, 规格型号列];
            var mainMaterialCell = worksheet.Cells[startRow, 物料编码列];
            var mainSalesCell = worksheet.Cells[startRow, 销售数量列];

            // 动态获取备注列
            int 备注列 = 获取备注列索引(worksheet);

            List<double> 区间线长列表 = new List<double>();

            string 规格型号 = mainSpecCell.Text;
            物料编码 = mainMaterialCell.Text;
            double.TryParse(mainSalesCell.Text, out 当前销售数量);

            // 调试输出当前销售数量
            //MessageBox.Show($"当前销售数量: {当前销售数量}\n" +
            //                $"原始文本: {mainSalesCell.Text}\n" +
            //                $"规格型号: {规格型号}\n" +
            //                $"物料编码: {物料编码}",
            //                "调试信息");

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
            else if (规格型号.Contains("C-SFB-"))
            {
                var parts = 规格型号.Split(new[] { "C-SFB-", "C-FR-" }, StringSplitOptions.RemoveEmptyEntries);
                if (parts.Length >= 1)
                {
                    string 型号部分 = parts[0];
                    var match = System.Text.RegularExpressions.Regex.Match(型号部分, @"A\d+");
                    if (match.Success)
                    {
                        当前型号 = match.Value;
                    }
                    else
                    {
                        var match1 = System.Text.RegularExpressions.Regex.Match(型号部分, @"W\d+");
                        当前型号 = match1.Value;
                    }
                }
            }

            //MessageBox.Show($"带销售数量型号: {带销售数量型号}\n",
            //                "调试信息");

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
            List<double> 线长列表 = new List<double>(); // 使用 double 类型来存储线长数值

            // 处理该区间内的所有行以获取出线方式
            for (int row = startRow; row <= endRow; row++)
            {
                var specCell = worksheet.Cells[row, 规格型号列];
                if (specCell.Value == null) continue;

                string 当前行规格型号 = specCell.Text;

                // 检查是否包含配件行标识
                if (当前行规格型号.Contains("B-硅胶注塑式") || 当前行规格型号.Contains("B-双层注塑式") || 当前行规格型号.Contains("B-硅胶双层注塑式") || 当前行规格型号.Contains("B-注塑式"))
                {
                    // 检查常见的出线方式关键词
                    if (当前行规格型号.Contains("端部出线"))
                    {
                        基础出线方式.Add("端部出线");
                    }
                    if (当前行规格型号.Contains("侧部出线") || 当前行规格型号.Contains("侧面出线"))
                    {
                        基础出线方式.Add("侧面出线");
                    }
                    if (当前行规格型号.Contains("底部出线"))
                    {
                        基础出线方式.Add("底部出线");
                    }

                    //获取无附件的线长长度
                    // 动态获取备注列的值
                    var remarkCell = worksheet.Cells[row, 备注列];
                    string 当前行备注 = remarkCell?.Text ?? "";
                    // 获取原始线长，可能有一个或两个
                    List<double> 原始线长列表 = 提取线长列表(当前行规格型号);
                    // 将多个线长相加
                    double 最终线长 = 原始线长列表.Sum(); // 将所有提取到的线长求和
                    // 检查备注列，如果有 "线长X米" 的内容，替换原始线长
                    double 备注线长 = 从备注中获取线长(当前行备注);
                    if (备注线长 > 0)
                    {
                        最终线长 = 备注线长;
                    }

                    区间线长列表.Add(最终线长);
                }
            }

            //标签要求
            string 文件夹路径 = Path.Combine(Application.StartupPath, "输出结果", 变量.订单编号);
            string 处理好订单文件夹路径 = Path.Combine(文件夹路径, "订单资料");
            if (!Directory.Exists(处理好订单文件夹路径))
            {
                Directory.CreateDirectory(处理好订单文件夹路径);
            }
            string 单位文件路径 = Path.Combine(处理好订单文件夹路径, "单位.txt");
            string 单位内容 = "m"; // 默认

            //MessageBox.Show(标签要求);

            if (!string.IsNullOrWhiteSpace(标签要求))
            {
                string t = 标签要求;

                // 1. 毫米
                if (t.Contains("毫米") || t.Contains("mm") || t.Contains("MM"))
                {
                    单位内容 = "mm";
                }
                // 2. m和Ft共存
                else if ((t.Contains("m") || t.Contains("M") || t.Contains("米制")) &&
                         (t.Contains("Ft") || t.Contains("FT") || t.Contains("ft") || t.Contains("英尺") || t.Contains("英制")))
                {
                    单位内容 = "m(Ft)";
                }
                // 3. m和IN共存
                else if ((t.Contains("m") || t.Contains("M")) &&
                         (t.Contains("IN") || t.Contains("in") || t.Contains("英寸")))
                {
                    单位内容 = "m(IN)";
                }
                // 4. 有Ft/英尺
                else if (t.Contains("Ft") || t.Contains("FT") || t.Contains("ft") || t.Contains("英尺"))
                {
                    单位内容 = "Ft";
                }
                // 5. 有IN/英寸
                else if (t.Contains("IN") || t.Contains("in") || t.Contains("英寸"))
                {
                    单位内容 = "IN";
                }
            }
            File.WriteAllText(单位文件路径, 单位内容);

            // 计算区间的总线长
            double 区间总线长 = 区间线长列表.Sum();
            线长 = 区间总线长.ToString();
            //MessageBox.Show($"最终线长：{区间总线长}", "错误");

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

            // 在处理订单数据之前，先创建一个字典来存储所有型号的数据
            Dictionary<string, List<(double 长度, int 数量, string 来源)>> 订单汇总数据 = new Dictionary<string, List<(double, int, string)>>();

            // 检查剪切长度
            if (剪切长度列 != -1)
            {
                string 剪切长度 = worksheet.Cells[startRow, 剪切长度列].Text?.Trim() ?? "";
                string 销售数量文本 = worksheet.Cells[startRow, 销售数量列].Text;
                double 当前行销售数量 = 0;
                double.TryParse(销售数量文本, out 当前行销售数量);

                // 确保型号在字典中存在
                if (!订单汇总数据.ContainsKey(当前型号))
                {
                    订单汇总数据[当前型号] = new List<(double, int, string)>();
                }

                // 处理剪切长度信息
                if (!string.IsNullOrEmpty(剪切长度))
                {
                    if (剪切长度.Contains("见附件"))
                    {
                        // 记录需要附件处理的型号信息
                        if (!变量.需要附件处理的型号.ContainsKey(当前型号))
                        {
                            变量.需要附件处理的型号[当前型号] = new List<(string 出线方式, double 销售数量)>();
                        }
                        // 使用当前出线方式而不是F列内容
                        变量.需要附件处理的型号[当前型号].Add((string.Join(",", 当前出线方式), 当前行销售数量));

                        // 创建匹配信息
                        var 匹配信息 = new 匹配信息(
                            变量.订单编号,
                            当前型号,
                            当前出线方式,
                            当前行销售数量,
                            "",
                            当前行销售数量,
                            0
                        );
                        变量.订单附件匹配列表.Add(匹配信息);

                        保存订单_附件型号Excel文件(当前型号, 当前出线方式, 当前行销售数量, 物料编码, 规格型号, true, 当前客户型号, "");

                        //// Debug输出
                        //MessageBox.Show($"添加匹配信息:\n" +
                        //               $"订单编号: {变量.订单编号}\n" +
                        //               $"型号: {当前型号}\n" +
                        //               $"出线方式: {string.Join(",", 当前出线方式)}\n" +
                        //               $"销售数量: {当前行销售数量}\n" +
                        //               $"当前匹配列表数量: {变量.订单附件匹配列表.Count}",
                        //               "匹配信息记录");
                    }
                    if (!剪切长度.Contains("见附件"))
                    {
                        // 处理直接包含长度信息的情况
                        string[] 长度组 = 剪切长度.Split(new[] { ',', '，', ';', '；', '+' }, StringSplitOptions.RemoveEmptyEntries);
                        List<(double 长度, int 数量)> 解析后的长度列表 = new List<(double, int)>();

                        foreach (string 单个长度 in 长度组)
                        {
                            var match = Regex.Match(单个长度.Trim(),
                                @"(\d+(?:\.\d+)?)\s*[Mm][^*]*\*\s*(\d+)\s*(?:PC|PCS|pc|pcs)",
                                RegexOptions.IgnoreCase);

                            if (match.Success)
                            {
                                double 长度 = double.Parse(match.Groups[1].Value);
                                int 数量 = int.Parse(match.Groups[2].Value);
                                订单汇总数据[当前型号].Add((长度, 数量, "直接长度"));
                                解析后的长度列表.Add((长度, 数量));

                                //#if DEBUG // 仅在调试模式下显示消息框
                                //                                MessageBox.Show($"解析分割后的长度:\n" +
                                //                                  $"型号: {当前型号}\n" +
                                //                                  $"长度: {长度}米\n" +
                                //                                  $"数量: {数量}个\n" +
                                //                                  $"原始文本: {单个长度}",
                                //                                  "处理结果");
                                //#endif
                            }
                        }

                        // 收集完所有长度信息后，只调用一次保存方法
                        if (解析后的长度列表.Any())
                        {
                            保存订单_备注型号Excel文件(当前型号, 当前出线方式, 当前行销售数量, 物料编码, 规格型号, 解析后的长度列表, 线长, 当前客户型号, "");
                        }
                    }
                }
                else if (当前行销售数量 > 0)
                {
                    // 处理只有销售数量的情况
                    订单汇总数据[当前型号].Add((当前行销售数量, 1, "销售数量"));
                }
            }

            // 在处理完所有数据后，统一生成Excel数据
            foreach (var 型号数据 in 订单汇总数据)
            {
                string 基础型号 = 型号数据.Key;
                var 数据列表 = 型号数据.Value;

                //StringBuilder sb = new StringBuilder();
                //sb.AppendLine($"型号: {基础型号}");
                //sb.AppendLine("数据列表内容:");

                //foreach (var (长度, 数量, 来源) in 数据列表)
                //{
                //    sb.AppendLine($"长度: {长度}, 数量: {数量}, 来源: {来源}");
                //}

                //// 如果有处理方式信息，也显示出来
                //if (变量.型号处理方式.TryGetValue(基础型号, out var 处理信息))
                //{
                //    var (处理方式, 重复次数) = 处理信息;
                //    sb.AppendLine($"处理方式: {处理方式}");
                //    sb.AppendLine($"重复次数: {重复次数}");
                //}

                //MessageBox.Show(sb.ToString(), "订单汇总数据");

                // 首先检查是否是需要附件处理的型号
                if (变量.需要附件处理的型号.ContainsKey(基础型号))
                {
                    //MessageBox.Show($"型号 {基础型号} 需要附件处理，跳过Excel生成", "提示");
                    continue; // 跳过需要附件处理的型号
                }

                // 检查是否已存在该型号的数据，如果存在则添加后缀
                string 当前处理型号 = 基础型号;
                int 后缀序号 = 1;
                while (变量.附件表数据.ContainsKey(当前处理型号))
                {
                    当前处理型号 = $"{基础型号}_{后缀序号++}";
                }

                // 使用新的型号标识创建数据列表
                变量.附件表数据[当前处理型号] = new List<string>();

                int 序号 = 1;
                double 总米数 = 0;
                bool 需要附件处理 = false;

                foreach (var (长度, 数量, 来源) in 数据列表)
                {
                    if (来源 == "附件")
                    {
                        // 标记需要等待附件处理
                        需要附件处理 = true;
                        总米数 = 长度; // 保存销售数量，供后续附件处理使用

                        // 添加匹配信息
                        var 匹配信息 = new 匹配信息(
                            变量.订单编号,
                            基础型号,
                            当前出线方式,
                            总米数,
                            当前处理型号,
                            总米数,
                            区间总线长);
                        变量.订单附件匹配列表.Add(匹配信息);

                        // Debug输出
                        MessageBox.Show($"型号 {当前处理型号} 需要从附件获取数据\n" +
                                      $"销售数量: {总米数}",
                                      "附件处理标记");

                        break; // 跳出循环，等待附件处理
                    }
                    else
                    {
                        // 处理直接长度和销售数量的情况
                        for (int i = 0; i < 数量; i++)
                        {
                            string 记录 = $"{序号++}, 1, {长度}, , , , , ";
                            变量.附件表数据[当前处理型号].Add(记录);
                            总米数 += 长度;
                        }
                    }
                }

                // 只处理非附件数据的总计和匹配信息
                if (!需要附件处理 && 总米数 > 0)
                {
                    变量.附件表数据[当前处理型号].Add($"总计, , {总米数:F3}, , , ");

                    var 匹配信息 = new 匹配信息(
                        变量.订单编号,
                        基础型号,
                        当前出线方式,
                        总米数,
                        当前处理型号,
                        总米数,
                        区间总线长);
                    变量.订单附件匹配列表.Add(匹配信息);
                }
            }

            // 如果没有剪切长度信息，使用销售数量
            if (当前销售数量 == 0 && double.TryParse(worksheet.Cells[startRow, 销售数量列].Text, out double 销售数量值))
            {
                当前销售数量 = 销售数量值;
            }

            // Debug输出
            //string debugInfo = $"处理区间：{startRow}-{endRow}\n" +
            //                  $"型号：{当前型号}\n" +
            //                  $"规格型号：{规格型号}\n" +
            //                  $"物料编码：{物料编码}\n" +
            //                  $"出线方式：{string.Join("，", 当前出线方式)}\n" +
            //                  $"销售数量：{当前销售数量}\n" +
            //                  $"区间总线长：{区间总线长}";

            //MessageBox.Show(debugInfo);
        }

        // 添加两个新方法来处理Excel文件保存
        private void 保存订单_备注型号Excel文件(string 型号, HashSet<string> 出线方式, double 销售数量, string 物料编码, string 规格型号, List<(double 长度, int 数量)> 长度列表, string 线材长度, string 客户型号, string 标签显示长度)
        {
            try
            {
                // 创建以订单编号为名的文件夹
                string 文件夹路径 = Path.Combine(Application.StartupPath, "输出结果", 变量.订单编号);
                string 处理好订单文件夹路径 = Path.Combine(文件夹路径, "订单资料");

                if (!Directory.Exists(处理好订单文件夹路径))
                {
                    Directory.CreateDirectory(处理好订单文件夹路径);
                }

                // 创建基础文件名
                string 基础文件名 = 型号 + "-" + 销售数量 + "-备注";

                // 检查是否已存在同名文件，并生成带有递增数字的文件名
                string 文件名 = 基础文件名;
                string 保存路径 = Path.Combine(处理好订单文件夹路径, 文件名 + ".xlsx");

                // 检查是否存在同型号的文件
                var 现有文件列表 = Directory.GetFiles(处理好订单文件夹路径, 型号 + "-备注*.xlsx");

                if (现有文件列表.Length > 0)
                {
                    // 找出最大的后缀数字
                    int 最大后缀 = 0;
                    foreach (var 文件 in 现有文件列表)
                    {
                        string 文件名称 = Path.GetFileNameWithoutExtension(文件);
                        // 尝试提取后缀数字
                        var 匹配 = Regex.Match(文件名称, 基础文件名 + @"-(\d+)$");
                        if (匹配.Success && int.TryParse(匹配.Groups[1].Value, out int 后缀))
                        {
                            最大后缀 = Math.Max(最大后缀, 后缀);
                        }
                    }

                    // 生成新的文件名，后缀数字加1
                    文件名 = $"{基础文件名}-{最大后缀 + 1}";
                    保存路径 = Path.Combine(处理好订单文件夹路径, 文件名 + ".xlsx");
                }
                else if (File.Exists(保存路径))
                {
                    // 如果基础文件名已存在但没有数字后缀的文件，添加-1后缀
                    文件名 = $"{基础文件名}-1";
                    保存路径 = Path.Combine(处理好订单文件夹路径, 文件名 + ".xlsx");
                }

                // 保存为Excel文件
                using (ExcelPackage package = new ExcelPackage())
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets.Add(型号);

                    // 添加标题行 - 按照新的格式要求
                    worksheet.Cells[1, 1].Value = "型号";
                    string 基础型号 = Regex.Replace(型号, @"-\d+(\.\d+)?$", ""); // 移除型号中可能的数字后缀
                    worksheet.Cells[1, 2].Value = 型号;
                    worksheet.Cells[2, 1].Value = "物料编码";
                    worksheet.Cells[2, 2].Value = 物料编码;
                    worksheet.Cells[3, 1].Value = "规格型号";
                    worksheet.Cells[3, 2].Value = 规格型号;
                    worksheet.Cells[4, 1].Value = "销售数量";
                    worksheet.Cells[4, 2].Value = 销售数量;
                    worksheet.Cells[5, 1].Value = "出线方式";
                    worksheet.Cells[5, 2].Value = 出线方式 != null ? string.Join("，", 出线方式) : "无";
                    worksheet.Cells[6, 1].Value = "客户型号";
                    worksheet.Cells[6, 2].Value = 客户型号;

                    // 添加数据表头
                    worksheet.Cells[7, 1].Value = "序号";
                    worksheet.Cells[7, 2].Value = "数量";
                    worksheet.Cells[7, 3].Value = "灯带长度";
                    worksheet.Cells[7, 4].Value = "线长长度";

                    // 设置表头样式
                    using (var range = worksheet.Cells[7, 1, 7, 4])
                    {
                        range.Style.Font.Bold = true;
                        range.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        range.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                        range.Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                        range.Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                        range.Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                        range.Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                    }

                    // 添加数据行
                    int row = 8;
                    int dataIndex = 1;
                    double 总长度 = 0;

                    foreach (var (长度, 数量) in 长度列表)
                    {
                        // 对于每个数量，添加一行
                        for (int i = 0; i < 数量; i++)
                        {
                            worksheet.Cells[row, 1].Value = dataIndex++;
                            worksheet.Cells[row, 2].Value = 1;
                            worksheet.Cells[row, 3].Value = 长度;
                            worksheet.Cells[row, 4].Value = 线材长度; // 线长长度

                            // 设置数据行样式
                            using (var range = worksheet.Cells[row, 1, row, 4])
                            {
                                range.Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                                range.Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                                range.Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                                range.Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            }

                            总长度 += 长度;
                            row++;
                        }
                    }

                    // 设置列宽
                    worksheet.Column(1).Width = 10;  // 序号
                    worksheet.Column(2).Width = 10;  // 数量
                    worksheet.Column(3).Width = 15;  // 灯带长度
                    worksheet.Column(4).Width = 15;  // 线长长度

                    // 保存文件
                    package.SaveAs(new FileInfo(保存路径));

                    // 在状态栏显示保存信息
                    uiTextBox_状态.AppendText($"已保存文件: {保存路径}" + Environment.NewLine);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"创建Excel文件时出错: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void 保存订单_附件型号Excel文件(string 型号, HashSet<string> 出线方式, double 销售数量, string 物料编码, string 规格型号, bool 是附件, string 客户型号, string 标签显示长度)
        {
            try
            {
                // 创建以订单编号为名的文件夹
                string 文件夹路径 = Path.Combine(Application.StartupPath, "输出结果", 变量.订单编号);
                string 处理好订单文件夹路径 = Path.Combine(文件夹路径, "订单资料");

                if (!Directory.Exists(处理好订单文件夹路径))
                {
                    Directory.CreateDirectory(处理好订单文件夹路径);
                }

                // 创建基础文件名
                string 基础文件名 = 型号 + "-" + 销售数量;
                // 如果是附件数据，添加标记
                if (是附件)
                {
                    基础文件名 += "-附件";
                }

                // 检查是否已存在同名文件，并生成带有递增数字的文件名
                string 文件名 = 基础文件名;
                string 保存路径 = Path.Combine(处理好订单文件夹路径, 文件名 + ".xlsx");

                // 检查是否存在同型号的文件
                var 现有文件列表 = Directory.GetFiles(处理好订单文件夹路径, 基础文件名 + "*.xlsx");

                if (现有文件列表.Length > 0)
                {
                    // 找出最大的后缀数字
                    int 最大后缀 = 0;
                    foreach (var 文件 in 现有文件列表)
                    {
                        string 文件名称 = Path.GetFileNameWithoutExtension(文件);
                        // 尝试提取后缀数字
                        var 匹配 = Regex.Match(文件名称, 基础文件名 + @"-(\d+)$");
                        if (匹配.Success && int.TryParse(匹配.Groups[1].Value, out int 后缀))
                        {
                            最大后缀 = Math.Max(最大后缀, 后缀);
                        }
                    }

                    // 生成新的文件名，后缀数字加1
                    文件名 = $"{基础文件名}-{最大后缀 + 1}";
                    保存路径 = Path.Combine(处理好订单文件夹路径, 文件名 + ".xlsx");
                }
                else if (File.Exists(保存路径))
                {
                    // 如果基础文件名已存在但没有数字后缀的文件，添加-1后缀
                    文件名 = $"{基础文件名}-1";
                    保存路径 = Path.Combine(处理好订单文件夹路径, 文件名 + ".xlsx");
                }

                // 保存为Excel文件
                using (ExcelPackage package = new ExcelPackage())
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets.Add(型号);

                    // 添加标题行 - 按照新的格式要求
                    worksheet.Cells[1, 1].Value = "型号";
                    string 基础型号 = Regex.Replace(型号, @"-\d+(\.\d+)?$", ""); // 移除型号中可能的数字后缀
                    worksheet.Cells[1, 2].Value = 型号;
                    worksheet.Cells[2, 1].Value = "物料编码";
                    worksheet.Cells[2, 2].Value = 物料编码;
                    worksheet.Cells[3, 1].Value = "规格型号";
                    worksheet.Cells[3, 2].Value = 规格型号;
                    worksheet.Cells[4, 1].Value = "销售数量";
                    worksheet.Cells[4, 2].Value = 销售数量;
                    worksheet.Cells[5, 1].Value = "出线方式";
                    worksheet.Cells[5, 2].Value = 出线方式 != null ? string.Join("，", 出线方式) : "无";
                    worksheet.Cells[6, 1].Value = "客户型号";
                    worksheet.Cells[6, 2].Value = 客户型号;

                    // 添加数据表头
                    worksheet.Cells[7, 1].Value = "序号";
                    worksheet.Cells[7, 2].Value = "数量";
                    worksheet.Cells[7, 3].Value = "灯带长度";
                    worksheet.Cells[7, 4].Value = "线长长度";
                    worksheet.Cells[7, 5].Value = "标签码1";  // 添加标签码列
                    worksheet.Cells[7, 6].Value = "标签码2";
                    worksheet.Cells[7, 7].Value = "标签码3";
                    worksheet.Cells[7, 8].Value = "标签码4";
                    worksheet.Cells[7, 9].Value = "标签显示长度";

                    // 设置表头样式
                    using (var range = worksheet.Cells[7, 1, 7, 9])
                    {
                        range.Style.Font.Bold = true;
                        range.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        range.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                        range.Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                        range.Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                        range.Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                        range.Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                    }

                    // 添加数据行
                    int row = 8;

                    if (是附件)
                    {
                        // 对于附件数据，添加一行数据
                        worksheet.Cells[row, 1].Value = "";
                        worksheet.Cells[row, 2].Value = "";
                        worksheet.Cells[row, 3].Value = "";
                        worksheet.Cells[row, 4].Value = ""; // 暂时留空
                        worksheet.Cells[row, 5].Value = "";
                        worksheet.Cells[row, 6].Value = "";
                        worksheet.Cells[row, 7].Value = "";
                        worksheet.Cells[row, 8].Value = "";
                        worksheet.Cells[row, 9].Value = "";

                        using (var range = worksheet.Cells[row, 1, row, 9])
                        {
                            range.Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            range.Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            range.Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            range.Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                        }

                        row++;
                    }

                    // 设置列宽
                    worksheet.Column(1).Width = 10;  // 序号
                    worksheet.Column(2).Width = 10;  // 数量
                    worksheet.Column(3).Width = 15;  // 灯带长度
                    worksheet.Column(4).Width = 15;  // 线长长度
                    worksheet.Column(5).Width = 20;
                    worksheet.Column(6).Width = 20;
                    worksheet.Column(7).Width = 20;
                    worksheet.Column(8).Width = 20;
                    worksheet.Column(9).Width = 20;

                    // 保存文件
                    package.SaveAs(new FileInfo(保存路径));

                    // 在状态栏显示保存信息
                    uiTextBox_状态.AppendText($"已保存文件: {保存路径}" + Environment.NewLine);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"创建Excel文件时出错: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private static HashSet<string> 已处理文件集合 = new HashSet<string>();

        private void 更新订单_附件型号Excel文件添加附件数据(string 型号, double 销售数量, string 工作表名称)
        {
            try
            {
                // 构建Excel文件路径
                string 文件夹路径 = Path.Combine(Application.StartupPath, "输出结果", 变量.订单编号);
                string 处理好订单文件夹路径 = Path.Combine(文件夹路径, "订单资料");

                // 基础文件名
                string 基础文件名 = 型号 + "-" + 销售数量 + "-附件";

                // 查找所有匹配的文件（包括带数字后缀的）
                var 匹配文件列表 = Directory.GetFiles(处理好订单文件夹路径, 基础文件名 + "*.xlsx")
                    .Where(f =>
                    {
                        string 文件名 = Path.GetFileNameWithoutExtension(f);
                        return 文件名 == 基础文件名 || Regex.IsMatch(文件名, 基础文件名 + @"-\d+$");
                    })
                    .ToList();

                if (匹配文件列表.Count == 0)
                {
                    string 文件列表信息 = "未找到文件: " + 基础文件名 + ".xlsx\n\n现有文件列表:\n";

                    var 文件列表 = Directory.GetFiles(处理好订单文件夹路径, "*.xlsx");
                    foreach (var 文件 in 文件列表)
                    {
                        文件列表信息 += Path.GetFileName(文件) + "\n";
                    }

                    MessageBox.Show(文件列表信息, "文件查找调试");

                    var 备选匹配文件列表 = 文件列表.Where(f => Path.GetFileNameWithoutExtension(f).StartsWith(型号)).ToList();
                    if (备选匹配文件列表.Count > 0)
                    {
                        匹配文件列表 = 备选匹配文件列表;
                        MessageBox.Show($"使用备选匹配文件: {string.Join(", ", 匹配文件列表.Select(Path.GetFileName))}", "文件匹配");
                    }
                    else
                    {
                        MessageBox.Show($"警告：未找到匹配的文件，无法更新附件数据", "错误");
                        return;
                    }
                }

                // 从附件Excel中读取数据并按序号首字母分组
                //Dictionary<string, List<(string 序号, string 数量, double 灯带长度, double 线长长度,string 标签码1,string 标签码2,string 标签码3,string 标签码4)>> 按序号分组的数据
                //    = new Dictionary<string, List<(string, string, double, double,string,string,string,string)>>();

                // 从附件工作表中提取数据
                List<(string 序号, string 数量, double 灯带长度, double 线长长度, string 标签码1, string 标签码2, string 标签码3, string 标签码4, string 标签显示长度)> 长度数据列表 = new List<(string, string, double, double, string, string, string, string, string)>();

                // 直接从附件Excel文件中读取数据
                string 附件文件路径 = 变量.附件excel地址;

                if (File.Exists(附件文件路径))
                {
                    using (ExcelPackage package = new ExcelPackage(new FileInfo(附件文件路径)))
                    {
                        // 查找指定的工作表
                        ExcelWorksheet worksheet = null;
                        foreach (var sheet in package.Workbook.Worksheets)
                        {
                            if (sheet.Name == 工作表名称)
                            {
                                worksheet = sheet;
                                break;
                            }
                        }

                        if (worksheet != null)
                        {
                            // 查找序号行
                            int 序号行号 = -1;
                            int 序号列号 = -1;
                            int 第一个长度列号 = -1;
                            int 数量列号 = -1;
                            int 灯带长度列号 = -1;
                            int 线长01列号 = -1;
                            int 线长02列号 = -1;

                            List<int> 标签码列号列表 = new List<int>();
                            int 标签码1列号 = -1;
                            int 标签码2列号 = -1;
                            int 标签码3列号 = -1;
                            int 标签码4列号 = -1;
                            int 标签显示长度列号 = -1;

                            // 在前20行中查找标题行
                            for (int row = 1; row <= Math.Min(20, worksheet.Dimension.End.Row); row++)
                            {
                                for (int col = 1; col <= worksheet.Dimension.End.Column; col++)
                                {
                                    string cellValue = worksheet.Cells[row, col].Text?.Trim() ?? "";
                                    if (cellValue.Contains("序号"))
                                    {
                                        序号行号 = row;
                                        序号列号 = col;
                                        break;
                                    }
                                }
                                if (序号行号 != -1) break;
                            }

                            if (序号行号 == -1)
                            {
                                // 如果找不到"序号"，尝试查找中文"序"字
                                for (int row = 1; row <= Math.Min(20, worksheet.Dimension.End.Row); row++)
                                {
                                    for (int col = 1; col <= worksheet.Dimension.End.Column; col++)
                                    {
                                        string cellValue = worksheet.Cells[row, col].Text?.Trim() ?? "";
                                        if (cellValue == "序")
                                        {
                                            序号行号 = row;
                                            序号列号 = col;
                                            break;
                                        }
                                    }
                                    if (序号行号 != -1) break;
                                }
                            }

                            if (序号行号 != -1)
                            {
                                // 在标题行查找其他列
                                for (int col = 1; col <= worksheet.Dimension.End.Column; col++)
                                {
                                    string cellValue = worksheet.Cells[序号行号, col].Text?.Trim() ?? "";

                                    if (cellValue.Contains("条数") || cellValue.Contains("数量"))
                                    {
                                        数量列号 = col;
                                    }
                                    else if (cellValue.Contains("总长度") || cellValue.Contains("灯带长度"))
                                    {
                                        灯带长度列号 = col;
                                    }
                                    else if (cellValue.Contains("01端线长"))
                                    {
                                        线长01列号 = col;
                                    }
                                    else if (cellValue.Contains("02端线长"))
                                    {
                                        线长02列号 = col;
                                    }
                                    else if (cellValue.Contains("标签显示长度(英寸)") || cellValue.Contains("标签显示长度(英尺)") || cellValue.Contains("英尺长度") || cellValue.Contains("英寸长度"))
                                    {
                                        标签显示长度列号 = col;
                                    }
                                }

                                // 获取工作表的实际列数
                                int 最大列数 = worksheet.Dimension.End.Column;

                                // 安全地遍历列
                                for (int col = 1; col <= 最大列数; col++)
                                {
                                    try
                                    {
                                        string cellValue = worksheet.Cells[序号行号, col].Text?.Trim() ?? "";

                                        if (cellValue.Contains("序号"))
                                        {
                                            序号列号 = col;
                                        }
                                        else if (cellValue.Contains("长度"))
                                        {
                                            第一个长度列号 = col;
                                            break;
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        MessageBox.Show($"处理第{col}列时出错: {ex.Message}", "错误");
                                        continue;
                                    }
                                }

                                // 获取标签码列

                                for (int col = 序号列号 + 1; col < 第一个长度列号; col++)
                                {
                                    if (col <= 最大列数)  // 确保不超出实际列范围
                                    {
                                        标签码列号列表.Add(col);
                                    }
                                }
                                //MessageBox.Show($"找到{标签码列号列表.Count}个标签码列", "调试信息");

                                // 直接使用标签码列号列表中的列号
                                int[] 标签码列号 = new int[4];

                                // 将找到的标签码列号复制到数组中
                                for (int i = 0; i < Math.Min(标签码列号列表.Count, 4); i++)
                                {
                                    标签码列号[i] = 标签码列号列表[i];
                                }

                                // 未找到的标签码列号设为0
                                for (int i = 标签码列号列表.Count; i < 4; i++)
                                {
                                    标签码列号[i] = 0;
                                }

                                // 将结果赋值给变量
                                标签码1列号 = 标签码列号[0];
                                标签码2列号 = 标签码列号[1];
                                标签码3列号 = 标签码列号[2];
                                标签码4列号 = 标签码列号[3];

                                // 打印每个列号的值进行确认
                                string 调试信息 = $"处理后的列号:\n" +
                                                 $"标签码1列号: {标签码列号[0]}\n" +
                                                 $"标签码2列号: {标签码列号[1]}\n" +
                                                 $"标签码3列号: {标签码列号[2]}\n" +
                                                 $"标签码4列号: {标签码列号[3]}";

                                //MessageBox.Show(调试信息, "列号分配结果");
                                //MessageBox.Show(标签码1列号.ToString(), 标签码2列号.ToString());

                                // 处理数据行
                                for (int row = 序号行号 + 1; row <= worksheet.Dimension.End.Row; row++)
                                {
                                    string 序号值 = worksheet.Cells[row, 序号列号].Text?.Trim() ?? "";

                                    // 跳过空行和总计行，剔除掉最后汇总行的判断
                                    if (string.IsNullOrEmpty(序号值) || 序号值.Contains("Grand Total") || 序号值.Contains("总计") || 序号值.Contains("合计"))
                                    {
                                        continue;
                                    }

                                    string 数量值 = 数量列号 != -1 ? worksheet.Cells[row, 数量列号].Text?.Trim() ?? "1" : "1";
                                    string 灯带长度值 = 灯带长度列号 != -1 ? worksheet.Cells[row, 灯带长度列号].Text?.Trim() ?? "0" : "0";

                                    string 标签显示长度值 = 标签显示长度列号 != -1 ? worksheet.Cells[row, 标签显示长度列号].Text?.Trim() ?? "0" : "0";

                                    // 获取01端线长和02端线长
                                    string 线长01值 = 线长01列号 != -1 ? worksheet.Cells[row, 线长01列号].Text?.Trim() ?? "0" : "0";
                                    string 线长02值 = 线长02列号 != -1 ? worksheet.Cells[row, 线长02列号].Text?.Trim() ?? "0" : "0";

                                    //获取标签码的值
                                    string 标签码1 = 标签码1列号 > 0 ? worksheet.Cells[row, 标签码1列号].Text?.Trim() ?? "" : "";
                                    string 标签码2 = 标签码2列号 > 0 ? worksheet.Cells[row, 标签码2列号].Text?.Trim() ?? "" : "";
                                    string 标签码3 = 标签码3列号 > 0 ? worksheet.Cells[row, 标签码3列号].Text?.Trim() ?? "" : "";
                                    string 标签码4 = 标签码4列号 > 0 ? worksheet.Cells[row, 标签码4列号].Text?.Trim() ?? "" : "";

                                    // 解析数值
                                    int 数量 = 1;
                                    double 灯带长度 = 0;
                                    double 线长01 = 0;
                                    double 线长02 = 0;

                                    int.TryParse(数量值, out 数量);
                                    double.TryParse(灯带长度值.Replace("m", "").Trim(), out 灯带长度);

                                    // 处理可能的"3m"格式
                                    if (线长01值.EndsWith("m", StringComparison.OrdinalIgnoreCase))
                                    {
                                        double.TryParse(线长01值.Replace("m", "").Trim(), out 线长01);
                                    }
                                    else
                                    {
                                        double.TryParse(线长01值, out 线长01);
                                    }

                                    if (线长02值.EndsWith("m", StringComparison.OrdinalIgnoreCase))
                                    {
                                        double.TryParse(线长02值.Replace("m", "").Trim(), out 线长02);
                                    }
                                    else if (线长02值.Contains("End Cap"))
                                    {
                                        线长02 = 0; // End Cap 不计算线长
                                    }
                                    else
                                    {
                                        double.TryParse(线长02值, out 线长02);
                                    }

                                    double 总线长 = 线长01 + 线长02;
                                    if (总线长 == 0) 总线长 = 3.0; // 默认值

                                    // 处理多条的情况
                                    if (数量 > 1)
                                    {
                                        double 分割后的灯带长度 = 灯带长度 / 数量;

                                        // 为每条创建一个单独的记录
                                        for (int i = 0; i < 数量; i++)
                                        {
                                            长度数据列表.Add((序号值 + "-" + (i + 1), "1", 分割后的灯带长度, 总线长, 标签码1, 标签码2, 标签码3, 标签码4, 标签显示长度值));
                                        }
                                    }
                                    else
                                    {
                                        长度数据列表.Add((序号值, 数量值, 灯带长度, 总线长, 标签码1, 标签码2, 标签码3, 标签码4, 标签显示长度值));
                                    }
                                }
                            }
                        }
                        else
                        {
                            MessageBox.Show($"未找到工作表: {工作表名称}", "工作表错误");
                        }
                    }
                }
                else
                {
                    MessageBox.Show($"附件文件不存在: {附件文件路径}", "文件错误");
                }

                // 更新所有匹配的Excel文件
                foreach (var 文件路径 in 匹配文件列表)
                {
                    using (ExcelPackage package = new ExcelPackage(new FileInfo(文件路径)))
                    {
                        if (package.Workbook.Worksheets.Count > 0)
                        {
                            ExcelWorksheet worksheet = package.Workbook.Worksheets[0]; // 获取第一个工作表

                            // 找到数据开始的行（通常是第8行，标题后）
                            int dataStartRow = 8;

                            // 清除现有数据行（如果有）
                            int lastRow = worksheet.Dimension.End.Row;
                            if (lastRow >= dataStartRow)
                            {
                                worksheet.DeleteRow(dataStartRow, lastRow - dataStartRow + 1);
                            }

                            // 添加新数据
                            int row = dataStartRow;
                            double 总灯带长度 = 0;

                            if (长度数据列表.Count > 0)
                            {
                                foreach (var (序号值, 数量值, 灯带长度, 线长长度, 标签码1, 标签码2, 标签码3, 标签码4, 标签显示长度值) in 长度数据列表)
                                {
                                    worksheet.Cells[row, 1].Value = 序号值;
                                    worksheet.Cells[row, 2].Value = 数量值;
                                    worksheet.Cells[row, 3].Value = 灯带长度;
                                    worksheet.Cells[row, 4].Value = 线长长度;
                                    worksheet.Cells[row, 5].Value = 标签码1;
                                    worksheet.Cells[row, 6].Value = 标签码2;
                                    worksheet.Cells[row, 7].Value = 标签码3;
                                    worksheet.Cells[row, 8].Value = 标签码4;
                                    worksheet.Cells[row, 9].Value = 标签显示长度值;

                                    using (var range = worksheet.Cells[row, 1, row, 9])
                                    {
                                        range.Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                                        range.Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                                        range.Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                                        range.Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                                    }

                                    总灯带长度 += 灯带长度;
                                    row++;
                                }
                            }
                            else
                            {
                                // 如果没有数据，添加一个空行
                                using (var range = worksheet.Cells[row, 1, row, 9])
                                {
                                    range.Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                                    range.Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                                    range.Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                                    range.Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                                }
                                row++;
                            }

                            // 保存文件
                            package.Save();

                            uiTextBox_状态.AppendText($"已更新文件: {Path.GetFileName(文件路径)}" + Environment.NewLine);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"更新Excel文件时出错: {ex.Message}\n\n{ex.StackTrace}", "错误");
            }
        }

        //处理同样型号，同样出线方式的备注订单
        private void 混合处理()
        {
            try
            {
                //MessageBox.Show("开始混合处理包装计算", "调试信息");

                string 文件夹路径 = Path.Combine(Application.StartupPath, "输出结果", 变量.订单编号);
                string 处理好订单文件夹路径 = Path.Combine(文件夹路径, "订单资料");

                if (!Directory.Exists(处理好订单文件夹路径))
                {
                    MessageBox.Show($"订单文件夹不存在: {处理好订单文件夹路径}", "错误");
                    return;
                }

                // 获取所有备注Excel文件
                var 所有文件 = Directory.GetFiles(处理好订单文件夹路径, "*.xlsx");
                var 备注文件列表 = 所有文件.Where(f => Path.GetFileNameWithoutExtension(f).EndsWith("-备注")).ToList();

                //MessageBox.Show($"找到 {备注文件列表.Count} 个备注Excel文件", "调试信息");

                // 用于存储已处理的型号
                Dictionary<string, List<string>> 型号文件映射 = new Dictionary<string, List<string>>();

                // 首先遍历备注文件，找出相同型号的文件，如果有客户型号按照相同客户型号合并，如果没有客户型号，按照原本型号来
                foreach (var 文件路径 in 备注文件列表)
                {
                    string 文件名 = Path.GetFileNameWithoutExtension(文件路径);
                    string 基础型号 = "";

                    // 提取型号信息
                    var 型号匹配 = Regex.Match(文件名, @"^([A-Z]\d+)");
                    if (型号匹配.Success)
                    {
                        基础型号 = 型号匹配.Groups[1].Value;

                        // 读取客户型号（假设在第6行第2列，如有不同请调整）
                        string 客户型号 = "";
                        using (ExcelPackage package = new ExcelPackage(new FileInfo(文件路径)))
                        {
                            var worksheet = package.Workbook.Worksheets[0];
                            客户型号 = worksheet.Cells[6, 2].Text?.Trim() ?? "";
                        }

                        // 分组key - 使用型号和客户型号组合，但不修改最终输出的型号名称
                        string 分组key = string.IsNullOrEmpty(客户型号) ? 基础型号 : $"{基础型号}|{客户型号}";

                        if (!型号文件映射.ContainsKey(分组key))
                        {
                            型号文件映射[分组key] = new List<string>();
                        }
                        型号文件映射[分组key].Add(文件路径);
                    }
                }

                // 处理相同型号的文件
                foreach (var kvp in 型号文件映射)
                {
                    string 分组key = kvp.Key;
                    var 相关文件列表 = kvp.Value;

                    // 从分组key中提取基础型号和客户型号
                    string 基础型号;
                    string 客户型号 = "";

                    if (分组key.Contains("|"))
                    {
                        var parts = 分组key.Split('|');
                        基础型号 = parts[0];
                        客户型号 = parts[1];
                    }
                    else
                    {
                        基础型号 = 分组key;
                    }

                    if (相关文件列表.Count > 1)
                    {
                        //MessageBox.Show($"发现型号 {基础型号} 有 {相关文件列表.Count} 个文件需要整合", "调试信息");

                        // 创建新的整合文件 - 只使用基础型号，不包含客户型号
                        string 整合文件名 = $"{基础型号}-整合.xlsx";
                        string 整合文件路径 = Path.Combine(处理好订单文件夹路径, 整合文件名);

                        double 总销售数量 = 0;
                        StringBuilder 销售数量表达式 = new StringBuilder();
                        List<(double 灯带长度, double 线材长度, int 数量)> 合并长度列表 = new List<(double, double, int)>();
                        string 出线方式 = "";
                        string 物料编码 = "";
                        string 规格型号 = "";

                        // 读取并合并所有相关文件的数据
                        foreach (var 文件路径 in 相关文件列表)
                        {
                            using (ExcelPackage package = new ExcelPackage(new FileInfo(文件路径)))
                            {
                                var worksheet = package.Workbook.Worksheets[0];

                                // 累加销售数量并构建表达式
                                if (worksheet.Cells[4, 2].Value != null)
                                {
                                    string 当前销售数量文本 = worksheet.Cells[4, 2].Text;
                                    double 当前销售数量;

                                    if (double.TryParse(当前销售数量文本, out 当前销售数量))
                                    {
                                        总销售数量 += 当前销售数量;

                                        // 添加到表达式
                                        if (销售数量表达式.Length > 0)
                                        {
                                            销售数量表达式.Append(" + ");
                                        }
                                        销售数量表达式.Append(当前销售数量文本);
                                    }
                                }

                                // 获取其他信息（使用第一个文件的信息）
                                if (string.IsNullOrEmpty(出线方式) && worksheet.Cells[5, 2].Value != null)
                                {
                                    出线方式 = worksheet.Cells[5, 2].Text;
                                }
                                if (string.IsNullOrEmpty(物料编码) && worksheet.Cells[2, 2].Value != null)
                                {
                                    物料编码 = worksheet.Cells[2, 2].Text;
                                }
                                if (string.IsNullOrEmpty(规格型号) && worksheet.Cells[3, 2].Value != null)
                                {
                                    规格型号 = worksheet.Cells[3, 2].Text;
                                }

                                // 读取长度数据
                                int dataStartRow = 8;
                                int lastRow = worksheet.Dimension.End.Row;

                                for (int row = dataStartRow; row <= lastRow; row++)
                                {
                                    // 检查B列是否为空
                                    string b列值 = worksheet.Cells[row, 2].Text?.Trim() ?? "";
                                    if (string.IsNullOrEmpty(b列值))
                                    {
                                        string a列值 = worksheet.Cells[row, 1].Text?.Trim() ?? "";
                                        if (!a列值.Contains("合计") && !a列值.Contains("总计"))
                                        {
                                            break; // 如果不是合计行，则结束数据读取
                                        }
                                        continue; // 如果是合计行，跳过这一行
                                    }

                                    if (worksheet.Cells[row, 3].Value != null)
                                    {
                                        double 灯带长度;
                                        double 线材长度 = 0;
                                        if (double.TryParse(worksheet.Cells[row, 3].Text, out 灯带长度))
                                        {
                                            int 数量 = 1;
                                            if (worksheet.Cells[row, 2].Value != null)
                                            {
                                                int.TryParse(worksheet.Cells[row, 2].Text, out 数量);
                                            }

                                            // 读取线材长度
                                            if (worksheet.Cells[row, 4].Value != null)
                                            {
                                                double.TryParse(worksheet.Cells[row, 4].Text, out 线材长度);
                                            }

                                            合并长度列表.Add((灯带长度, 线材长度, 数量));
                                        }
                                    }
                                }
                            }
                        }

                        // 创建整合后的Excel文件
                        using (ExcelPackage newPackage = new ExcelPackage())
                        {
                            var newWorksheet = newPackage.Workbook.Worksheets.Add(基础型号);

                            // 写入基本信息
                            newWorksheet.Cells[1, 1].Value = "型号";
                            newWorksheet.Cells[1, 2].Value = 基础型号;
                            newWorksheet.Cells[2, 1].Value = "物料编码";
                            newWorksheet.Cells[2, 2].Value = 物料编码;
                            newWorksheet.Cells[3, 1].Value = "规格型号";
                            newWorksheet.Cells[3, 2].Value = 规格型号;
                            newWorksheet.Cells[4, 1].Value = "销售数量";
                            newWorksheet.Cells[4, 2].Value = 销售数量表达式.ToString(); // 使用表达式而不是计算结果
                            newWorksheet.Cells[5, 1].Value = "出线方式";
                            newWorksheet.Cells[5, 2].Value = 出线方式;
                            newWorksheet.Cells[6, 1].Value = "客户型号"; // 添加客户型号行
                            newWorksheet.Cells[6, 2].Value = 客户型号;  // 保持客户型号不变

                            // 写入表头
                            newWorksheet.Cells[7, 1].Value = "序号";
                            newWorksheet.Cells[7, 2].Value = "数量";
                            newWorksheet.Cells[7, 3].Value = "灯带长度";
                            newWorksheet.Cells[7, 4].Value = "线长长度";

                            // 设置表头样式
                            using (var range = newWorksheet.Cells[7, 1, 7, 4])
                            {
                                range.Style.Font.Bold = true;
                                range.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                range.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                                range.Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                                range.Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                                range.Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                                range.Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            }

                            // 写入合并的数据
                            int currentRow = 8;
                            int 序号 = 1;
                            foreach (var (灯带长度, 线材长度, 数量) in 合并长度列表)
                            {
                                newWorksheet.Cells[currentRow, 1].Value = 序号++;
                                newWorksheet.Cells[currentRow, 2].Value = 数量;
                                newWorksheet.Cells[currentRow, 3].Value = 灯带长度;
                                newWorksheet.Cells[currentRow, 4].Value = 线材长度;

                                using (var range = newWorksheet.Cells[currentRow, 1, currentRow, 4])
                                {
                                    range.Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                                    range.Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                                    range.Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                                    range.Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                                }

                                currentRow++;
                            }

                            // 设置列宽
                            newWorksheet.Column(1).Width = 10;
                            newWorksheet.Column(2).Width = 10;
                            newWorksheet.Column(3).Width = 15;
                            newWorksheet.Column(4).Width = 15;

                            // 保存新文件
                            newPackage.SaveAs(new FileInfo(整合文件路径));
                        }

                        // 删除原文件
                        foreach (var 文件路径 in 相关文件列表)
                        {
                            File.Delete(文件路径);
                        }

                        //MessageBox.Show($"已整合型号 {基础型号} 的文件，销售数量表达式: {销售数量表达式}", "整合完成");
                    }
                }

                //MessageBox.Show("文件整合完成", "处理完成");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"混合处理出错: {ex.Message}\n\n{ex.StackTrace}", "错误");
            }
        }

        private void 查找合适包装_0(List<(string 序号, double 灯带长度, double 线材长度, string 标签码1, string 标签码2, string 标签码3, string 标签码4, string 客户型号, string 标签显示长度)> 长度列表, string 型号, string 出线方式)
        {
            try
            {
                // 检查包装数据是否已加载
                if (包装数据字典 == null || 包装数据字典.Count == 0)
                {
                    MessageBox.Show("包装数据未加载，正在加载...", "提示");
                    包装数据字典 = 加载包装数据();

                    if (包装数据字典.Count == 0)
                    {
                        MessageBox.Show("无法加载包装数据，测试终止", "错误");
                        return;
                    }
                }

                // 判断出线方式是否为空
                bool isNoWireType = string.IsNullOrWhiteSpace(出线方式);
                string[] 出线方式数组;
                if (isNoWireType)
                {
                    出线方式数组 = new[] { "直发" };
                }
                else
                {
                    出线方式数组 = 出线方式.Split(new[] { ',', '，', ';', '；' }, StringSplitOptions.RemoveEmptyEntries)
                                            .Select(s => s.Trim())
                                            .ToArray();
                }

                // 第一步：找到支持该型号和出线方式的所有包装
                List<包装数据> 匹配包装列表 = new List<包装数据>();
                foreach (var kvp in 包装数据字典)
                {
                    var 包装 = kvp.Value;
                    foreach (var 产品信息 in 包装.支持产品型号列表)
                    {
                        if (产品信息.产品型号 == 型号)
                        {
                            if (isNoWireType)
                            {
                                // 没有出线方式时，只有装箱产品型号规格为空或为“直发”才算匹配
                                if (string.IsNullOrWhiteSpace(产品信息.装箱产品型号规格) || 产品信息.装箱产品型号规格.Contains("直发"))
                                {
                                    匹配包装列表.Add(包装);
                                    break;
                                }
                            }
                            else
                            {
                                // 原有逻辑
                                bool 至少一个匹配 = false;
                                foreach (var 单个出线方式 in 出线方式数组)
                                {
                                    if (产品信息.装箱产品型号规格.Contains(单个出线方式))
                                    {
                                        至少一个匹配 = true;
                                        break;
                                    }
                                }
                                if (至少一个匹配)
                                {
                                    匹配包装列表.Add(包装);
                                    break;
                                }
                            }
                        }
                    }
                }

                if (匹配包装列表.Count == 0)
                {
                    // 提供更详细的错误信息
                    StringBuilder errorMessage = new StringBuilder();
                    errorMessage.AppendLine($"未找到支持型号 {型号} 且出线方式为以下任一项的包装:");

                    foreach (var 单个出线方式 in 出线方式数组)
                    {
                        errorMessage.AppendLine($"- {单个出线方式}");
                    }

                    // 显示所有可用的包装信息
                    errorMessage.AppendLine("\n可用的包装信息:");
                    foreach (var kvp in 包装数据字典)
                    {
                        var 包装 = kvp.Value;
                        foreach (var 产品信息 in 包装.支持产品型号列表)
                        {
                            if (产品信息.产品型号 == 型号)
                            {
                                errorMessage.AppendLine($"- 包装名称: {包装.包装名称}, 支持规格: {产品信息.装箱产品型号规格}");
                            }
                        }
                    }

                    MessageBox.Show(errorMessage.ToString(), "未找到匹配包装");
                    return;
                }

                // 创建第一步结果显示
                StringBuilder sb = new StringBuilder();

                sb.AppendLine($"找到 {匹配包装列表.Count} 个支持的包装方案:");
                foreach (var 包装 in 匹配包装列表)
                {
                    var 产品信息 = 包装.支持产品型号列表.FirstOrDefault(p => p.产品型号 == 型号);
                    if (产品信息 == null)
                    {
                        continue; // 跳过没有找到产品信息的包装
                    }
                    sb.AppendLine($"包装名称: {包装.包装名称}");
                    sb.AppendLine($"包装类型: {包装.包装类型}");
                    sb.AppendLine($"物料码: {包装.物料码}");
                    sb.AppendLine($"灯带最大长度: {产品信息.灯带最大长度:F2}m");
                    sb.AppendLine($"线材最大长度: {产品信息.线材最大长度:F2}m");
                    sb.AppendLine($"最低阈值: {产品信息.最低阈值:F2}m");
                    sb.AppendLine("----------------------------------------\n");
                }

                // 显示所有待处理灯带信息
                foreach (var 灯带 in 长度列表)
                {
                    sb.AppendLine($"序号 {灯带.序号}:"); // 修改这里，直接使用序号
                    sb.AppendLine($"灯带长度: {灯带.灯带长度:F2}m");
                    sb.AppendLine($"线材长度: {灯带.线材长度:F2}m");
                    sb.AppendLine();
                }

                // 计算总长度
                double 总灯带长度 = 长度列表.Sum(x => x.灯带长度);
                double 总线材长度 = 长度列表.Sum(x => x.线材长度);

                sb.AppendLine($"包装内总灯带长度: {总灯带长度:F2}m");
                sb.AppendLine($"包装内总线材长度: {总线材长度:F2}m");
                sb.AppendLine("\n----------------------------------------\n");

                // 创建结果窗口
                Form 结果窗口 = new Form
                {
                    Width = 800,
                    Height = 600,
                    Text = "第一步 - 匹配包装列表"
                };

                RichTextBox 文本框 = new RichTextBox
                {
                    Dock = DockStyle.Fill,
                    ReadOnly = true,
                    Font = new Font("Microsoft YaHei", 10),
                    Text = sb.ToString()
                };

                //结果窗口.Controls.Add(文本框);
                //结果窗口.ShowDialog();
                保存日志(sb.ToString(), "匹配到的包装列表", Path.GetFileName(当前处理文件名));

                // 执行两种方案

                var 方案一结果 = 执行方案一(匹配包装列表, 长度列表, 型号);
                var 方案二结果 = 执行方案二(匹配包装列表, 长度列表, 型号);
                var 节约方案结果 = 执行节约方案(匹配包装列表, 长度列表, 型号);
                //var RU客户方案结果=执行RU客户方案(匹配包装列表, 长度列表, 型号);
                //var Luminii客户方案结果=执行Luminii客户方案(匹配包装列表, 长度列表, 型号);

                //MessageBox.Show("当前选择的算法是：" + 当前选择算法, "算法选择结果", MessageBoxButtons.OK, MessageBoxIcon.Information);
                //显示Luminii客户方案结果(匹配包装列表, 长度列表, 型号);
                if (当前选择算法 == "常规三选一最优方案")
                {
                    显示方案比较结果_2(方案一结果, 方案二结果, 节约方案结果);
                }
                else if (当前选择算法 == "RU客户方案" || 变量.客户代码.Contains("12141") || 变量.客户代码.Contains("15065"))
                {
                    显示RU客户方案结果(匹配包装列表, 长度列表, 型号);
                }
                else if (当前选择算法 == "Luminii-常规")
                {
                    显示Luminii客户方案结果(匹配包装列表, 长度列表, 型号);
                }
                else if (当前选择算法 == "Luinii-不同标签码2不能混装")
                {
                    显示Luminii不同标签码2方案结果(匹配包装列表, 长度列表, 型号);
                }
                else if (当前选择算法 == "不同序号不可混装")
                {
                    显示不同序号不可混装方案结果(匹配包装列表, 长度列表, 型号);
                }
                else if (当前选择算法 == "不同标签码1不能混装")
                {
                    显示不同标签码1不能混装方案结果(匹配包装列表, 长度列表, 型号);
                }
                else if (当前选择算法 == "不同标签码2不能混装")
                {
                    显示不同标签码2不能混装方案结果(匹配包装列表, 长度列表, 型号);
                }
                else if (当前选择算法 == "不同标签码1不能混装(且外箱不能混装)")
                {
                    显示不同标签码1不能混装方案结果(匹配包装列表, 长度列表, 型号);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"查找包装出错: {ex.Message}\n\n{ex.StackTrace}", "错误");
            }
        }

        //先选出单条包装的灯带，在把剩余的灯带进行多条包装选择。
        private (List<(string 包装名称, string 物料码, List<(string 序号, double 灯带长度, double 线材长度, string 标签码1, string 标签码2, string 标签码3, string 标签码4, string 客户型号, string 标签显示长度)> 包含灯带)> 多条包装列表,
         List<(string 包装名称, string 物料码, string 序号, double 灯带长度, double 线材长度, string 标签码1, string 标签码2, string 标签码3, string 标签码4, string 客户型号, string 标签显示长度)> 单条包装列表)
执行方案一(List<包装数据> 匹配包装列表, List<(string 序号, double 灯带长度, double 线材长度, string 标签码1, string 标签码2, string 标签码3, string 标签码4, string 客户型号, string 标签显示长度)> 长度列表, string 型号)
        {
            // 添加辅助方法：判断是否需要计算线长
            bool 需要计算线长(List<(string 序号, double 灯带长度, double 线材长度, string 标签码1, string 标签码2, string 标签码3, string 标签码4, string 客户型号, string 标签显示长度)> 灯带列表)
            {
                if (灯带列表.Count > 4) return true;  // 超过4条需要计算
                return 灯带列表.Any(x => x.线材长度 > 5.0);  // 有任何一条超过5米需要计算
            }

            // 排除Luminii包装
            var 非Luminii包装列表 = 匹配包装列表.Where(p => !p.包装名称.Contains("Luminii")).ToList();

            var 单条包装列表 = 非Luminii包装列表.Where(p => p.包装类型.Contains("单条"))
                .OrderBy(p =>
                {
                    if (p.包装名称.Contains("330")) return 1;
                    if (p.包装名称.Contains("470")) return 2;
                    if (p.包装名称.Contains("600")) return 3;
                    return 4;
                })
                .ToList();

            List<(string 包装名称, string 物料码, string 序号, double 灯带长度, double 线材长度, string 标签码1, string 标签码2, string 标签码3, string 标签码4, string 客户型号, string 标签显示长度)> 单条包装列表记录
                = new List<(string, string, string, double, double, string, string, string, string, string, string)>();
            List<string> 需要多条包装的灯带序号 = new List<string>();

            // 单条包装匹配逻辑
            for (int i = 0; i < 长度列表.Count; i++)
            {
                var (序号, 灯带长度, 线材长度, 标签码1, 标签码2, 标签码3, 标签码4, 客户型号, 标签显示长度) = 长度列表[i];
                bool 找到单条包装 = false;

                foreach (var 包装 in 单条包装列表)
                {
                    var 产品信息 = 包装.支持产品型号列表.FirstOrDefault(p => p.产品型号 == 型号);
                    // 创建单条灯带的列表用于检查线长计算条件
                    var 单条灯带列表 = new List<(string 序号, double 灯带长度, double 线材长度, string 标签码1, string 标签码2, string 标签码3, string 标签码4, string 客户型号, string 标签显示长度)>
            {
                (序号, 灯带长度, 线材长度, 标签码1, 标签码2, 标签码3, 标签码4,客户型号,标签显示长度)
            };

                    bool 忽略线长限制 = !需要计算线长(单条灯带列表);

                    if (灯带长度 <= 产品信息.灯带最大长度 &&
                        灯带长度 >= 产品信息.最低阈值 &&
                        (忽略线长限制 || 线材长度 <= 产品信息.线材最大长度))
                    {
                        单条包装列表记录.Add((包装.包装名称, 包装.物料码, 序号, 灯带长度, 线材长度, 标签码1, 标签码2, 标签码3, 标签码4, 客户型号, 标签显示长度));
                        找到单条包装 = true;
                        break;
                    }
                }

                if (!找到单条包装)
                {
                    需要多条包装的灯带序号.Add(序号);
                }
            }

            List<(string 包装名称, string 物料码, List<(string 序号, double 灯带长度, double 线材长度, string 标签码1, string 标签码2, string 标签码3, string 标签码4, string 客户型号, string 标签显示长度)> 包含灯带)> 使用的包装列表
                = new List<(string, string, List<(string, double, double, string, string, string, string, string, string)>)>();

            if (需要多条包装的灯带序号.Count > 0)
            {
                var 待处理灯带 = 需要多条包装的灯带序号
                    .Select(序号 =>
                    {
                        var 原始数据 = 长度列表.FirstOrDefault(x => x.序号 == 序号);
                        return (序号: 原始数据.序号, 灯带长度: 原始数据.灯带长度, 线材长度: 原始数据.线材长度, 标签码1: 原始数据.标签码1, 标签码2: 原始数据.标签码2, 标签码3: 原始数据.标签码3, 标签码4: 原始数据.标签码4, 客户型号: 原始数据.客户型号, 标签显示长度: 原始数据.标签显示长度);
                    })
                    .ToList();

                // 使用比例算法进行分组
                List<List<(string 序号, double 灯带长度, double 线材长度, string 标签码1, string 标签码2, string 标签码3, string 标签码4, string 客户型号, string 标签显示长度)>> 分组结果 = new List<List<(string, double, double, string, string, string, string, string, string)>>();
                var 当前组 = new List<(string 序号, double 灯带长度, double 线材长度, string 标签码1, string 标签码2, string 标签码3, string 标签码4, string 客户型号, string 标签显示长度)>();
                double current灯带米数 = 0;
                double current线材长度 = 0;
                double 权重值 = uiTextBox_权重值.Text.ToDouble();

                foreach (var 灯带 in 待处理灯带)
                {
                    // 计算预计累计值
                    double 预计累计灯带米数 = current灯带米数 + 灯带.灯带长度;
                    double 预计累计线材长度 = current线材长度 + 灯带.线材长度;

                    var 空结果 = (
                        多条包装列表: new List<(string 包装名称, string 物料码, List<(string 序号, double 灯带长度, double 线材长度, string 标签码1, string 标签码2, string 标签码3, string 标签码4, string 客户型号, string 标签显示长度)> 包含灯带)>(),
                        单条包装列表: new List<(string 包装名称, string 物料码, string 序号, double 灯带长度, double 线材长度, string 标签码1, string 标签码2, string 标签码3, string 标签码4, string 客户型号, string 标签显示长度)>()
                    );

                    var 匹配包装 = 非Luminii包装列表.Where(p => p.包装名称.Contains("470") && (p.包装类型.Contains("多条") || p.包装类型.Contains("2条") || p.包装类型.Contains("3条"))).ToList();

                    if (!匹配包装.Any())
                    {
                        MessageBox.Show("未找到470包装规格", "提示");
                        return 空结果;
                    }

                    var 当前包装 = 匹配包装[0];
                    var 匹配产品 = 当前包装.支持产品型号列表.Where(p => p.产品型号 == 型号).ToList();

                    if (!匹配产品.Any())
                    {
                        MessageBox.Show($"未找到型号 {型号} 的产品信息", "提示");
                        return 空结果;
                    }

                    var 产品信息 = 匹配产品[0];
                    double 最大灯带米数 = 产品信息.灯带最大长度;
                    double 最大线材长度 = 产品信息.线材最大长度;

                    // 检查是否需要计算线长限制
                    var 预计组 = new List<(string 序号, double 灯带长度, double 线材长度, string 标签码1, string 标签码2, string 标签码3, string 标签码4, string 客户型号, string 标签显示长度)>(当前组);
                    预计组.Add(灯带);
                    bool 忽略线长限制 = !需要计算线长(预计组);

                    // 计算权重系数
                    double 累计灯带权重系数 = 预计累计灯带米数 / 最大灯带米数;
                    double 累计线材权重系数 = 预计累计线材长度 / 最大线材长度;

                    // 调整最大限制
                    double 调整后最大线材长度 = 最大线材长度 * (1 - 累计灯带权重系数 * 权重值);
                    double 调整后最大灯带米数 = 最大灯带米数 * (1 - 累计线材权重系数 * 权重值);

                    // 检查是否超出限制
                    bool 超出灯带限制 = 预计累计灯带米数 > 调整后最大灯带米数;
                    bool 超出线材限制 = !忽略线长限制 && 预计累计线材长度 > 调整后最大线材长度;

                    if (超出灯带限制 || 超出线材限制)
                    {
                        if (当前组.Count > 0)
                        {
                            分组结果.Add(new List<(string, double, double, string, string, string, string, string, string)>(当前组));
                            当前组.Clear();
                        }
                        current灯带米数 = 0;
                        current线材长度 = 0;
                    }

                    当前组.Add(灯带);
                    current灯带米数 += 灯带.灯带长度;
                    current线材长度 += 灯带.线材长度;
                }

                if (当前组.Count > 0)
                {
                    分组结果.Add(new List<(string, double, double, string, string, string, string, string, string)>(当前组));
                }

                // 为每个分组选择合适的包装
                foreach (var 分组 in 分组结果)
                {
                    double 分组总灯带长度 = 分组.Sum(x => x.灯带长度);
                    double 分组总线材长度 = 分组.Sum(x => x.线材长度);
                    bool 忽略线长限制 = !需要计算线长(分组);

                    // 先尝试使用470包装
                    var 四七零包装列表 = 非Luminii包装列表.Where(p => p.包装名称.Contains("470") &&
                        (p.包装类型.Contains("多条") || p.包装类型.Contains("2条") || p.包装类型.Contains("3条")))
                        .ToList();

                    if (四七零包装列表.Any())
                    {
                        var 四七零包装 = 四七零包装列表.FirstOrDefault();
                        var 产品信息 = 四七零包装.支持产品型号列表.FirstOrDefault(p => p.产品型号 == 型号);

                        if (分组总灯带长度 <= 产品信息.灯带最大长度 &&
                            (忽略线长限制 || 分组总线材长度 <= 产品信息.线材最大长度))
                        {
                            使用的包装列表.Add((四七零包装.包装名称, 四七零包装.物料码, 分组));
                            continue;
                        }
                    }

                    // 如果470包装不合适，尝试600包装
                    var 六百包装列表 = 非Luminii包装列表.Where(p => p.包装名称.Contains("600") &&
                        (p.包装类型.Contains("多条") || p.包装类型.Contains("2条") || p.包装类型.Contains("3条")))
                        .ToList();

                    if (六百包装列表.Any())
                    {
                        var 六百包装 = 六百包装列表.FirstOrDefault();
                        var 产品信息 = 六百包装.支持产品型号列表.FirstOrDefault(p => p.产品型号 == 型号);

                        if (分组总灯带长度 <= 产品信息.灯带最大长度 &&
                            (忽略线长限制 || 分组总线材长度 <= 产品信息.线材最大长度))
                        {
                            使用的包装列表.Add((六百包装.包装名称, 六百包装.物料码, 分组));
                        }
                        else
                        {
                            throw new Exception("无法找到合适的包装方案处理当前分组");
                        }
                    }
                }
            }

            return (使用的包装列表, 单条包装列表记录);
        }

        private (List<(string 包装名称, string 物料码, List<(string 序号, double 灯带长度, double 线材长度, string 标签码1, string 标签码2, string 标签码3, string 标签码4, string 客户型号, string 标签显示长度)> 包含灯带)> 多条包装列表,
         List<(string 包装名称, string 物料码, string 序号, double 灯带长度, double 线材长度, string 标签码1, string 标签码2, string 标签码3, string 标签码4, string 客户型号, string 标签显示长度)> 单条包装列表)
执行方案二(List<包装数据> 匹配包装列表, List<(string 序号, double 灯带长度, double 线材长度, string 标签码1, string 标签码2, string 标签码3, string 标签码4, string 客户型号, string 标签显示长度)> 长度列表, string 型号)
        {
            // 排除Luminii包装
            var 非Luminii包装列表 = 匹配包装列表.Where(p => !p.包装名称.Contains("Luminii")).ToList();

            List<(string 包装名称, string 物料码, List<(string 序号, double 灯带长度, double 线材长度, string 标签码1, string 标签码2, string 标签码3, string 标签码4, string 客户型号, string 标签显示长度)> 包含灯带)> 使用的包装列表
                = new List<(string, string, List<(string, double, double, string, string, string, string, string, string)>)>();

            List<(string 包装名称, string 物料码, string 序号, double 灯带长度, double 线材长度, string 标签码1, string 标签码2, string 标签码3, string 标签码4, string 客户型号, string 标签显示长度)> 单条包装列表记录
                = new List<(string, string, string, double, double, string, string, string, string, string, string)>();

            // 复制一份长度列表，用于处理
            List<(string 序号, double 灯带长度, double 线材长度, string 标签码1, string 标签码2, string 标签码3, string 标签码4, string 客户型号, string 标签显示长度)> 待处理灯带 = new List<(string, double, double, string, string, string, string, string, string)>(长度列表);

            // 获取600和470的多条包装，从非Luminii包装列表中获取
            var 六百包装 = 非Luminii包装列表.FirstOrDefault(p => p.包装名称?.Contains("600") == true &&
                (p.包装类型?.Contains("多条") == true || p.包装类型?.Contains("2条") == true || p.包装类型?.Contains("3条") == true));

            var 四七零包装 = 非Luminii包装列表.FirstOrDefault(p => p.包装名称?.Contains("470") == true &&
                (p.包装类型?.Contains("多条") == true || p.包装类型?.Contains("2条") == true || p.包装类型?.Contains("3条") == true));

            double 权重值 = uiTextBox_权重值.Text.ToDouble();

            // 处理多条包装
            while (待处理灯带.Count >= 2)
            {
                // 尝试找出最多能放入一个包装的灯带
                List<(string 序号, double 灯带长度, double 线材长度, string 标签码1, string 标签码2, string 标签码3, string 标签码4, string 客户型号, string 标签显示长度)> 当前组 = new List<(string, double, double, string, string, string, string, string, string)>();
                double 当前组灯带总长 = 0;
                double 当前组线材总长 = 0;

                for (int i = 0; i < 待处理灯带.Count; i++)
                {
                    var 当前灯带 = 待处理灯带[i];
                    double 预计累计灯带米数 = 当前组灯带总长 + 当前灯带.灯带长度;
                    double 预计累计线材长度 = 当前组线材总长 + 当前灯带.线材长度;

                    // 创建临时组用于检查
                    var 临时组 = new List<(string 序号, double 灯带长度, double 线材长度, string 标签码1, string 标签码2, string 标签码3, string 标签码4, string 客户型号, string 标签显示长度)>(当前组);
                    临时组.Add(当前灯带);

                    bool 可以用600 = 六百包装 != null && 检查包装是否可用(六百包装, 型号, 预计累计灯带米数, 预计累计线材长度, 权重值, 临时组);
                    bool 可以用470 = 四七零包装 != null && 检查包装是否可用(四七零包装, 型号, 预计累计灯带米数, 预计累计线材长度, 权重值, 临时组);

                    if (可以用600)
                    {
                        当前组.Add(当前灯带);
                        当前组灯带总长 = 预计累计灯带米数;
                        当前组线材总长 = 预计累计线材长度;
                    }
                    else
                    {
                        break;
                    }
                }

                // 处理当前组的逻辑
                if (当前组.Count >= 2)
                {
                    bool 可以用470 = 四七零包装 != null && 检查包装是否可用(四七零包装, 型号, 当前组灯带总长, 当前组线材总长, 权重值, 当前组);

                    if (可以用470)
                    {
                        使用的包装列表.Add((四七零包装.包装名称, 四七零包装.物料码, 当前组));
                    }
                    else
                    {
                        使用的包装列表.Add((六百包装.包装名称, 六百包装.物料码, 当前组));
                    }

                    foreach (var 灯带 in 当前组)
                    {
                        待处理灯带.Remove(灯带);
                    }
                }
                else
                {
                    break;
                }
            }

            // 处理剩余的单条灯带
            foreach (var 灯带 in 待处理灯带.ToList())  // 使用ToList()创建副本以避免修改集合时的问题
            {
                var 单条灯带列表 = new List<(string 序号, double 灯带长度, double 线材长度, string 标签码1, string 标签码2, string 标签码3, string 标签码4, string 客户型号, string 标签显示长度)> { 灯带 };
                bool 忽略线长限制 = !需要计算线长(单条灯带列表);

                var 单条包装列表 = 非Luminii包装列表.Where(p => p.包装类型?.Contains("单条") == true)
                    .OrderBy(p =>
                    {
                        if (p.包装名称?.Contains("330") == true) return 1;
                        if (p.包装名称?.Contains("470") == true) return 2;
                        if (p.包装名称?.Contains("600") == true) return 3;
                        return 4;
                    })
                    .ToList();

                bool 找到单条包装 = false;

                foreach (var 包装 in 单条包装列表)
                {
                    var 产品信息 = 包装.支持产品型号列表?.FirstOrDefault(p => p.产品型号 == 型号);
                    if (产品信息 != null &&
                        灯带.灯带长度 <= 产品信息.灯带最大长度 &&
                        灯带.灯带长度 >= 产品信息.最低阈值 &&
                        (忽略线长限制 || 灯带.线材长度 <= 产品信息.线材最大长度))
                    {
                        单条包装列表记录.Add((包装.包装名称, 包装.物料码, 灯带.序号, 灯带.灯带长度, 灯带.线材长度, 灯带.标签码1, 灯带.标签码2, 灯带.标签码3, 灯带.标签码4, 灯带.客户型号, 灯带.标签显示长度));
                        找到单条包装 = true;
                        break;
                    }
                }

                // 如果找不到合适的单条包装，尝试使用多条包装
                if (!找到单条包装)
                {
                    bool 可以用470单条 = 四七零包装 != null && 检查包装是否可用(四七零包装, 型号, 灯带.灯带长度, 灯带.线材长度, 权重值, 单条灯带列表);

                    if (可以用470单条)
                    {
                        使用的包装列表.Add((四七零包装.包装名称, 四七零包装.物料码, 单条灯带列表));
                    }
                    else if (六百包装 != null)
                    {
                        bool 可以用600单条 = 检查包装是否可用(六百包装, 型号, 灯带.灯带长度, 灯带.线材长度, 权重值, 单条灯带列表);

                        if (可以用600单条)
                        {
                            使用的包装列表.Add((六百包装.包装名称, 六百包装.物料码, 单条灯带列表));
                        }
                        else
                        {
                            throw new Exception($"无法找到合适的包装方案处理灯带 {灯带.序号}");
                        }
                    }
                }
            }

            return (使用的包装列表, 单条包装列表记录);
        }

        //有权重系数
        private (List<(string 包装名称, string 物料码, List<(string 序号, double 灯带长度, double 线材长度, string 标签码1, string 标签码2, string 标签码3, string 标签码4, string 客户型号, string 标签显示长度)> 包含灯带)> 多条包装列表,
                 List<(string 包装名称, string 物料码, string 序号, double 灯带长度, double 线材长度, string 标签码1, string 标签码2, string 标签码3, string 标签码4, string 客户型号, string 标签显示长度)> 单条包装列表)
        执行节约方案(List<包装数据> 匹配包装列表, List<(string 序号, double 灯带长度, double 线材长度, string 标签码1, string 标签码2, string 标签码3, string 标签码4, string 客户型号, string 标签显示长度)> 长度列表, string 型号)
        {
            List<(string 包装名称, string 物料码, List<(string 序号, double 灯带长度, double 线材长度, string 标签码1, string 标签码2, string 标签码3, string 标签码4, string 客户型号, string 标签显示长度)> 包含灯带)> 使用的包装列表
                = new List<(string, string, List<(string, double, double, string, string, string, string, string, string)>)>();

            List<(string 序号, double 灯带长度, double 线材长度, string 标签码1, string 标签码2, string 标签码3, string 标签码4, string 客户型号, string 标签显示长度)> 待处理灯带 = 长度列表.ToList(); // 保持原始顺序

            List<(string 包装名称, string 物料码, string 序号, double 灯带长度, double 线材长度, string 标签码1, string 标签码2, string 标签码3, string 标签码4, string 客户型号, string 标签显示长度)> 单条包装列表记录
                = new List<(string, string, string, double, double, string, string, string, string, string, string)>();

            double 权重值 = uiTextBox_权重值.Text.ToDouble(); // 可以调整这个值来改变限制的严格程度

            // 首先排除Luminii包装
            var 非Luminii包装列表 = 匹配包装列表.Where(p => !p.包装名称.Contains("Luminii")).ToList();

            // 获取470和600的多条包装，从非Luminii包装列表中获取
            var 四七零包装 = 非Luminii包装列表.FirstOrDefault(p => p.包装名称?.Contains("470") == true &&
                (p.包装类型?.Contains("多条") == true || p.包装类型?.Contains("2条") == true || p.包装类型?.Contains("3条") == true));

            var 六百包装 = 非Luminii包装列表.FirstOrDefault(p => p.包装名称?.Contains("600") == true &&
                (p.包装类型?.Contains("多条") == true || p.包装类型?.Contains("2条") == true || p.包装类型?.Contains("3条") == true));

            while (待处理灯带.Any())
            {
                var 当前组 = new List<(string 序号, double 灯带长度, double 线材长度, string 标签码1, string 标签码2, string 标签码3, string 标签码4, string 客户型号, string 标签显示长度)>();
                double current灯带米数 = 0;
                double current线材长度 = 0;

                // 处理第一条灯带
                var 第一个灯带 = 待处理灯带[0];
                当前组.Add(第一个灯带);
                current灯带米数 = 第一个灯带.灯带长度;
                current线材长度 = 第一个灯带.线材长度;
                待处理灯带.RemoveAt(0);

                // 尝试继续添加后续灯带
                while (待处理灯带.Any())
                {
                    var 下一个灯带 = 待处理灯带[0];
                    double 预计累计灯带米数 = current灯带米数 + 下一个灯带.灯带长度;
                    double 预计累计线材长度 = current线材长度 + 下一个灯带.线材长度;

                    // 检查470包装是否可用
                    bool 可以用470 = false;
                    if (四七零包装 != null)
                    {
                        var 临时组 = new List<(string 序号, double 灯带长度, double 线材长度, string 标签码1, string 标签码2, string 标签码3, string 标签码4, string 客户型号, string 标签显示长度)>(当前组);
                        临时组.Add(下一个灯带);
                        可以用470 = 检查包装是否可用(四七零包装, 型号, 预计累计灯带米数, 预计累计线材长度, 权重值, 临时组);
                    }

                    // 检查600包装是否可用
                    bool 可以用600 = false;
                    if (六百包装 != null)
                    {
                        var 临时组 = new List<(string 序号, double 灯带长度, double 线材长度, string 标签码1, string 标签码2, string 标签码3, string 标签码4, string 客户型号, string 标签显示长度)>(当前组);
                        临时组.Add(下一个灯带);
                        可以用600 = 检查包装是否可用(六百包装, 型号, 预计累计灯带米数, 预计累计线材长度, 权重值, 临时组);
                    }

                    // 只有当至少一种包装可用时，才添加到当前组
                    if (可以用470 || 可以用600)
                    {
                        当前组.Add(下一个灯带);
                        current灯带米数 = 预计累计灯带米数;
                        current线材长度 = 预计累计线材长度;
                        待处理灯带.RemoveAt(0);
                    }
                    else
                    {
                        break; // 当前包装已经装不下，退出循环
                    }
                }

                // 处理当前组
                if (当前组.Count == 1)
                {
                    // 尝试使用单条包装
                    var 单条包装列表 = 非Luminii包装列表
                        .Where(p => p.包装类型?.Contains("单条") == true)
                        .OrderBy(p =>
                        {
                            if (p.包装名称?.Contains("330") == true) return 1;
                            if (p.包装名称?.Contains("470") == true) return 2;
                            if (p.包装名称?.Contains("600") == true) return 3;
                            return 4;
                        })
                        .ToList();

                    bool 找到单条包装 = false;
                    foreach (var 包装 in 单条包装列表)
                    {
                        var 单条产品信息 = 包装.支持产品型号列表?.FirstOrDefault(p => p.产品型号 == 型号);
                        if (单条产品信息 == null) continue;

                        var 灯带 = 当前组[0];
                        var 单条灯带列表 = new List<(string 序号, double 灯带长度, double 线材长度, string 标签码1, string 标签码2, string 标签码3, string 标签码4, string 客户型号, string 标签显示长度)> { 灯带 };
                        bool 忽略线长限制 = !需要计算线长(单条灯带列表);

                        if (灯带.灯带长度 <= 单条产品信息.灯带最大长度 &&
                            灯带.灯带长度 >= 单条产品信息.最低阈值 &&
                            (忽略线长限制 || 灯带.线材长度 <= 单条产品信息.线材最大长度))
                        {
                            单条包装列表记录.Add((包装.包装名称, 包装.物料码, 灯带.序号, 灯带.灯带长度, 灯带.线材长度, 灯带.标签码1, 灯带.标签码2, 灯带.标签码3, 灯带.标签码4, 灯带.客户型号, 灯带.标签显示长度));
                            找到单条包装 = true;
                            break;
                        }
                    }

                    if (!找到单条包装)
                    {
                        // 如果找不到合适的单条包装，尝试使用多条包装
                        bool 可以用470单条 = false;
                        if (四七零包装 != null)
                        {
                            var 单条灯带列表 = new List<(string 序号, double 灯带长度, double 线材长度, string 标签码1, string 标签码2, string 标签码3, string 标签码4, string 客户型号, string 标签显示长度)> { 当前组[0] };
                            可以用470单条 = 检查包装是否可用(四七零包装, 型号, 当前组[0].灯带长度, 当前组[0].线材长度, 权重值, 单条灯带列表);

                            if (可以用470单条)
                            {
                                使用的包装列表.Add((四七零包装.包装名称, 四七零包装.物料码, 当前组));
                            }
                        }

                        // 如果470不行，尝试600
                        if (!可以用470单条 && 六百包装 != null)
                        {
                            var 单条灯带列表 = new List<(string 序号, double 灯带长度, double 线材长度, string 标签码1, string 标签码2, string 标签码3, string 标签码4, string 客户型号, string 标签显示长度)> { 当前组[0] };
                            bool 可以用600单条 = 检查包装是否可用(六百包装, 型号, 当前组[0].灯带长度, 当前组[0].线材长度, 权重值, 单条灯带列表);

                            if (可以用600单条)
                            {
                                使用的包装列表.Add((六百包装.包装名称, 六百包装.物料码, 当前组));
                            }
                            else
                            {
                                throw new Exception($"无法找到合适的包装方案处理灯带 {当前组[0].序号}");
                            }
                        }
                    }
                }
                else
                {
                    // 多条包装处理
                    bool 可以用470多条 = false;
                    if (四七零包装 != null)
                    {
                        可以用470多条 = 检查包装是否可用(四七零包装, 型号, current灯带米数, current线材长度, 权重值, 当前组);
                    }

                    // 优先使用470包装
                    if (可以用470多条)
                    {
                        使用的包装列表.Add((四七零包装.包装名称, 四七零包装.物料码, 当前组));
                    }
                    else if (六百包装 != null)
                    {
                        bool 可以用600多条 = 检查包装是否可用(六百包装, 型号, current灯带米数, current线材长度, 权重值, 当前组);
                        if (可以用600多条)
                        {
                            使用的包装列表.Add((六百包装.包装名称, 六百包装.物料码, 当前组));
                        }
                        else
                        {
                            throw new Exception("无法找到合适的多条包装");
                        }
                    }
                    else
                    {
                        throw new Exception("无法找到合适的多条包装");
                    }
                }
            }

            return (使用的包装列表, 单条包装列表记录);
        }

        /// <summary>
        /// 执行UL客户方案 - 每条灯带单独包装，按照330-470-600顺序选择最合适的包装
        /// </summary>
        private (List<(string 包装名称, string 物料码, List<(string 序号, double 灯带长度, double 线材长度, string 标签码1, string 标签码2, string 标签码3, string 标签码4, string 客户型号, string 标签显示长度)> 包含灯带)> 多条包装列表,
         List<(string 包装名称, string 物料码, string 序号, double 灯带长度, double 线材长度, string 标签码1, string 标签码2, string 标签码3, string 标签码4, string 客户型号, string 标签显示长度)> 单条包装列表)
执行RU客户方案(List<包装数据> 匹配包装列表, List<(string 序号, double 灯带长度, double 线材长度, string 标签码1, string 标签码2, string 标签码3, string 标签码4, string 客户型号, string 标签显示长度)> 长度列表, string 型号)
        {
            // 排除Luminii包装
            var 非Luminii包装列表 = 匹配包装列表.Where(p => !p.包装名称.Contains("Luminii")).ToList();

            // 创建空的结果列表
            List<(string 包装名称, string 物料码, List<(string 序号, double 灯带长度, double 线材长度, string 标签码1, string 标签码2, string 标签码3, string 标签码4, string 客户型号, string 标签显示长度)> 包含灯带)> 多条包装列表 =
                new List<(string 包装名称, string 物料码, List<(string 序号, double 灯带长度, double 线材长度, string 标签码1, string 标签码2, string 标签码3, string 标签码4, string 客户型号, string 标签显示长度)> 包含灯带)>();

            List<(string 包装名称, string 物料码, string 序号, double 灯带长度, double 线材长度, string 标签码1, string 标签码2, string 标签码3, string 标签码4, string 客户型号, string 标签显示长度)> 单条包装列表 =
                new List<(string 包装名称, string 物料码, string 序号, double 灯带长度, double 线材长度, string 标签码1, string 标签码2, string 标签码3, string 标签码4, string 客户型号, string 标签显示长度)>();

            // 定义UL客户包装规格顺序
            List<int> 包装规格顺序 = new List<int> { 330, 470, 600 };

            // 按照包装名称中的数字大小排序匹配包装列表，使用非Luminii包装列表
            非Luminii包装列表 = 非Luminii包装列表.OrderBy(p =>
            {
                // 从包装名称中提取数字
                var match = Regex.Match(p.包装名称, @"\d+");
                if (match.Success && int.TryParse(match.Value, out int size))
                {
                    // 找到该尺寸在顺序列表中的位置
                    int index = 包装规格顺序.IndexOf(size);
                    return index >= 0 ? index : int.MaxValue; // 如果不在列表中，放到最后
                }
                return int.MaxValue;
            }).ToList();

            // 创建待处理灯带的副本
            var 待处理灯带 = new List<(string 序号, double 灯带长度, double 线材长度, string 标签码1, string 标签码2, string 标签码3, string 标签码4, string 客户型号, string 标签显示长度)>(长度列表);

            // 为每条灯带选择最合适的包装
            while (待处理灯带.Count > 0)
            {
                var 当前灯带 = 待处理灯带[0];
                待处理灯带.RemoveAt(0);

                bool 已找到包装 = false;

                // 按照优先顺序尝试每种包装，使用非Luminii包装列表
                foreach (var 包装 in 非Luminii包装列表)
                {
                    // 查找该型号的产品信息
                    var 产品信息 = 包装.支持产品型号列表?.FirstOrDefault(p => p.产品型号 == 型号);
                    if (产品信息 == null) continue;

                    // 检查灯带长度和线材长度是否符合要求
                    if (当前灯带.灯带长度 <= 产品信息.灯带最大长度 && 当前灯带.线材长度 <= 产品信息.线材最大长度)
                    {
                        // 找到合适的包装，添加到单条包装列表
                        单条包装列表.Add((包装.包装名称, 包装.物料码, 当前灯带.序号, 当前灯带.灯带长度, 当前灯带.线材长度, 当前灯带.标签码1, 当前灯带.标签码2, 当前灯带.标签码3, 当前灯带.标签码4, 当前灯带.客户型号, 当前灯带.标签显示长度));
                        已找到包装 = true;
                        break;
                    }
                }

                // 如果没有找到合适的包装，使用最大的包装，从非Luminii包装列表中选择
                if (!已找到包装 && 非Luminii包装列表.Count > 0)
                {
                    var 最大包装 = 非Luminii包装列表.Last();
                    var 产品信息 = 最大包装.支持产品型号列表?.FirstOrDefault(p => p.产品型号 == 型号);

                    if (产品信息 != null)
                    {
                        单条包装列表.Add((最大包装.包装名称, 最大包装.物料码, 当前灯带.序号, 当前灯带.灯带长度, 当前灯带.线材长度, 当前灯带.标签码1, 当前灯带.标签码2, 当前灯带.标签码3, 当前灯带.标签码4, 当前灯带.客户型号, 当前灯带.标签显示长度));
                    }
                    else
                    {
                        // 如果没有找到支持该型号的包装，记录错误
                        Console.WriteLine($"错误：找不到支持型号 {型号} 的非Luminii包装，灯带序号：{当前灯带.序号}");
                    }
                }
            }

            return (多条包装列表, 单条包装列表);
        }

        /// <summary>
        /// 显示UL客户方案结果
        /// </summary>
        private void 显示RU客户方案结果(List<包装数据> 匹配包装列表, List<(string 序号, double 灯带长度, double 线材长度, string 标签码1, string 标签码2, string 标签码3, string 标签码4, string 客户型号, string 标签显示长度)> 长度列表, string 型号)
        {
            // 执行UL客户方案
            var RU方案 = 执行RU客户方案(匹配包装列表, 长度列表, 型号);

            int 方案包装数 = RU方案.多条包装列表.Count + RU方案.单条包装列表.Count;

            StringBuilder 结果 = new StringBuilder();
            结果.AppendLine("UL客户包装方案结果：");
            结果.AppendLine($"需要包装数：{方案包装数}");
            结果.AppendLine($"  多条包装：{RU方案.多条包装列表.Count} 个");
            结果.AppendLine($"  单条包装：{RU方案.单条包装列表.Count} 个");
            结果.AppendLine();

            结果.AppendLine("----------------------------------------");
            结果.AppendLine("包装方案详情：");

            // 按照包装物料码分组统计
            var 包装统计 = RU方案.单条包装列表
                .GroupBy(x => x.物料码)
                .Select(g => new { 物料码 = g.Key, 数量 = g.Count(), 包装名称 = g.First().包装名称 })
                .OrderBy(x => x.包装名称)
                .ToList();

            结果.AppendLine("包装使用统计：");
            foreach (var 统计 in 包装统计)
            {
                结果.AppendLine($"  {统计.包装名称} (物料码: {统计.物料码}): {统计.数量} 个");
            }

            结果.AppendLine("\n单条包装详情：");
            foreach (var 包装 in RU方案.单条包装列表)
            {
                结果.AppendLine($"  序号: {包装.序号}, 灯带长度: {包装.灯带长度:F3}m, 线材长度: {包装.线材长度:F3}m, 包装: {包装.包装名称} (物料码: {包装.物料码})");
            }

            // 保存日志
            保存日志(结果.ToString(), "UL客户包装方案结果", Path.GetFileName(当前处理文件名));

            // 创建Excel文件(方案组合结果EXCEL)
            string 文件夹路径 = Path.Combine(Application.StartupPath, "输出结果", 变量.订单编号);
            if (!Directory.Exists(文件夹路径))
            {
                Directory.CreateDirectory(文件夹路径);
            }

            string 文件路径 = Path.Combine(文件夹路径, "RU_" + Path.GetFileName(当前处理文件名));
            using (ExcelPackage package = new ExcelPackage(new FileInfo(文件路径)))
            {
                int 包装序号 = 1;

                // 处理单条包装
                foreach (var 包装 in RU方案.单条包装列表)
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets.Add($"第{包装序号}盒");

                    // 设置表头
                    worksheet.Cells[1, 1].Value = "序号";
                    worksheet.Cells[1, 2].Value = "条数";
                    worksheet.Cells[1, 3].Value = "米数";
                    worksheet.Cells[1, 4].Value = "标签码1";
                    worksheet.Cells[1, 5].Value = "标签码2";
                    worksheet.Cells[1, 6].Value = "实际剪切长度";
                    worksheet.Cells[1, 7].Value = "实际剪切长度/米";
                    worksheet.Cells[1, 8].Value = "线长";
                    worksheet.Cells[1, 9].Value = "包装编码";
                    worksheet.Cells[1, 13].Value = "客户型号"; // 新增
                    worksheet.Cells[1, 14].Value = "标签显示长度"; // 新增

                    // 填充单条包装数据
                    worksheet.Cells[2, 1].Value = 包装.序号;  // 直接使用原始序号
                    worksheet.Cells[2, 2].Value = 1;
                    worksheet.Cells[2, 3].Value = Math.Round(包装.灯带长度, 3);
                    worksheet.Cells[2, 8].Value = Math.Round(包装.线材长度, 3);
                    worksheet.Cells[2, 9].Value = 包装.物料码;
                    worksheet.Cells[2, 13].Value = 包装.客户型号;
                    worksheet.Cells[2, 14].Value = 包装.标签显示长度;

                    包装序号++;
                }

                package.Save();
            }

            包装方案 方案对象 = new 包装方案
            {
                多条包装列表 = RU方案.多条包装列表,
                单条包装列表 = RU方案.单条包装列表
            };
            生成包装材料需求流转单_4(文件夹路径, "RU_" + 当前处理文件名, 方案对象);

            //MessageBox.Show($"RU客户方案已生成，共需要{方案包装数}个包装。", "RU客户方案");
        }

        private (List<(string 包装名称, string 物料码, List<(string 序号, double 灯带长度, double 线材长度, string 标签码1, string 标签码2, string 标签码3, string 标签码4, string 客户型号, string 标签显示长度)> 包含灯带)> 多条包装列表,
         List<(string 包装名称, string 物料码, string 序号, double 灯带长度, double 线材长度, string 标签码1, string 标签码2, string 标签码3, string 标签码4, string 客户型号, string 标签显示长度)> 单条包装列表)
执行Luminii客户方案(List<包装数据> 匹配包装列表, List<(string 序号, double 灯带长度, double 线材长度, string 标签码1, string 标签码2, string 标签码3, string 标签码4, string 客户型号, string 标签显示长度)> 长度列表, string 型号)
        {
            // 筛选 Luminii 的包装
            var Luminii包装列表 = 匹配包装列表.Where(p => p.包装名称.Contains("Luminii")).ToList();

            var 多条包装列表 = new List<(string 包装名称, string 物料码, List<(string 序号, double 灯带长度, double 线材长度, string 标签码1, string 标签码2, string 标签码3, string 标签码4, string 客户型号, string 标签显示长度)> 包含灯带)>();
            var 单条包装列表 = new List<(string 包装名称, string 物料码, string 序号, double 灯带长度, double 线材长度, string 标签码1, string 标签码2, string 标签码3, string 标签码4, string 客户型号, string 标签显示长度)>();

            // 获取Luminii多条包装
            var 多条包装 = Luminii包装列表.FirstOrDefault(p =>
                p.包装类型?.Contains("多条") == true ||
                p.包装类型?.Contains("2条") == true ||
                p.包装类型?.Contains("3条") == true);

            // 获取Luminii单条包装
            var 单条包装列表数据 = Luminii包装列表
                .Where(p => p.包装类型?.Contains("单条") == true)
                .OrderBy(p => p.包装名称)
                .ToList();

            if (多条包装 == null)
            {
                MessageBox.Show("未找到Luminii多条包装方案", "提示");
                return (多条包装列表, 单条包装列表);
            }

            double 权重值 = uiTextBox_权重值.Text.ToDouble();

            // 不同于标签码1不能混装方案，这里不按标签码分组
            // 直接处理所有灯带，因为不需要区分标签码
            var 待处理灯带 = new List<(string 序号, double 灯带长度, double 线材长度, string 标签码1, string 标签码2, string 标签码3, string 标签码4, string 客户型号, string 标签显示长度)>(长度列表);

            // 先尝试将所有灯带放在一起
            double 总灯带长度 = 待处理灯带.Sum(x => x.灯带长度);
            double 总线材长度 = 待处理灯带.Sum(x => x.线材长度);
            bool 忽略线长限制_整组 = !需要计算线长(待处理灯带);
            bool 整组放入一个包装 = false;

            // 检查多条包装是否可以容纳所有灯带
            var 产品信息 = 多条包装.支持产品型号列表?.FirstOrDefault(p => p.产品型号 == 型号);
            if (产品信息 != null)
            {
                bool 可以用多条包装 = 产品信息 != null &&
                                  总灯带长度 <= 产品信息.灯带最大长度 &&
                                  (忽略线长限制_整组 || 总线材长度 <= 产品信息.线材最大长度);

                Console.WriteLine($"检查是否可以将整组放入Luminii多条包装: 总长={总灯带长度}m, 最大长度={产品信息?.灯带最大长度}m, 可用={可以用多条包装}");

                if (可以用多条包装)
                {
                    多条包装列表.Add((多条包装.包装名称, 多条包装.物料码, 待处理灯带));
                    整组放入一个包装 = true;
                }
            }

            // 如果整组不能放入一个包装，则分批处理
            if (!整组放入一个包装)
            {
                while (待处理灯带.Count > 0)
                {
                    var 当前组 = new List<(string 序号, double 灯带长度, double 线材长度, string 标签码1, string 标签码2, string 标签码3, string 标签码4, string 客户型号, string 标签显示长度)>();
                    double current灯带米数 = 0;
                    double current线材长度 = 0;

                    // 处理第一条灯带
                    var 第一个灯带 = 待处理灯带[0];
                    当前组.Add(第一个灯带);
                    current灯带米数 = 第一个灯带.灯带长度;
                    current线材长度 = 第一个灯带.线材长度;
                    待处理灯带.RemoveAt(0);

                    // 尝试继续添加灯带 - 这里不同于标签码1方案，不需要检查标签码是否相同
                    int 最大迭代次数 = 1000;
                    int 当前迭代 = 0;

                    while (待处理灯带.Any() && 当前迭代 < 最大迭代次数)
                    {
                        当前迭代++;
                        var 下一个灯带 = 待处理灯带[0];

                        double 预计累计灯带米数 = current灯带米数 + 下一个灯带.灯带长度;
                        double 预计累计线材长度 = current线材长度 + 下一个灯带.线材长度;

                        var 当前灯带组 = 当前组.Concat(new[] { 下一个灯带 }).ToList();

                        // 检查多条包装是否可用
                        bool 忽略线长限制 = !需要计算线长(当前灯带组);
                        bool 可以用多条包装 = false;

                        if (产品信息 != null)
                        {
                            可以用多条包装 = 预计累计灯带米数 <= 产品信息.灯带最大长度 &&
                                             (忽略线长限制 || 预计累计线材长度 <= 产品信息.线材最大长度);
                        }

                        // 添加调试信息
                        string 调试信息 = $"尝试合并灯带 序号:{下一个灯带.序号}, 标签码1:{下一个灯带.标签码1}\n" +
                                       $"累计灯带长度:{预计累计灯带米数}m, 累计线材长度:{预计累计线材长度}m\n" +
                                       $"Luminii多条包装可用:{可以用多条包装}";

                        调试信息 += $"\nLuminii多条包装最大长度:{产品信息?.灯带最大长度}m, 最低阈值:{产品信息?.最低阈值}m";

                        // 使用日志输出而不是MessageBox，避免阻塞
                        Console.WriteLine(调试信息);

                        if (可以用多条包装)
                        {
                            当前组.Add(下一个灯带);
                            current灯带米数 = 预计累计灯带米数;
                            current线材长度 = 预计累计线材长度;
                            待处理灯带.RemoveAt(0);
                            Console.WriteLine($"添加灯带成功，当前组数量: {当前组.Count}, 剩余待处理: {待处理灯带.Count}");
                        }
                        else
                        {
                            Console.WriteLine($"无法添加灯带 {下一个灯带.序号} 到当前组");
                            break;
                        }
                    }

                    // 处理当前组
                    if (当前组.Count > 0)
                    {
                        bool 忽略线长限制 = !需要计算线长(当前组);

                        if (产品信息 != null)
                        {
                            bool 可以用多条包装 = current灯带米数 <= 产品信息.灯带最大长度 &&
                                               (忽略线长限制 || current线材长度 <= 产品信息.线材最大长度);

                            if (可以用多条包装)
                            {
                                多条包装列表.Add((多条包装.包装名称, 多条包装.物料码, 当前组));
                                continue;
                            }
                        }

                        // 检查是否可以使用最低阈值限制
                        if (产品信息 != null)
                        {
                            bool 可以用多条包装 = current灯带米数 <= 产品信息.灯带最大长度 &&
                                               current灯带米数 >= 产品信息.最低阈值 &&
                                               (忽略线长限制 || current线材长度 <= 产品信息.线材最大长度);

                            if (可以用多条包装)
                            {
                                多条包装列表.Add((多条包装.包装名称, 多条包装.物料码, 当前组));
                                continue;
                            }
                        }

                        // 如果只有一条灯带且多条包装不可用，尝试单条包装
                        if (当前组.Count == 1 && 单条包装列表数据.Any())
                        {
                            var 灯带 = 当前组[0];
                            bool 找到单条包装 = false;

                            foreach (var 包装 in 单条包装列表数据)
                            {
                                bool 忽略线长_单条 = !需要计算线长(当前组);
                                var 单条产品信息 = 包装.支持产品型号列表?.FirstOrDefault(p => p.产品型号 == 型号);

                                if (单条产品信息 == null) continue;

                                if (灯带.灯带长度 <= 单条产品信息.灯带最大长度 &&
                                    灯带.灯带长度 >= 单条产品信息.最低阈值 &&
                                            (忽略线长_单条 || 灯带.线材长度 <= 单条产品信息.线材最大长度))
                                {
                                    单条包装列表.Add((包装.包装名称, 包装.物料码, 灯带.序号, 灯带.灯带长度, 灯带.线材长度, 灯带.标签码1, 灯带.标签码2, 灯带.标签码3, 灯带.标签码4, 灯带.客户型号, 灯带.标签显示长度));
                                    找到单条包装 = true;
                                    break;
                                }
                            }

                            // 如果没找到单条包装，强制使用多条包装
                            if (!找到单条包装)
                            {
                                多条包装列表.Add((多条包装.包装名称, 多条包装.物料码, 当前组));
                            }
                        }
                        else
                        {
                            // 多条灯带但无法使用多条包装，强制使用
                            多条包装列表.Add((多条包装.包装名称, 多条包装.物料码, 当前组));
                        }
                    }
                }
            }

            return (多条包装列表, 单条包装列表);
        }

        private void 显示Luminii客户方案结果(List<包装数据> 匹配包装列表, List<(string 序号, double 灯带长度, double 线材长度, string 标签码1, string 标签码2, string 标签码3, string 标签码4, string 客户型号, string 标签显示长度)> 长度列表, string 型号)
        {
            // 执行Luminii客户方案
            var Luminii方案 = 执行Luminii客户方案(匹配包装列表, 长度列表, 型号);

            int 方案包装数 = Luminii方案.多条包装列表.Count + Luminii方案.单条包装列表.Count;

            StringBuilder 结果 = new StringBuilder();
            结果.AppendLine("Luminii客户包装方案结果：");
            结果.AppendLine($"需要包装数：{方案包装数}");
            结果.AppendLine($"  多条包装：{Luminii方案.多条包装列表.Count} 个");
            结果.AppendLine($"  单条包装：{Luminii方案.单条包装列表.Count} 个");
            结果.AppendLine();

            // 显示多条包装详情
            if (Luminii方案.多条包装列表.Any())
            {
                结果.AppendLine("多条包装详情：");
                foreach (var 包装 in Luminii方案.多条包装列表)
                {
                    结果.AppendLine($"包装名称: {包装.包装名称}");
                    结果.AppendLine($"物料码: {包装.物料码}");
                    结果.AppendLine("包含灯带:");
                    double 包装内总灯带长度 = 0;
                    double 包装内总线材长度 = 0;

                    foreach (var 灯带 in 包装.包含灯带)
                    {
                        结果.AppendLine($"  序号: {灯带.序号}");
                        结果.AppendLine($"    灯带长度: {灯带.灯带长度:F2}m");
                        结果.AppendLine($"    线材长度: {灯带.线材长度:F2}m");
                        结果.AppendLine($"    标签码1: {灯带.标签码1}");
                        结果.AppendLine($"    标签码2: {灯带.标签码2}");
                        结果.AppendLine($"    标签码3: {灯带.标签码3}");
                        结果.AppendLine($"    标签码4: {灯带.标签码4}");
                        包装内总灯带长度 += 灯带.灯带长度;
                        包装内总线材长度 += 灯带.线材长度;
                    }
                    结果.AppendLine($"包装内总灯带长度: {包装内总灯带长度:F2}m");
                    结果.AppendLine($"包装内总线材长度: {包装内总线材长度:F2}m");
                    结果.AppendLine("----------------------------------------");
                }
            }

            // 显示单条包装详情
            if (Luminii方案.单条包装列表.Any())
            {
                结果.AppendLine("\n单条包装详情：");
                foreach (var (包装名称, 物料码, 序号, 灯带长度, 线材长度, 标签码1, 标签码2, 标签码3, 标签码4, 客户型号, 标签显示长度) in Luminii方案.单条包装列表)
                {
                    结果.AppendLine($"包装名称: {包装名称}");
                    结果.AppendLine($"物料码: {物料码}");
                    结果.AppendLine($"序号: {序号}");
                    结果.AppendLine($"  灯带长度: {灯带长度:F2}m");
                    结果.AppendLine($"  线材长度: {线材长度:F2}m");
                    结果.AppendLine($"  标签码1: {标签码1}");
                    结果.AppendLine($"  标签码2: {标签码2}");
                    结果.AppendLine($"  标签码3: {标签码3}");
                    结果.AppendLine($"  标签码4: {标签码4}");
                    结果.AppendLine("----------------------------------------");
                }
            }

            // 保存日志
            保存日志(结果.ToString(), "Luminii客户包装方案结果", Path.GetFileName(当前处理文件名));

            //MessageBox.Show("准备生成Luminii方案的Excel文件和流转单");

            // 创建Excel文件
            string 文件夹路径 = Path.Combine(Application.StartupPath, "输出结果", 变量.订单编号);
            if (!Directory.Exists(文件夹路径))
            {
                Directory.CreateDirectory(文件夹路径);
            }

            string 文件路径 = Path.Combine(文件夹路径, "Luminii_" + Path.GetFileName(当前处理文件名));
            using (ExcelPackage package = new ExcelPackage(new FileInfo(文件路径)))
            {
                int 包装序号 = 1;

                // 处理多条包装
                foreach (var 包装 in Luminii方案.多条包装列表)
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets.Add($"第{包装序号}盒");

                    // 设置表头
                    worksheet.Cells[1, 1].Value = "序号";
                    worksheet.Cells[1, 2].Value = "条数";
                    worksheet.Cells[1, 3].Value = "米数";
                    worksheet.Cells[1, 4].Value = "标签码1";
                    worksheet.Cells[1, 5].Value = "标签码2";
                    worksheet.Cells[1, 6].Value = "标签码3";
                    worksheet.Cells[1, 7].Value = "标签码4";
                    worksheet.Cells[1, 8].Value = "线长";
                    worksheet.Cells[1, 9].Value = "包装编码";
                    worksheet.Cells[1, 13].Value = "客户型号";
                    worksheet.Cells[1, 14].Value = "标签显示长度";

                    // 填充数据
                    int row = 2;
                    foreach (var 灯带 in 包装.包含灯带)
                    {
                        worksheet.Cells[row, 1].Value = 灯带.序号;
                        worksheet.Cells[row, 2].Value = 1;
                        worksheet.Cells[row, 3].Value = Math.Round(灯带.灯带长度, 3);
                        worksheet.Cells[row, 4].Value = 灯带.标签码1;
                        worksheet.Cells[row, 5].Value = 灯带.标签码2;
                        worksheet.Cells[row, 6].Value = 灯带.标签码3;
                        worksheet.Cells[row, 7].Value = 灯带.标签码4;
                        worksheet.Cells[row, 8].Value = Math.Round(灯带.线材长度, 3);
                        worksheet.Cells[row, 9].Value = 包装.物料码;
                        worksheet.Cells[row, 13].Value = 灯带.客户型号;
                        worksheet.Cells[row, 14].Value = 灯带.标签显示长度;

                        // 设置单元格样式
                        using (var range = worksheet.Cells[row, 1, row, 9])
                        {
                            range.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                            range.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                            range.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                            range.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        }

                        row++;
                    }

                    包装序号++;
                }

                // 处理单条包装
                foreach (var 包装 in Luminii方案.单条包装列表)
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets.Add($"第{包装序号}盒");

                    // 设置表头
                    worksheet.Cells[1, 1].Value = "序号";
                    worksheet.Cells[1, 2].Value = "条数";
                    worksheet.Cells[1, 3].Value = "米数";
                    worksheet.Cells[1, 4].Value = "标签码1";
                    worksheet.Cells[1, 5].Value = "标签码2";
                    worksheet.Cells[1, 6].Value = "标签码3";
                    worksheet.Cells[1, 7].Value = "标签码4";
                    worksheet.Cells[1, 8].Value = "线长";
                    worksheet.Cells[1, 9].Value = "包装编码";
                    worksheet.Cells[1, 13].Value = "客户型号";
                    worksheet.Cells[1, 14].Value = "标签显示长度";

                    // 填充单条包装数据
                    worksheet.Cells[2, 1].Value = 包装.序号;
                    worksheet.Cells[2, 2].Value = 1;
                    worksheet.Cells[2, 3].Value = Math.Round(包装.灯带长度, 3);
                    worksheet.Cells[2, 4].Value = 包装.标签码1;
                    worksheet.Cells[2, 5].Value = 包装.标签码2;
                    worksheet.Cells[2, 6].Value = 包装.标签码3;
                    worksheet.Cells[2, 7].Value = 包装.标签码4;
                    worksheet.Cells[2, 8].Value = Math.Round(包装.线材长度, 3);
                    worksheet.Cells[2, 9].Value = 包装.物料码;
                    worksheet.Cells[2, 13].Value = 包装.客户型号;
                    worksheet.Cells[2, 14].Value = 包装.标签显示长度;

                    // 设置单元格样式
                    using (var range = worksheet.Cells[2, 1, 2, 9])
                    {
                        range.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                        range.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                        range.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                        range.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    }

                    包装序号++;
                }

                package.Save();
            }

            包装方案 选定方案 = new 包装方案
            {
                多条包装列表 = Luminii方案.多条包装列表,
                单条包装列表 = Luminii方案.单条包装列表
            };

            // 生成包装材料需求流转单
            生成包装材料需求流转单_4(文件夹路径, "Luminii_" + 当前处理文件名, 选定方案);

            uiTextBox_状态.AppendText($"Luminii客户方案已生成，共需要{方案包装数}个包装。" + Environment.NewLine);
            uiTextBox_状态.AppendText("------------------------" + Environment.NewLine);
        }

        private (List<(string 包装名称, string 物料码, List<(string 序号, double 灯带长度, double 线材长度, string 标签码1, string 标签码2, string 标签码3, string 标签码4, string 客户型号, string 标签显示长度)> 包含灯带)> 多条包装列表,
         List<(string 包装名称, string 物料码, string 序号, double 灯带长度, double 线材长度, string 标签码1, string 标签码2, string 标签码3, string 标签码4, string 客户型号, string 标签显示长度)> 单条包装列表)
执行Luminii不同标签码2方案(List<包装数据> 匹配包装列表, List<(string 序号, double 灯带长度, double 线材长度, string 标签码1, string 标签码2, string 标签码3, string 标签码4, string 客户型号, string 标签显示长度)> 长度列表, string 型号)
        {
            // 添加辅助方法：判断是否需要计算线长
            bool 需要计算线长(List<(string 序号, double 灯带长度, double 线材长度, string 标签码1, string 标签码2, string 标签码3, string 标签码4, string 客户型号, string 标签显示长度)> 灯带列表)
            {
                if (灯带列表.Count > 4) return true;  // 超过4条需要计算
                return 灯带列表.Any(x => x.线材长度 > 5.0);  // 有任何一条超过5米需要计算
            }

            List<(string 包装名称, string 物料码, List<(string 序号, double 灯带长度, double 线材长度, string 标签码1, string 标签码2, string 标签码3, string 标签码4, string 客户型号, string 标签显示长度)> 包含灯带)> 使用的包装列表
                = new List<(string, string, List<(string, double, double, string, string, string, string, string, string)>)>();

            List<(string 包装名称, string 物料码, string 序号, double 灯带长度, double 线材长度, string 标签码1, string 标签码2, string 标签码3, string 标签码4, string 客户型号, string 标签显示长度)> 单条包装列表记录
                = new List<(string, string, string, double, double, string, string, string, string, string, string)>();

            double 权重值 = uiTextBox_权重值.Text.ToDouble();

            // 只选择Luminii包装
            var Luminii包装列表 = 匹配包装列表.Where(p => p.包装名称.Contains("Luminii")).ToList();

            if (!Luminii包装列表.Any())
            {
                throw new Exception("未找到Luminii包装");
            }

            // 获取600的多条包装
            var 六百包装 = Luminii包装列表.FirstOrDefault(p => p.包装种类?.Contains("600") == true &&
                (p.包装类型?.Contains("多条") == true || p.包装类型?.Contains("2条") == true || p.包装类型?.Contains("3条") == true));

            if (六百包装 == null)
            {
                throw new Exception("未找到Luminii 600包装");
            }

            // 按标签码2分组
            var 分组后的灯带 = 长度列表.GroupBy(x => x.标签码2 ?? "").ToList();

            foreach (var 组 in 分组后的灯带)
            {
                var 待处理灯带 = 组.ToList();
                var 当前组 = new List<(string 序号, double 灯带长度, double 线材长度, string 标签码1, string 标签码2, string 标签码3, string 标签码4, string 客户型号, string 标签显示长度)>();
                double current灯带米数 = 0;
                double current线材长度 = 0;

                // 获取产品信息
                var 产品信息 = 六百包装.支持产品型号列表?.FirstOrDefault(p => p.产品型号 == 型号);
                if (产品信息 == null)
                {
                    throw new Exception($"未找到型号 {型号} 的产品信息");
                }

                foreach (var 灯带 in 待处理灯带)
                {
                    // 判断添加这条灯带后是否需要计算线长限制
                    var 预计组 = new List<(string 序号, double 灯带长度, double 线材长度, string 标签码1, string 标签码2, string 标签码3, string 标签码4, string 客户型号, string 标签显示长度)>(当前组);
                    预计组.Add(灯带);
                    bool 需要考虑线长 = 需要计算线长(预计组);  // 如果加入这条后超过4条或有线长>5米，则需要考虑线长

                    bool 可以添加;
                    if (需要考虑线长)
                    {
                        // 需要计算线长时，使用权重计算
                        double 累计灯带权重系数 = (current灯带米数 + 灯带.灯带长度) / 产品信息.灯带最大长度;
                        double 累计线材权重系数 = (current线材长度 + 灯带.线材长度) / 产品信息.线材最大长度;

                        double 调整后最大线材长度 = 产品信息.线材最大长度 * (1 - 累计灯带权重系数 * 权重值);
                        double 调整后最大灯带米数 = 产品信息.灯带最大长度 * (1 - 累计线材权重系数 * 权重值);

                        可以添加 = (current灯带米数 + 灯带.灯带长度) <= 调整后最大灯带米数 &&
                                  (current线材长度 + 灯带.线材长度) <= 调整后最大线材长度;
                    }
                    else
                    {
                        // 不需要计算线长时，只检查灯带长度
                        可以添加 = (current灯带米数 + 灯带.灯带长度) <= 产品信息.灯带最大长度;
                    }

                    if (可以添加)
                    {
                        当前组.Add(灯带);
                        current灯带米数 += 灯带.灯带长度;
                        current线材长度 += 灯带.线材长度;
                    }
                    else
                    {
                        // 当前组已满，保存并创建新组
                        if (当前组.Any())
                        {
                            使用的包装列表.Add((六百包装.包装名称, 六百包装.物料码, 当前组));
                        }
                        当前组 = new List<(string 序号, double 灯带长度, double 线材长度, string 标签码1, string 标签码2, string 标签码3, string 标签码4, string 客户型号, string 标签显示长度)> { 灯带 };
                        current灯带米数 = 灯带.灯带长度;
                        current线材长度 = 灯带.线材长度;
                    }
                }

                // 处理最后一组
                if (当前组.Any())
                {
                    使用的包装列表.Add((六百包装.包装名称, 六百包装.物料码, 当前组));
                }
            }

            return (使用的包装列表, 单条包装列表记录);
        }

        private void 显示Luminii不同标签码2方案结果(List<包装数据> 匹配包装列表, List<(string 序号, double 灯带长度, double 线材长度, string 标签码1, string 标签码2, string 标签码3, string 标签码4, string 客户型号, string 标签显示长度)> 长度列表, string 型号)
        {
            var 方案结果 = 执行Luminii不同标签码2方案(匹配包装列表, 长度列表, 型号);

            int 方案包装数 = 方案结果.多条包装列表.Count + 方案结果.单条包装列表.Count;

            StringBuilder 结果 = new StringBuilder();
            结果.AppendLine("Luminii-不同标签码2不能混装方案结果：");
            结果.AppendLine($"需要包装数：{方案包装数}");
            结果.AppendLine($"  多条包装：{方案结果.多条包装列表.Count} 个");
            结果.AppendLine($"  单条包装：{方案结果.单条包装列表.Count} 个");
            结果.AppendLine();

            if (方案结果.多条包装列表.Any())
            {
                结果.AppendLine("多条包装详情：");
                foreach (var 包装 in 方案结果.多条包装列表)
                {
                    结果.AppendLine($"包装名称: {包装.包装名称}");
                    结果.AppendLine($"物料码: {包装.物料码}");
                    结果.AppendLine("包含灯带:");
                    double 包装内总灯带长度 = 0;
                    double 包装内总线材长度 = 0;

                    foreach (var 灯带 in 包装.包含灯带)
                    {
                        结果.AppendLine($"  序号: {灯带.序号}");
                        结果.AppendLine($"    灯带长度: {灯带.灯带长度:F2}m");
                        结果.AppendLine($"    线材长度: {灯带.线材长度:F2}m");
                        结果.AppendLine($"    标签码1: {灯带.标签码1}");
                        结果.AppendLine($"    标签码2: {灯带.标签码2}");
                        结果.AppendLine($"    标签码3: {灯带.标签码3}");
                        结果.AppendLine($"    标签码4: {灯带.标签码4}");
                        包装内总灯带长度 += 灯带.灯带长度;
                        包装内总线材长度 += 灯带.线材长度;
                    }
                    结果.AppendLine($"包装内总灯带长度: {包装内总灯带长度:F2}m");
                    结果.AppendLine($"包装内总线材长度: {包装内总线材长度:F2}m");
                    结果.AppendLine("----------------------------------------");
                }
            }

            if (方案结果.单条包装列表.Any())
            {
                结果.AppendLine("\n单条包装详情：");
                foreach (var (包装名称, 物料码, 序号, 灯带长度, 线材长度, 标签码1, 标签码2, 标签码3, 标签码4, 客户型号, 标签显示长度) in 方案结果.单条包装列表)
                {
                    结果.AppendLine($"包装名称: {包装名称}");
                    结果.AppendLine($"物料码: {物料码}");
                    结果.AppendLine($"序号: {序号}");
                    结果.AppendLine($"  灯带长度: {灯带长度:F2}m");
                    结果.AppendLine($"  线材长度: {线材长度:F2}m");
                    结果.AppendLine($"  标签码1: {标签码1}");
                    结果.AppendLine($"  标签码2: {标签码2}");
                    结果.AppendLine($"  标签码3: {标签码3}");
                    结果.AppendLine($"  标签码4: {标签码4}");
                    结果.AppendLine("----------------------------------------");
                }
            }

            保存日志(结果.ToString(), "Luminii-不同标签码2不能混装方案结果", Path.GetFileName(当前处理文件名));

            string 文件夹路径 = Path.Combine(Application.StartupPath, "输出结果", 变量.订单编号);
            if (!Directory.Exists(文件夹路径))
            {
                Directory.CreateDirectory(文件夹路径);
            }

            string 文件路径 = Path.Combine(文件夹路径, "Luminii_不同标签码2_" + Path.GetFileName(当前处理文件名));
            using (ExcelPackage package = new ExcelPackage(new FileInfo(文件路径)))
            {
                int 包装序号 = 1;

                foreach (var 包装 in 方案结果.多条包装列表)
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets.Add($"第{包装序号}盒");

                    worksheet.Cells[1, 1].Value = "序号";
                    worksheet.Cells[1, 2].Value = "条数";
                    worksheet.Cells[1, 3].Value = "米数";
                    worksheet.Cells[1, 4].Value = "标签码1";
                    worksheet.Cells[1, 5].Value = "标签码2";
                    worksheet.Cells[1, 6].Value = "标签码3";
                    worksheet.Cells[1, 7].Value = "标签码4";
                    worksheet.Cells[1, 8].Value = "线长";
                    worksheet.Cells[1, 9].Value = "包装编码";
                    worksheet.Cells[1, 13].Value = "客户型号";
                    worksheet.Cells[1, 14].Value = "标签显示长度";

                    int row = 2;
                    foreach (var 灯带 in 包装.包含灯带)
                    {
                        worksheet.Cells[row, 1].Value = 灯带.序号;
                        worksheet.Cells[row, 2].Value = 1;
                        worksheet.Cells[row, 3].Value = Math.Round(灯带.灯带长度, 3);
                        worksheet.Cells[row, 4].Value = 灯带.标签码1;
                        worksheet.Cells[row, 5].Value = 灯带.标签码2;
                        worksheet.Cells[row, 6].Value = 灯带.标签码3;
                        worksheet.Cells[row, 7].Value = 灯带.标签码4;
                        worksheet.Cells[row, 8].Value = Math.Round(灯带.线材长度, 3);
                        worksheet.Cells[row, 9].Value = 包装.物料码;
                        worksheet.Cells[row, 13].Value = 灯带.客户型号;
                        worksheet.Cells[row, 14].Value = 灯带.标签显示长度;

                        using (var range = worksheet.Cells[row, 1, row, 9])
                        {
                            range.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                            range.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                            range.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                            range.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        }

                        row++;
                    }

                    包装序号++;
                }

                foreach (var 包装 in 方案结果.单条包装列表)
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets.Add($"第{包装序号}盒");

                    worksheet.Cells[1, 1].Value = "序号";
                    worksheet.Cells[1, 2].Value = "条数";
                    worksheet.Cells[1, 3].Value = "米数";
                    worksheet.Cells[1, 4].Value = "标签码1";
                    worksheet.Cells[1, 5].Value = "标签码2";
                    worksheet.Cells[1, 6].Value = "标签码3";
                    worksheet.Cells[1, 7].Value = "标签码4";
                    worksheet.Cells[1, 8].Value = "线长";
                    worksheet.Cells[1, 9].Value = "包装编码";
                    worksheet.Cells[1, 13].Value = "客户型号";
                    worksheet.Cells[1, 14].Value = "标签显示长度";

                    worksheet.Cells[2, 1].Value = 包装.序号;
                    worksheet.Cells[2, 2].Value = 1;
                    worksheet.Cells[2, 3].Value = Math.Round(包装.灯带长度, 3);
                    worksheet.Cells[2, 4].Value = 包装.标签码1;
                    worksheet.Cells[2, 5].Value = 包装.标签码2;
                    worksheet.Cells[2, 6].Value = 包装.标签码3;
                    worksheet.Cells[2, 7].Value = 包装.标签码4;
                    worksheet.Cells[2, 8].Value = Math.Round(包装.线材长度, 3);
                    worksheet.Cells[2, 9].Value = 包装.物料码;
                    worksheet.Cells[2, 13].Value = 包装.客户型号;
                    worksheet.Cells[2, 14].Value = 包装.标签显示长度;

                    using (var range = worksheet.Cells[2, 1, 2, 9])
                    {
                        range.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                        range.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                        range.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                        range.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    }

                    包装序号++;
                }

                package.Save();
            }
            包装方案 选定方案 = new 包装方案
            {
                多条包装列表 = 方案结果.多条包装列表,
                单条包装列表 = 方案结果.单条包装列表
            };

            // 生成包装材料需求流转单
            生成包装材料需求流转单_4(文件夹路径, "Luminii_不同标签码2_" + Path.GetFileName(当前处理文件名), 选定方案);

            uiTextBox_状态.AppendText($"Luminii-不同标签码2不能混装方案已生成，共需要{方案包装数}个包装。" + Environment.NewLine);
            uiTextBox_状态.AppendText("------------------------" + Environment.NewLine);
        }

        private (List<(string 包装名称, string 物料码, List<(string 序号, double 灯带长度, double 线材长度, string 标签码1, string 标签码2, string 标签码3, string 标签码4, string 客户型号, string 标签显示长度)> 包含灯带)> 多条包装列表,
         List<(string 包装名称, string 物料码, string 序号, double 灯带长度, double 线材长度, string 标签码1, string 标签码2, string 标签码3, string 标签码4, string 客户型号, string 标签显示长度)> 单条包装列表)
执行不同序号不可混装方案(List<包装数据> 匹配包装列表, List<(string 序号, double 灯带长度, double 线材长度, string 标签码1, string 标签码2, string 标签码3, string 标签码4, string 客户型号, string 标签显示长度)> 长度列表, string 型号)
        {
            List<(string 包装名称, string 物料码, List<(string 序号, double 灯带长度, double 线材长度, string 标签码1, string 标签码2, string 标签码3, string 标签码4, string 客户型号, string 标签显示长度)> 包含灯带)> 使用的包装列表
                = new List<(string, string, List<(string, double, double, string, string, string, string, string, string)>)>();

            List<(string 包装名称, string 物料码, string 序号, double 灯带长度, double 线材长度, string 标签码1, string 标签码2, string 标签码3, string 标签码4, string 客户型号, string 标签显示长度)> 单条包装列表记录
                = new List<(string, string, string, double, double, string, string, string, string, string, string)>();

            double 权重值 = uiTextBox_权重值.Text.ToDouble();

            // 排除Luminii包装
            var 非Luminii包装列表 = 匹配包装列表.Where(p => !p.包装名称.Contains("Luminii")).ToList();

            // 获取470和600的多条包装
            var 四七零包装 = 非Luminii包装列表.FirstOrDefault(p => p.包装名称?.Contains("470") == true &&
                (p.包装类型?.Contains("多条") == true || p.包装类型?.Contains("2条") == true || p.包装类型?.Contains("3条") == true));

            var 六百包装 = 非Luminii包装列表.FirstOrDefault(p => p.包装名称?.Contains("600") == true &&
                (p.包装类型?.Contains("多条") == true || p.包装类型?.Contains("2条") == true || p.包装类型?.Contains("3条") == true));

            // 按序号前缀分组 - 直接在这里修改，不使用新方法
            var 序号分组 = 长度列表.GroupBy(x =>
            {
                if (x.序号 != null && x.序号.Contains("-"))
                    return x.序号.Split('-')[0];
                return x.序号;
            }).ToList();

            foreach (var 组 in 序号分组)
            {
                var 待处理灯带 = 组.ToList();

                while (待处理灯带.Any())
                {
                    var 当前组 = new List<(string 序号, double 灯带长度, double 线材长度, string 标签码1, string 标签码2, string 标签码3, string 标签码4, string 客户型号, string 标签显示长度)>();
                    double current灯带米数 = 0;
                    double current线材长度 = 0;

                    // 提取序号前缀
                    string 当前序号前缀;
                    if (待处理灯带[0].序号 != null && 待处理灯带[0].序号.Contains("-"))
                        当前序号前缀 = 待处理灯带[0].序号.Split('-')[0];
                    else
                        当前序号前缀 = 待处理灯带[0].序号;

                    var 多条产品信息 = 四七零包装?.支持产品型号列表?.FirstOrDefault(p => p.产品型号 == 型号) ??
                                    六百包装?.支持产品型号列表?.FirstOrDefault(p => p.产品型号 == 型号);

                    if (多条产品信息 == null) continue;

                    // 处理第一条灯带
                    var 第一个灯带 = 待处理灯带[0];
                    当前组.Add(第一个灯带);
                    current灯带米数 = 第一个灯带.灯带长度;
                    current线材长度 = 第一个灯带.线材长度;
                    待处理灯带.RemoveAt(0);

                    // 尝试继续添加相同序号前缀的后续灯带
                    while (待处理灯带.Any())
                    {
                        var 下一个灯带 = 待处理灯带[0];

                        // 提取下一个灯带的序号前缀
                        string 下一个序号前缀;
                        if (下一个灯带.序号 != null && 下一个灯带.序号.Contains("-"))
                            下一个序号前缀 = 下一个灯带.序号.Split('-')[0];
                        else
                            下一个序号前缀 = 下一个灯带.序号;

                        // 检查序号前缀是否相同
                        if (下一个序号前缀 != 当前序号前缀)
                        {
                            break; // 如果不同，停止当前分组
                        }

                        double 预计累计灯带米数 = current灯带米数 + 下一个灯带.灯带长度;
                        double 预计累计线材长度 = current线材长度 + 下一个灯带.线材长度;

                        var 当前灯带组 = 当前组.Concat(new[] { 下一个灯带 }).ToList();

                        // 检查是否需要计算线长
                        bool 忽略线长限制 = !需要计算线长(当前灯带组);

                        // 检查470包装是否可用
                        bool 可以用470 = false;
                        if (四七零包装 != null)
                        {
                            var 产品信息470 = 四七零包装.支持产品型号列表?.FirstOrDefault(p => p.产品型号 == 型号);
                            可以用470 = 产品信息470 != null &&
                                      预计累计灯带米数 <= 产品信息470.灯带最大长度 &&
                                      预计累计灯带米数 >= 产品信息470.最低阈值 &&
                                      (忽略线长限制 || 预计累计线材长度 <= 产品信息470.线材最大长度);
                        }

                        // 检查600包装是否可用
                        bool 可以用600 = false;
                        if (六百包装 != null)
                        {
                            var 产品信息600 = 六百包装.支持产品型号列表?.FirstOrDefault(p => p.产品型号 == 型号);
                            可以用600 = 产品信息600 != null &&
                                      预计累计灯带米数 <= 产品信息600.灯带最大长度 &&
                                      预计累计灯带米数 >= 产品信息600.最低阈值 &&
                                      (忽略线长限制 || 预计累计线材长度 <= 产品信息600.线材最大长度);
                        }

                        if (可以用470 || 可以用600)
                        {
                            当前组.Add(下一个灯带);
                            current灯带米数 = 预计累计灯带米数;
                            current线材长度 = 预计累计线材长度;
                            待处理灯带.RemoveAt(0);
                        }
                        else
                        {
                            break;
                        }
                    }

                    // 处理当前组
                    if (当前组.Count == 1)
                    {
                        // 尝试使用单条包装
                        var 单条包装列表 = 非Luminii包装列表
                            .Where(p => p.包装类型?.Contains("单条") == true)
                            .OrderBy(p =>
                            {
                                if (p.包装名称?.Contains("330") == true) return 1;
                                if (p.包装名称?.Contains("470") == true) return 2;
                                if (p.包装名称?.Contains("600") == true) return 3;
                                return 4;
                            })
                            .ToList();

                        bool 找到单条包装 = false;
                        foreach (var 包装 in 单条包装列表)
                        {
                            var 灯带 = 当前组[0];
                            bool 忽略线长限制 = !需要计算线长(当前组);

                            var 单条产品信息 = 包装.支持产品型号列表?.FirstOrDefault(p => p.产品型号 == 型号);
                            if (单条产品信息 == null) continue;

                            if (灯带.灯带长度 <= 单条产品信息.灯带最大长度 &&
                                灯带.灯带长度 >= 单条产品信息.最低阈值 &&
                                (忽略线长限制 || 灯带.线材长度 <= 单条产品信息.线材最大长度))
                            {
                                单条包装列表记录.Add((包装.包装名称, 包装.物料码, 灯带.序号, 灯带.灯带长度, 灯带.线材长度, 灯带.标签码1, 灯带.标签码2, 灯带.标签码3, 灯带.标签码4, 灯带.客户型号, 灯带.标签显示长度));
                                找到单条包装 = true;
                                break;
                            }
                        }

                        if (!找到单条包装)
                        {
                            // 如果找不到合适的单条包装，尝试使用多条包装
                            var 灯带 = 当前组[0];
                            bool 忽略线长限制 = !需要计算线长(当前组);

                            var 产品信息470 = 四七零包装?.支持产品型号列表?.FirstOrDefault(p => p.产品型号 == 型号);
                            bool 可以用470单条 = 四七零包装 != null && 产品信息470 != null &&
                                               灯带.灯带长度 <= 产品信息470.灯带最大长度 &&
                                               灯带.灯带长度 >= 产品信息470.最低阈值 &&
                                               (忽略线长限制 || 灯带.线材长度 <= 产品信息470.线材最大长度);

                            if (可以用470单条)
                            {
                                使用的包装列表.Add((四七零包装.包装名称, 四七零包装.物料码, 当前组));
                            }
                            else if (六百包装 != null)
                            {
                                var 产品信息600 = 六百包装.支持产品型号列表?.FirstOrDefault(p => p.产品型号 == 型号);
                                bool 可以用600单条 = 产品信息600 != null &&
                                                   灯带.灯带长度 <= 产品信息600.灯带最大长度 &&
                                                   灯带.灯带长度 >= 产品信息600.最低阈值 &&
                                                   (忽略线长限制 || 灯带.线材长度 <= 产品信息600.线材最大长度);

                                if (可以用600单条)
                                {
                                    使用的包装列表.Add((六百包装.包装名称, 六百包装.物料码, 当前组));
                                }
                                else
                                {
                                    throw new Exception($"无法找到合适的包装方案处理灯带 {灯带.序号}");
                                }
                            }
                        }
                    }
                    else
                    {
                        // 多条包装处理
                        bool 忽略线长限制 = !需要计算线长(当前组);

                        var 产品信息470 = 四七零包装?.支持产品型号列表?.FirstOrDefault(p => p.产品型号 == 型号);
                        bool 可以用470多条 = 四七零包装 != null && 产品信息470 != null &&
                                           current灯带米数 <= 产品信息470.灯带最大长度 &&
                                           current灯带米数 >= 产品信息470.最低阈值 &&
                                           (忽略线长限制 || current线材长度 <= 产品信息470.线材最大长度);

                        if (可以用470多条)
                        {
                            使用的包装列表.Add((四七零包装.包装名称, 四七零包装.物料码, 当前组));
                        }
                        else if (六百包装 != null)
                        {
                            var 产品信息600 = 六百包装.支持产品型号列表?.FirstOrDefault(p => p.产品型号 == 型号);
                            bool 可以用600多条 = 产品信息600 != null &&
                                               current灯带米数 <= 产品信息600.灯带最大长度 &&
                                               current灯带米数 >= 产品信息600.最低阈值 &&
                                               (忽略线长限制 || current线材长度 <= 产品信息600.线材最大长度);

                            if (可以用600多条)
                            {
                                使用的包装列表.Add((六百包装.包装名称, 六百包装.物料码, 当前组));
                            }
                            else
                            {
                                throw new Exception("无法找到合适的多条包装");
                            }
                        }
                    }
                }
            }

            return (使用的包装列表, 单条包装列表记录);
        }

        private void 显示不同序号不可混装方案结果(List<包装数据> 匹配包装列表, List<(string 序号, double 灯带长度, double 线材长度, string 标签码1, string 标签码2, string 标签码3, string 标签码4, string 客户型号, string 标签显示长度)> 长度列表, string 型号)
        {
            var 方案结果 = 执行不同序号不可混装方案(匹配包装列表, 长度列表, 型号);

            int 方案包装数 = 方案结果.多条包装列表.Count + 方案结果.单条包装列表.Count;

            StringBuilder 结果 = new StringBuilder();
            结果.AppendLine("不同序号不可混装方案结果：");
            结果.AppendLine($"需要包装数：{方案包装数}");
            结果.AppendLine($"  多条包装：{方案结果.多条包装列表.Count} 个");
            结果.AppendLine($"  单条包装：{方案结果.单条包装列表.Count} 个");
            结果.AppendLine();

            if (方案结果.多条包装列表.Any())
            {
                结果.AppendLine("多条包装详情：");
                foreach (var 包装 in 方案结果.多条包装列表)
                {
                    结果.AppendLine($"包装名称: {包装.包装名称}");
                    结果.AppendLine($"物料码: {包装.物料码}");
                    结果.AppendLine("包含灯带:");
                    double 包装内总灯带长度 = 0;
                    double 包装内总线材长度 = 0;

                    foreach (var 灯带 in 包装.包含灯带)
                    {
                        结果.AppendLine($"  序号: {灯带.序号}");
                        结果.AppendLine($"    灯带长度: {灯带.灯带长度:F2}m");
                        结果.AppendLine($"    线材长度: {灯带.线材长度:F2}m");
                        结果.AppendLine($"    标签码1: {灯带.标签码1}");
                        结果.AppendLine($"    标签码2: {灯带.标签码2}");
                        结果.AppendLine($"    标签码3: {灯带.标签码3}");
                        结果.AppendLine($"    标签码4: {灯带.标签码4}");
                        包装内总灯带长度 += 灯带.灯带长度;
                        包装内总线材长度 += 灯带.线材长度;
                    }
                    结果.AppendLine($"包装内总灯带长度: {包装内总灯带长度:F2}m");
                    结果.AppendLine($"包装内总线材长度: {包装内总线材长度:F2}m");
                    结果.AppendLine("----------------------------------------");
                }
            }

            if (方案结果.单条包装列表.Any())
            {
                结果.AppendLine("\n单条包装详情：");
                foreach (var (包装名称, 物料码, 序号, 灯带长度, 线材长度, 标签码1, 标签码2, 标签码3, 标签码4, 客户型号, 标签显示长度) in 方案结果.单条包装列表)
                {
                    结果.AppendLine($"包装名称: {包装名称}");
                    结果.AppendLine($"物料码: {物料码}");
                    结果.AppendLine($"序号: {序号}");
                    结果.AppendLine($"  灯带长度: {灯带长度:F2}m");
                    结果.AppendLine($"  线材长度: {线材长度:F2}m");
                    结果.AppendLine($"  标签码1: {标签码1}");
                    结果.AppendLine($"  标签码2: {标签码2}");
                    结果.AppendLine($"  标签码3: {标签码3}");
                    结果.AppendLine($"  标签码4: {标签码4}");
                    结果.AppendLine("----------------------------------------");
                }
            }

            保存日志(结果.ToString(), "不同序号不可混装方案结果", Path.GetFileName(当前处理文件名));

            string outputDirectory = Path.Combine(Application.StartupPath, "输出结果", 变量.订单编号);
            if (!Directory.Exists(outputDirectory))
            {
                Directory.CreateDirectory(outputDirectory);
            }

            string outputFilePath = Path.Combine(outputDirectory, "不同序号不可混装_" + Path.GetFileName(当前处理文件名));
            using (ExcelPackage package = new ExcelPackage(new FileInfo(outputFilePath)))
            {
                int 包装序号 = 1;

                foreach (var 包装 in 方案结果.多条包装列表)
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets.Add($"第{包装序号}盒");

                    worksheet.Cells[1, 1].Value = "序号";
                    worksheet.Cells[1, 2].Value = "条数";
                    worksheet.Cells[1, 3].Value = "米数";
                    worksheet.Cells[1, 4].Value = "标签码1";
                    worksheet.Cells[1, 5].Value = "标签码2";
                    worksheet.Cells[1, 6].Value = "标签码3";
                    worksheet.Cells[1, 7].Value = "标签码4";
                    worksheet.Cells[1, 8].Value = "线长";
                    worksheet.Cells[1, 9].Value = "包装编码";
                    worksheet.Cells[1, 13].Value = "客户型号";
                    worksheet.Cells[1, 14].Value = "标签显示长度";

                    int row = 2;
                    foreach (var 灯带 in 包装.包含灯带)
                    {
                        worksheet.Cells[row, 1].Value = 灯带.序号;
                        worksheet.Cells[row, 2].Value = 1;
                        worksheet.Cells[row, 3].Value = Math.Round(灯带.灯带长度, 3);
                        worksheet.Cells[row, 4].Value = 灯带.标签码1;
                        worksheet.Cells[row, 5].Value = 灯带.标签码2;
                        worksheet.Cells[row, 6].Value = 灯带.标签码3;
                        worksheet.Cells[row, 7].Value = 灯带.标签码4;
                        worksheet.Cells[row, 8].Value = Math.Round(灯带.线材长度, 3);
                        worksheet.Cells[row, 9].Value = 包装.物料码;
                        worksheet.Cells[row, 13].Value = 灯带.客户型号;
                        worksheet.Cells[row, 14].Value = 灯带.标签显示长度;

                        using (var range = worksheet.Cells[row, 1, row, 9])
                        {
                            range.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                            range.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                            range.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                            range.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        }

                        row++;
                    }

                    包装序号++;
                }

                foreach (var 包装 in 方案结果.单条包装列表)
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets.Add($"第{包装序号}盒");

                    worksheet.Cells[1, 1].Value = "序号";
                    worksheet.Cells[1, 2].Value = "条数";
                    worksheet.Cells[1, 3].Value = "米数";
                    worksheet.Cells[1, 4].Value = "标签码1";
                    worksheet.Cells[1, 5].Value = "标签码2";
                    worksheet.Cells[1, 6].Value = "标签码3";
                    worksheet.Cells[1, 7].Value = "标签码4";
                    worksheet.Cells[1, 8].Value = "线长";
                    worksheet.Cells[1, 9].Value = "包装编码";
                    worksheet.Cells[1, 13].Value = "客户型号";
                    worksheet.Cells[1, 14].Value = "标签显示长度";

                    worksheet.Cells[2, 1].Value = 包装.序号;
                    worksheet.Cells[2, 2].Value = 1;
                    worksheet.Cells[2, 3].Value = Math.Round(包装.灯带长度, 3);
                    worksheet.Cells[2, 4].Value = 包装.标签码1;
                    worksheet.Cells[2, 5].Value = 包装.标签码2;
                    worksheet.Cells[2, 6].Value = 包装.标签码3;
                    worksheet.Cells[2, 7].Value = 包装.标签码4;
                    worksheet.Cells[2, 8].Value = Math.Round(包装.线材长度, 3);
                    worksheet.Cells[2, 9].Value = 包装.物料码;
                    worksheet.Cells[2, 13].Value = 包装.客户型号;
                    worksheet.Cells[2, 14].Value = 包装.标签显示长度;

                    using (var range = worksheet.Cells[2, 1, 2, 9])
                    {
                        range.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                        range.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                        range.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                        range.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    }

                    包装序号++;
                }

                package.Save();
            }

            // 创建包装方案对象，用于生成流转单
            包装方案 选定方案 = new 包装方案
            {
                多条包装列表 = 方案结果.多条包装列表,
                单条包装列表 = 方案结果.单条包装列表,
                方案名称 = "不同序号不可混装方案"
            };

            // 生成包装材料需求流转单
            生成包装材料需求流转单_4(outputDirectory, "不同序号不可混装_" + Path.GetFileName(当前处理文件名), 选定方案);

            uiTextBox_状态.AppendText($"不同序号不可混装方案已生成，共需要{方案包装数}个包装。" + Environment.NewLine);
            uiTextBox_状态.AppendText("------------------------" + Environment.NewLine);
        }

        private (List<(string 包装名称, string 物料码, List<(string 序号, double 灯带长度, double 线材长度, string 标签码1, string 标签码2, string 标签码3, string 标签码4, string 客户型号, string 标签显示长度)> 包含灯带)> 多条包装列表,
         List<(string 包装名称, string 物料码, string 序号, double 灯带长度, double 线材长度, string 标签码1, string 标签码2, string 标签码3, string 标签码4, string 客户型号, string 标签显示长度)> 单条包装列表)
执行不同标签码1不能混装方案(List<包装数据> 匹配包装列表, List<(string 序号, double 灯带长度, double 线材长度, string 标签码1, string 标签码2, string 标签码3, string 标签码4, string 客户型号, string 标签显示长度)> 长度列表, string 型号)
        {
            var 非Luminii包装列表 = 匹配包装列表.Where(p => !p.包装名称.Contains("Luminii")).ToList();

            var 多条包装列表 = new List<(string 包装名称, string 物料码, List<(string 序号, double 灯带长度, double 线材长度, string 标签码1, string 标签码2, string 标签码3, string 标签码4, string 客户型号, string 标签显示长度)> 包含灯带)>();
            var 单条包装列表 = new List<(string 包装名称, string 物料码, string 序号, double 灯带长度, double 线材长度, string 标签码1, string 标签码2, string 标签码3, string 标签码4, string 客户型号, string 标签显示长度)>();

            var 四七零包装 = 非Luminii包装列表.FirstOrDefault(p => p.包装名称?.Contains("470") == true &&
                (p.包装类型?.Contains("多条") == true || p.包装类型?.Contains("2条") == true || p.包装类型?.Contains("3条") == true));
            var 六百包装 = 非Luminii包装列表.FirstOrDefault(p => p.包装名称?.Contains("600") == true &&
                (p.包装类型?.Contains("多条") == true || p.包装类型?.Contains("2条") == true || p.包装类型?.Contains("3条") == true));

            var 单条包装列表数据 = 非Luminii包装列表
                .Where(p => p.包装类型?.Contains("单条") == true)
                .OrderBy(p =>
                {
                    if (p.包装名称?.Contains("330") == true) return 1;
                    if (p.包装名称?.Contains("470") == true) return 2;
                    if (p.包装名称?.Contains("600") == true) return 3;
                    return 4;
                })
                .ToList();

            if (四七零包装 == null && 六百包装 == null)
            {
                MessageBox.Show("未找到合适的多条包装方案", "提示");
                return (多条包装列表, 单条包装列表);
            }

            double 权重值 = uiTextBox_权重值.Text.ToDouble();

            var 标签码1分组 = 长度列表.GroupBy(x => x.标签码1).ToList();

            foreach (var 组 in 标签码1分组)
            {
                var 待处理灯带 = 组.ToList();
                string 当前标签码1 = 待处理灯带[0].标签码1;

                while (待处理灯带.Any())
                {
                    var 当前组 = new List<(string 序号, double 灯带长度, double 线材长度, string 标签码1, string 标签码2, string 标签码3, string 标签码4, string 客户型号, string 标签显示长度)>();
                    double current灯带米数 = 0;
                    double current线材长度 = 0;

                    var 第一个灯带 = 待处理灯带[0];
                    当前组.Add(第一个灯带);
                    current灯带米数 = 第一个灯带.灯带长度;
                    current线材长度 = 第一个灯带.线材长度;
                    待处理灯带.RemoveAt(0);

                    int 最大迭代次数 = 1000;
                    int 当前迭代 = 0;

                    while (待处理灯带.Any() && 当前迭代 < 最大迭代次数)
                    {
                        当前迭代++;
                        var 下一个灯带 = 待处理灯带[0];
                        if (下一个灯带.标签码1 != 当前标签码1)
                        {
                            break;
                        }
                        double 预计累计灯带米数 = current灯带米数 + 下一个灯带.灯带长度;
                        double 预计累计线材长度 = current线材长度 + 下一个灯带.线材长度;
                        var 当前灯带组 = 当前组.Concat(new[] { 下一个灯带 }).ToList();

                        bool 可以用470 = 四七零包装 != null && 检查包装是否可用(四七零包装, 型号, 预计累计灯带米数, 预计累计线材长度, 权重值, 当前灯带组);
                        bool 可以用600 = 六百包装 != null && 检查包装是否可用(六百包装, 型号, 预计累计灯带米数, 预计累计线材长度, 权重值, 当前灯带组);

                        if (可以用470 || 可以用600)
                        {
                            当前组.Add(下一个灯带);
                            current灯带米数 = 预计累计灯带米数;
                            current线材长度 = 预计累计线材长度;
                            待处理灯带.RemoveAt(0);
                        }
                        else
                        {
                            break;
                        }
                    }

                    if (当前组.Count == 1)
                    {
                        bool 找到单条包装 = false;
                        foreach (var 包装 in 单条包装列表数据)
                        {
                            var 单条产品信息 = 包装.支持产品型号列表?.FirstOrDefault(p => p.产品型号 == 型号);
                            if (单条产品信息 == null) continue;
                            var 灯带 = 当前组[0];
                            var 单条灯带列表 = new List<(string 序号, double 灯带长度, double 线材长度, string 标签码1, string 标签码2, string 标签码3, string 标签码4, string 客户型号, string 标签显示长度)> { 灯带 };
                            bool 忽略线长限制 = !需要计算线长(单条灯带列表);
                            if (灯带.灯带长度 <= 单条产品信息.灯带最大长度 &&
                                灯带.灯带长度 >= 单条产品信息.最低阈值 &&
                                (忽略线长限制 || 灯带.线材长度 <= 单条产品信息.线材最大长度))
                            {
                                单条包装列表.Add((包装.包装名称, 包装.物料码, 灯带.序号, 灯带.灯带长度, 灯带.线材长度, 灯带.标签码1, 灯带.标签码2, 灯带.标签码3, 灯带.标签码4, 灯带.客户型号, 灯带.标签显示长度));
                                找到单条包装 = true;
                                break;
                            }
                        }
                        if (!找到单条包装)
                        {
                            bool 可以用470单条 = false;
                            if (四七零包装 != null)
                            {
                                var 单条灯带列表 = new List<(string 序号, double 灯带长度, double 线材长度, string 标签码1, string 标签码2, string 标签码3, string 标签码4, string 客户型号, string 标签显示长度)> { 当前组[0] };
                                可以用470单条 = 检查包装是否可用(四七零包装, 型号, 当前组[0].灯带长度, 当前组[0].线材长度, 权重值, 单条灯带列表);
                                if (可以用470单条)
                                {
                                    多条包装列表.Add((四七零包装.包装名称, 四七零包装.物料码, 当前组));
                                }
                            }
                            if (!可以用470单条 && 六百包装 != null)
                            {
                                var 单条灯带列表 = new List<(string 序号, double 灯带长度, double 线材长度, string 标签码1, string 标签码2, string 标签码3, string 标签码4, string 客户型号, string 标签显示长度)> { 当前组[0] };
                                bool 可以用600单条 = 检查包装是否可用(六百包装, 型号, 当前组[0].灯带长度, 当前组[0].线材长度, 权重值, 单条灯带列表);
                                if (可以用600单条)
                                {
                                    多条包装列表.Add((六百包装.包装名称, 六百包装.物料码, 当前组));
                                }
                                else
                                {
                                    throw new Exception($"无法找到合适的包装方案处理灯带 {当前组[0].序号}");
                                }
                            }
                        }
                    }
                    else
                    {
                        bool 可以用470多条 = false;
                        if (四七零包装 != null)
                        {
                            可以用470多条 = 检查包装是否可用(四七零包装, 型号, current灯带米数, current线材长度, 权重值, 当前组);
                        }
                        if (可以用470多条)
                        {
                            多条包装列表.Add((四七零包装.包装名称, 四七零包装.物料码, 当前组));
                        }
                        else if (六百包装 != null)
                        {
                            bool 可以用600多条 = 检查包装是否可用(六百包装, 型号, current灯带米数, current线材长度, 权重值, 当前组);
                            if (可以用600多条)
                            {
                                多条包装列表.Add((六百包装.包装名称, 六百包装.物料码, 当前组));
                            }
                            else
                            {
                                throw new Exception("无法找到合适的多条包装");
                            }
                        }
                        else
                        {
                            throw new Exception("无法找到合适的多条包装");
                        }
                    }
                }
            }
            return (多条包装列表, 单条包装列表);
        }

        private (List<(string 包装名称, string 物料码, List<(string 序号, double 灯带长度, double 线材长度, string 标签码1, string 标签码2, string 标签码3, string 标签码4, string 客户型号, string 标签显示长度)> 包含灯带)> 多条包装列表,
         List<(string 包装名称, string 物料码, string 序号, double 灯带长度, double 线材长度, string 标签码1, string 标签码2, string 标签码3, string 标签码4, string 客户型号, string 标签显示长度)> 单条包装列表)
执行不同标签码1不能混装方案1(List<包装数据> 匹配包装列表, List<(string 序号, double 灯带长度, double 线材长度, string 标签码1, string 标签码2, string 标签码3, string 标签码4, string 客户型号, string 标签显示长度)> 长度列表, string 型号)
        {
            // 排除 Luminii 的包装
            var 非Luminii包装列表 = 匹配包装列表.Where(p => !p.包装名称.Contains("Luminii")).ToList();

            var 多条包装列表 = new List<(string 包装名称, string 物料码, List<(string 序号, double 灯带长度, double 线材长度, string 标签码1, string 标签码2, string 标签码3, string 标签码4, string 客户型号, string 标签显示长度)> 包含灯带)>();
            var 单条包装列表 = new List<(string 包装名称, string 物料码, string 序号, double 灯带长度, double 线材长度, string 标签码1, string 标签码2, string 标签码3, string 标签码4, string 客户型号, string 标签显示长度)>();

            // 获取470和600的多条包装
            var 四七零包装 = 非Luminii包装列表.FirstOrDefault(p => p.包装名称?.Contains("470") == true &&
                (p.包装类型?.Contains("多条") == true || p.包装类型?.Contains("2条") == true || p.包装类型?.Contains("3条") == true));

            var 六百包装 = 非Luminii包装列表.FirstOrDefault(p => p.包装名称?.Contains("600") == true &&
                (p.包装类型?.Contains("多条") == true || p.包装类型?.Contains("2条") == true || p.包装类型?.Contains("3条") == true));

            // 获取单条包装列表，按优先级排序
            var 单条包装列表数据 = 非Luminii包装列表
                .Where(p => p.包装类型?.Contains("单条") == true)
                .OrderBy(p =>
                {
                    if (p.包装名称?.Contains("330") == true) return 1;
                    if (p.包装名称?.Contains("470") == true) return 2;
                    if (p.包装名称?.Contains("600") == true) return 3;
                    return 4;
                })
                .ToList();

            if (四七零包装 == null && 六百包装 == null)
            {
                MessageBox.Show("未找到合适的多条包装方案", "提示");
                return (多条包装列表, 单条包装列表);
            }

            double 权重值 = uiTextBox_权重值.Text.ToDouble();

            // 按标签码1分组
            var 标签码1分组 = 长度列表.GroupBy(x => x.标签码1).ToList();

            foreach (var 组 in 标签码1分组)
            {
                var 待处理灯带 = 组.ToList();
                string 当前标签码1 = 待处理灯带[0].标签码1;

                // 先尝试将整组灯带放在一起（因为它们有相同的标签码1）
                double 总灯带长度 = 待处理灯带.Sum(x => x.灯带长度);
                double 总线材长度 = 待处理灯带.Sum(x => x.线材长度);
                bool 忽略线长限制_整组 = !需要计算线长(待处理灯带);
                bool 整组放入一个包装 = false;

                // 先检查600包装
                if (六百包装 != null)
                {
                    var 产品信息600 = 六百包装.支持产品型号列表?.FirstOrDefault(p => p.产品型号 == 型号);
                    bool 可以用600 = 产品信息600 != null &&
                                    总灯带长度 <= 产品信息600.灯带最大长度 &&
                                    (忽略线长限制_整组 || 总线材长度 <= 产品信息600.线材最大长度);

                    Console.WriteLine($"检查是否可以将整组放入600包装: 总长={总灯带长度}m, 最大长度={产品信息600?.灯带最大长度}m, 可用={可以用600}");

                    if (可以用600)
                    {
                        多条包装列表.Add((六百包装.包装名称, 六百包装.物料码, 待处理灯带));
                        整组放入一个包装 = true;
                    }
                }

                // 如果600不行，再检查470
                if (!整组放入一个包装 && 四七零包装 != null)
                {
                    var 产品信息470 = 四七零包装.支持产品型号列表?.FirstOrDefault(p => p.产品型号 == 型号);
                    bool 可以用470 = 产品信息470 != null &&
                                    总灯带长度 <= 产品信息470.灯带最大长度 &&
                                    (忽略线长限制_整组 || 总线材长度 <= 产品信息470.线材最大长度);

                    Console.WriteLine($"检查是否可以将整组放入470包装: 总长={总灯带长度}m, 最大长度={产品信息470?.灯带最大长度}m, 可用={可以用470}");

                    if (可以用470)
                    {
                        多条包装列表.Add((四七零包装.包装名称, 四七零包装.物料码, 待处理灯带));
                        整组放入一个包装 = true;
                    }
                }

                // 如果整组能放入一个包装，继续处理下一组
                if (整组放入一个包装)
                {
                    continue;
                }

                // 如果整组不能放入一个包装，则使用原有逻辑分批处理
                while (待处理灯带.Count > 0)
                {
                    var 当前组 = new List<(string 序号, double 灯带长度, double 线材长度, string 标签码1, string 标签码2, string 标签码3, string 标签码4, string 客户型号, string 标签显示长度)>();
                    double current灯带米数 = 0;
                    double current线材长度 = 0;

                    // 处理第一条灯带
                    var 第一个灯带 = 待处理灯带[0];
                    当前组.Add(第一个灯带);
                    current灯带米数 = 第一个灯带.灯带长度;
                    current线材长度 = 第一个灯带.线材长度;
                    待处理灯带.RemoveAt(0);

                    // 尝试继续添加相同标签码1的灯带
                    // 添加最大迭代次数防止无限循环
                    int 最大迭代次数 = 1000;
                    int 当前迭代 = 0;

                    while (待处理灯带.Any() && 当前迭代 < 最大迭代次数)
                    {
                        当前迭代++;
                        var 下一个灯带 = 待处理灯带[0];

                        // 检查标签码1是否相同 - 实际上这是多余的检查，因为我们已经按标签码1分组
                        if (下一个灯带.标签码1 != 当前标签码1)
                        {
                            break; // 如果不同，停止当前分组
                        }

                        double 预计累计灯带米数 = current灯带米数 + 下一个灯带.灯带长度;
                        double 预计累计线材长度 = current线材长度 + 下一个灯带.线材长度;

                        var 当前灯带组 = 当前组.Concat(new[] { 下一个灯带 }).ToList();

                        // 检查包装是否可用 - 检查600和470
                        bool 可以用600 = 六百包装 != null && 检查包装是否可用(六百包装, 型号, 预计累计灯带米数, 预计累计线材长度, 权重值, 当前灯带组);
                        bool 可以用470 = 四七零包装 != null && 检查包装是否可用(四七零包装, 型号, 预计累计灯带米数, 预计累计线材长度, 权重值, 当前灯带组);

                        // 添加调试信息
                        string 调试信息 = $"尝试合并灯带 序号:{下一个灯带.序号}, 标签码1:{下一个灯带.标签码1}\n" +
                                       $"累计灯带长度:{预计累计灯带米数}m, 累计线材长度:{预计累计线材长度}m\n" +
                                       $"600可用:{可以用600}, 470可用:{可以用470}";

                        var 产品信息600 = 六百包装?.支持产品型号列表?.FirstOrDefault(p => p.产品型号 == 型号);
                        var 产品信息470 = 四七零包装?.支持产品型号列表?.FirstOrDefault(p => p.产品型号 == 型号);

                        调试信息 += $"\n600最大长度:{产品信息600?.灯带最大长度}m, 最低阈值:{产品信息600?.最低阈值}m";
                        调试信息 += $"\n470最大长度:{产品信息470?.灯带最大长度}m, 最低阈值:{产品信息470?.最低阈值}m";

                        // 使用日志输出而不是MessageBox，避免阻塞
                        Console.WriteLine(调试信息);

                        if (可以用600 || 可以用470)
                        {
                            当前组.Add(下一个灯带);
                            current灯带米数 = 预计累计灯带米数;
                            current线材长度 = 预计累计线材长度;
                            待处理灯带.RemoveAt(0);
                            Console.WriteLine($"添加灯带成功，当前组数量: {当前组.Count}, 剩余待处理: {待处理灯带.Count}");
                        }
                        else
                        {
                            Console.WriteLine($"无法添加灯带 {下一个灯带.序号} 到当前组");
                            break;
                        }
                    }

                    // 处理当前组
                    if (当前组.Count > 0)
                    {
                        bool 忽略线长限制 = !需要计算线长(当前组);
                        bool 所有标签码相同 = 当前组.All(x => x.标签码1 == 当前组[0].标签码1);

                        // 先尝试600包装 - 通常容量更大
                        if (六百包装 != null)
                        {
                            var 产品信息600 = 六百包装.支持产品型号列表?.FirstOrDefault(p => p.产品型号 == 型号);
                            bool 可以用600 = 产品信息600 != null &&
                                            current灯带米数 <= 产品信息600.灯带最大长度 &&
                                            (忽略线长限制 || current线材长度 <= 产品信息600.线材最大长度);

                            if (可以用600)
                            {
                                多条包装列表.Add((六百包装.包装名称, 六百包装.物料码, 当前组));
                                continue;
                            }
                        }

                        // 然后尝试470包装
                        if (四七零包装 != null)
                        {
                            var 产品信息470 = 四七零包装.支持产品型号列表?.FirstOrDefault(p => p.产品型号 == 型号);
                            bool 可以用470 = 产品信息470 != null &&
                                            current灯带米数 <= 产品信息470.灯带最大长度 &&
                                            (忽略线长限制 || current线材长度 <= 产品信息470.线材最大长度);

                            if (可以用470)
                            {
                                多条包装列表.Add((四七零包装.包装名称, 四七零包装.物料码, 当前组));
                                continue;
                            }
                        }

                        // 检查是否可以使用最低阈值限制
                        if (六百包装 != null)
                        {
                            var 产品信息600 = 六百包装.支持产品型号列表?.FirstOrDefault(p => p.产品型号 == 型号);
                            bool 可以用600 = 产品信息600 != null &&
                                            current灯带米数 <= 产品信息600.灯带最大长度 &&
                                            current灯带米数 >= 产品信息600.最低阈值 &&
                                            (忽略线长限制 || current线材长度 <= 产品信息600.线材最大长度);

                            if (可以用600)
                            {
                                多条包装列表.Add((六百包装.包装名称, 六百包装.物料码, 当前组));
                                continue;
                            }
                        }

                        if (四七零包装 != null)
                        {
                            var 产品信息470 = 四七零包装.支持产品型号列表?.FirstOrDefault(p => p.产品型号 == 型号);
                            bool 可以用470 = 产品信息470 != null &&
                                            current灯带米数 <= 产品信息470.灯带最大长度 &&
                                            current灯带米数 >= 产品信息470.最低阈值 &&
                                            (忽略线长限制 || current线材长度 <= 产品信息470.线材最大长度);

                            if (可以用470)
                            {
                                多条包装列表.Add((四七零包装.包装名称, 四七零包装.物料码, 当前组));
                                continue;
                            }
                        }

                        if (当前组.Count == 1 && 单条包装列表数据.Any())
                        {
                            var 灯带 = 当前组[0];
                            bool 找到单条包装 = false;

                            foreach (var 包装 in 单条包装列表数据)
                            {
                                bool 忽略线长_单条 = !需要计算线长(当前组);
                                var 单条产品信息 = 包装.支持产品型号列表?.FirstOrDefault(p => p.产品型号 == 型号);

                                if (单条产品信息 == null) continue;

                                if (灯带.灯带长度 <= 单条产品信息.灯带最大长度 &&
                                    灯带.灯带长度 >= 单条产品信息.最低阈值 &&
                                            (忽略线长_单条 || 灯带.线材长度 <= 单条产品信息.线材最大长度))
                                {
                                    单条包装列表.Add((包装.包装名称, 包装.物料码, 灯带.序号, 灯带.灯带长度, 灯带.线材长度, 灯带.标签码1, 灯带.标签码2, 灯带.标签码3, 灯带.标签码4, 灯带.客户型号, 灯带.标签显示长度));
                                    找到单条包装 = true;
                                    break;
                                }
                            }

                            if (!找到单条包装)
                            {
                                if (四七零包装 != null)
                                {
                                    多条包装列表.Add((四七零包装.包装名称, 四七零包装.物料码, 当前组));
                                }
                                else if (六百包装 != null)
                                {
                                    多条包装列表.Add((六百包装.包装名称, 六百包装.物料码, 当前组));
                                }
                                else
                                {
                                    throw new Exception($"无法找到合适的包装方案处理灯带 {灯带.序号}");
                                }
                            }
                        }
                        else
                        {
                            if (四七零包装 != null)
                            {
                                多条包装列表.Add((四七零包装.包装名称, 四七零包装.物料码, 当前组));
                            }
                            else if (六百包装 != null)
                            {
                                多条包装列表.Add((六百包装.包装名称, 六百包装.物料码, 当前组));
                            }
                            else
                            {
                                throw new Exception("无法找到合适的多条包装");
                            }
                        }
                    }
                }
            }

            return (多条包装列表, 单条包装列表);
        }

        private void 显示不同标签码1不能混装方案结果(List<包装数据> 匹配包装列表, List<(string 序号, double 灯带长度, double 线材长度, string 标签码1, string 标签码2, string 标签码3, string 标签码4, string 客户型号, string 标签显示长度)> 长度列表, string 型号)
        {
            var 方案结果 = 执行不同标签码1不能混装方案(匹配包装列表, 长度列表, 型号);

            int 方案包装数 = 方案结果.多条包装列表.Count + 方案结果.单条包装列表.Count;

            StringBuilder 结果 = new StringBuilder();
            结果.AppendLine("不同标签码1不能混装方案结果：");
            结果.AppendLine($"需要包装数：{方案包装数}");
            结果.AppendLine($"  多条包装：{方案结果.多条包装列表.Count} 个");
            结果.AppendLine($"  单条包装：{方案结果.单条包装列表.Count} 个");
            结果.AppendLine();

            if (方案结果.多条包装列表.Any())
            {
                结果.AppendLine("多条包装详情：");
                foreach (var 包装 in 方案结果.多条包装列表)
                {
                    结果.AppendLine($"包装名称: {包装.包装名称}");
                    结果.AppendLine($"物料码: {包装.物料码}");
                    结果.AppendLine("包含灯带:");
                    double 包装内总灯带长度 = 0;
                    double 包装内总线材长度 = 0;

                    foreach (var 灯带 in 包装.包含灯带)
                    {
                        结果.AppendLine($"  序号: {灯带.序号}");
                        结果.AppendLine($"    灯带长度: {灯带.灯带长度:F2}m");
                        结果.AppendLine($"    线材长度: {灯带.线材长度:F2}m");
                        结果.AppendLine($"    标签码1: {灯带.标签码1}");
                        结果.AppendLine($"    标签码2: {灯带.标签码2}");
                        结果.AppendLine($"    标签码3: {灯带.标签码3}");
                        结果.AppendLine($"    标签码4: {灯带.标签码4}");
                        包装内总灯带长度 += 灯带.灯带长度;
                        包装内总线材长度 += 灯带.线材长度;
                    }
                    结果.AppendLine($"包装内总灯带长度: {包装内总灯带长度:F2}m");
                    结果.AppendLine($"包装内总线材长度: {包装内总线材长度:F2}m");
                    结果.AppendLine("----------------------------------------");
                }
            }

            if (方案结果.单条包装列表.Any())
            {
                结果.AppendLine("\n单条包装详情：");
                foreach (var (包装名称, 物料码, 序号, 灯带长度, 线材长度, 标签码1, 标签码2, 标签码3, 标签码4, 客户型号, 标签显示长度) in 方案结果.单条包装列表)
                {
                    结果.AppendLine($"包装名称: {包装名称}");
                    结果.AppendLine($"物料码: {物料码}");
                    结果.AppendLine($"序号: {序号}");
                    结果.AppendLine($"  灯带长度: {灯带长度:F2}m");
                    结果.AppendLine($"  线材长度: {线材长度:F2}m");
                    结果.AppendLine($"  标签码1: {标签码1}");
                    结果.AppendLine($"  标签码2: {标签码2}");
                    结果.AppendLine($"  标签码3: {标签码3}");
                    结果.AppendLine($"  标签码4: {标签码4}");
                    结果.AppendLine("----------------------------------------");
                }
            }

            保存日志(结果.ToString(), "不同标签码1不能混装方案结果", Path.GetFileName(当前处理文件名));

            string outputDirectory = Path.Combine(Application.StartupPath, "输出结果", 变量.订单编号);
            if (!Directory.Exists(outputDirectory))
            {
                Directory.CreateDirectory(outputDirectory);
            }

            string outputFilePath = Path.Combine(outputDirectory, "不同标签码1不能混装_" + Path.GetFileName(当前处理文件名));
            using (ExcelPackage package = new ExcelPackage(new FileInfo(outputFilePath)))
            {
                int 包装序号 = 1;

                foreach (var 包装 in 方案结果.多条包装列表)
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets.Add($"第{包装序号}盒");

                    worksheet.Cells[1, 1].Value = "序号";
                    worksheet.Cells[1, 2].Value = "条数";
                    worksheet.Cells[1, 3].Value = "米数";
                    worksheet.Cells[1, 4].Value = "标签码1";
                    worksheet.Cells[1, 5].Value = "标签码2";
                    worksheet.Cells[1, 6].Value = "标签码3";
                    worksheet.Cells[1, 7].Value = "标签码4";
                    worksheet.Cells[1, 8].Value = "线长";
                    worksheet.Cells[1, 9].Value = "包装编码";
                    worksheet.Cells[1, 13].Value = "客户型号";
                    worksheet.Cells[1, 14].Value = "标签显示长度";

                    int row = 2;
                    foreach (var 灯带 in 包装.包含灯带)
                    {
                        worksheet.Cells[row, 1].Value = 灯带.序号;
                        worksheet.Cells[row, 2].Value = 1;
                        worksheet.Cells[row, 3].Value = Math.Round(灯带.灯带长度, 3);
                        worksheet.Cells[row, 4].Value = 灯带.标签码1;
                        worksheet.Cells[row, 5].Value = 灯带.标签码2;
                        worksheet.Cells[row, 6].Value = 灯带.标签码3;
                        worksheet.Cells[row, 7].Value = 灯带.标签码4;
                        worksheet.Cells[row, 8].Value = Math.Round(灯带.线材长度, 3);
                        worksheet.Cells[row, 9].Value = 包装.物料码;
                        worksheet.Cells[row, 13].Value = 灯带.客户型号;
                        worksheet.Cells[row, 14].Value = 灯带.标签显示长度;

                        using (var range = worksheet.Cells[row, 1, row, 9])
                        {
                            range.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                            range.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                            range.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                            range.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        }

                        row++;
                    }

                    包装序号++;
                }

                foreach (var 包装 in 方案结果.单条包装列表)
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets.Add($"第{包装序号}盒");

                    worksheet.Cells[1, 1].Value = "序号";
                    worksheet.Cells[1, 2].Value = "条数";
                    worksheet.Cells[1, 3].Value = "米数";
                    worksheet.Cells[1, 4].Value = "标签码1";
                    worksheet.Cells[1, 5].Value = "标签码2";
                    worksheet.Cells[1, 6].Value = "标签码3";
                    worksheet.Cells[1, 7].Value = "标签码4";
                    worksheet.Cells[1, 8].Value = "线长";
                    worksheet.Cells[1, 9].Value = "包装编码";
                    worksheet.Cells[1, 13].Value = "客户型号";
                    worksheet.Cells[1, 14].Value = "标签显示长度";

                    worksheet.Cells[2, 1].Value = 包装.序号;
                    worksheet.Cells[2, 2].Value = 1;
                    worksheet.Cells[2, 3].Value = Math.Round(包装.灯带长度, 3);
                    worksheet.Cells[2, 4].Value = 包装.标签码1;
                    worksheet.Cells[2, 5].Value = 包装.标签码2;
                    worksheet.Cells[2, 6].Value = 包装.标签码3;
                    worksheet.Cells[2, 7].Value = 包装.标签码4;
                    worksheet.Cells[2, 8].Value = Math.Round(包装.线材长度, 3);
                    worksheet.Cells[2, 9].Value = 包装.物料码;
                    worksheet.Cells[2, 13].Value = 包装.客户型号;
                    worksheet.Cells[2, 14].Value = 包装.标签显示长度;

                    using (var range = worksheet.Cells[2, 1, 2, 9])
                    {
                        range.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                        range.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                        range.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                        range.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    }

                    包装序号++;
                }

                package.Save();
            }

            // 创建包装方案对象，用于生成流转单
            包装方案 选定方案 = new 包装方案
            {
                多条包装列表 = 方案结果.多条包装列表,
                单条包装列表 = 方案结果.单条包装列表,
                方案名称 = "不同标签码1不能混装方案"
            };

            构建标签码1分组字典(选定方案);

            //string dictMsg = string.Join("\n", 标签码1分组字典.Select(kv => $"物料码: {kv.Key} -> 标签码1: [{string.Join(", ", kv.Value)}]"));
            //MessageBox.Show(dictMsg, "标签码1分组字典内容");

            // 生成包装材料需求流转单
            生成包装材料需求流转单_4(outputDirectory, "不同标签码1不能混装_" + Path.GetFileName(当前处理文件名), 选定方案);

            uiTextBox_状态.AppendText($"不同标签码1不能混装方案已生成，共需要{方案包装数}个包装。" + Environment.NewLine);
            uiTextBox_状态.AppendText("------------------------" + Environment.NewLine);
        }

        private void 构建标签码1分组字典(包装方案 方案)
        {
            标签码1分组字典.Clear();
            // 多条包装：每个“包装”就是一盒
            foreach (var 包装 in 方案.多条包装列表)
            {
                string 物料码 = 包装.物料码;
                // 取这一盒的标签码1（假设一盒内所有灯带标签码1都一样，取第一个即可）
                string 标签码1 = 包装.包含灯带.FirstOrDefault().标签码1;
                if (!标签码1分组字典.ContainsKey(物料码))
                    标签码1分组字典[物料码] = new List<string>();
                标签码1分组字典[物料码].Add(标签码1);
            }
            // 单条包装：每个包装就是一盒
            foreach (var 包装 in 方案.单条包装列表)
            {
                string 物料码 = 包装.物料码;
                string 标签码1 = 包装.标签码1;
                if (!标签码1分组字典.ContainsKey(物料码))
                    标签码1分组字典[物料码] = new List<string>();
                标签码1分组字典[物料码].Add(标签码1);
            }
        }

        private void 显示不同标签码2不能混装方案结果(List<包装数据> 匹配包装列表, List<(string 序号, double 灯带长度, double 线材长度, string 标签码1, string 标签码2, string 标签码3, string 标签码4, string 客户型号, string 标签显示长度)> 长度列表, string 型号)
        {
            // 排除 Luminii 的包装
            var 非Luminii包装列表 = 匹配包装列表.Where(p => !p.包装名称.Contains("Luminii")).ToList();

            var 方案结果 = 执行不同标签码2不能混装方案(非Luminii包装列表, 长度列表, 型号);

            int 方案包装数 = 方案结果.多条包装列表.Count + 方案结果.单条包装列表.Count;

            StringBuilder 结果 = new StringBuilder();
            结果.AppendLine("不同标签码2不能混装方案结果：");
            结果.AppendLine($"需要包装数：{方案包装数}");
            结果.AppendLine($"  多条包装：{方案结果.多条包装列表.Count} 个");
            结果.AppendLine($"  单条包装：{方案结果.单条包装列表.Count} 个");
            结果.AppendLine();

            if (方案结果.多条包装列表.Any())
            {
                结果.AppendLine("多条包装详情：");
                foreach (var 包装 in 方案结果.多条包装列表)
                {
                    结果.AppendLine($"包装名称: {包装.包装名称}");
                    结果.AppendLine($"物料码: {包装.物料码}");
                    结果.AppendLine("包含灯带:");
                    double 包装内总灯带长度 = 0;
                    double 包装内总线材长度 = 0;

                    foreach (var 灯带 in 包装.包含灯带)
                    {
                        结果.AppendLine($"  序号: {灯带.序号}");
                        结果.AppendLine($"    灯带长度: {灯带.灯带长度:F2}m");
                        结果.AppendLine($"    线材长度: {灯带.线材长度:F2}m");
                        结果.AppendLine($"    标签码1: {灯带.标签码1}");
                        结果.AppendLine($"    标签码2: {灯带.标签码2}");
                        结果.AppendLine($"    标签码3: {灯带.标签码3}");
                        结果.AppendLine($"    标签码4: {灯带.标签码4}");
                        包装内总灯带长度 += 灯带.灯带长度;
                        包装内总线材长度 += 灯带.线材长度;
                    }
                    结果.AppendLine($"包装内总灯带长度: {包装内总灯带长度:F2}m");
                    结果.AppendLine($"包装内总线材长度: {包装内总线材长度:F2}m");
                    结果.AppendLine("----------------------------------------");
                }
            }

            if (方案结果.单条包装列表.Any())
            {
                结果.AppendLine("\n单条包装详情：");
                foreach (var (包装名称, 物料码, 序号, 灯带长度, 线材长度, 标签码1, 标签码2, 标签码3, 标签码4, 客户型号, 标签显示长度) in 方案结果.单条包装列表)
                {
                    结果.AppendLine($"包装名称: {包装名称}");
                    结果.AppendLine($"物料码: {物料码}");
                    结果.AppendLine($"序号: {序号}");
                    结果.AppendLine($"  灯带长度: {灯带长度:F2}m");
                    结果.AppendLine($"  线材长度: {线材长度:F2}m");
                    结果.AppendLine($"  标签码1: {标签码1}");
                    结果.AppendLine($"  标签码2: {标签码2}");
                    结果.AppendLine($"  标签码3: {标签码3}");
                    结果.AppendLine($"  标签码4: {标签码4}");
                    结果.AppendLine("----------------------------------------");
                }
            }

            保存日志(结果.ToString(), "不同标签码2不能混装方案结果", Path.GetFileName(当前处理文件名));

            string outputDirectory = Path.Combine(Application.StartupPath, "输出结果", 变量.订单编号);
            if (!Directory.Exists(outputDirectory))
            {
                Directory.CreateDirectory(outputDirectory);
            }

            string outputFilePath = Path.Combine(outputDirectory, "不同标签码2不能混装_" + Path.GetFileName(当前处理文件名));
            using (ExcelPackage package = new ExcelPackage(new FileInfo(outputFilePath)))
            {
                int 包装序号 = 1;

                foreach (var 包装 in 方案结果.多条包装列表)
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets.Add($"第{包装序号}盒");

                    worksheet.Cells[1, 1].Value = "序号";
                    worksheet.Cells[1, 2].Value = "条数";
                    worksheet.Cells[1, 3].Value = "米数";
                    worksheet.Cells[1, 4].Value = "标签码1";
                    worksheet.Cells[1, 5].Value = "标签码2";
                    worksheet.Cells[1, 6].Value = "标签码3";
                    worksheet.Cells[1, 7].Value = "标签码4";
                    worksheet.Cells[1, 8].Value = "线长";
                    worksheet.Cells[1, 9].Value = "包装编码";
                    worksheet.Cells[1, 13].Value = "客户型号";
                    worksheet.Cells[1, 14].Value = "标签显示长度";

                    int row = 2;
                    foreach (var 灯带 in 包装.包含灯带)
                    {
                        worksheet.Cells[row, 1].Value = 灯带.序号;
                        worksheet.Cells[row, 2].Value = 1;
                        worksheet.Cells[row, 3].Value = Math.Round(灯带.灯带长度, 3);
                        worksheet.Cells[row, 4].Value = 灯带.标签码1;
                        worksheet.Cells[row, 5].Value = 灯带.标签码2;
                        worksheet.Cells[row, 6].Value = 灯带.标签码3;
                        worksheet.Cells[row, 7].Value = 灯带.标签码4;
                        worksheet.Cells[row, 8].Value = Math.Round(灯带.线材长度, 3);
                        worksheet.Cells[row, 9].Value = 包装.物料码;
                        worksheet.Cells[row, 13].Value = 灯带.客户型号;
                        worksheet.Cells[row, 14].Value = 灯带.标签显示长度;

                        using (var range = worksheet.Cells[row, 1, row, 9])
                        {
                            range.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                            range.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                            range.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                            range.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        }

                        row++;
                    }

                    包装序号++;
                }

                foreach (var 包装 in 方案结果.单条包装列表)
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets.Add($"第{包装序号}盒");

                    worksheet.Cells[1, 1].Value = "序号";
                    worksheet.Cells[1, 2].Value = "条数";
                    worksheet.Cells[1, 3].Value = "米数";
                    worksheet.Cells[1, 4].Value = "标签码1";
                    worksheet.Cells[1, 5].Value = "标签码2";
                    worksheet.Cells[1, 6].Value = "标签码3";
                    worksheet.Cells[1, 7].Value = "标签码4";
                    worksheet.Cells[1, 8].Value = "线长";
                    worksheet.Cells[1, 9].Value = "包装编码";
                    worksheet.Cells[1, 13].Value = "客户型号";
                    worksheet.Cells[1, 14].Value = "标签显示长度";

                    worksheet.Cells[2, 1].Value = 包装.序号;
                    worksheet.Cells[2, 2].Value = 1;
                    worksheet.Cells[2, 3].Value = Math.Round(包装.灯带长度, 3);
                    worksheet.Cells[2, 4].Value = 包装.标签码1;
                    worksheet.Cells[2, 5].Value = 包装.标签码2;
                    worksheet.Cells[2, 6].Value = 包装.标签码3;
                    worksheet.Cells[2, 7].Value = 包装.标签码4;
                    worksheet.Cells[2, 8].Value = Math.Round(包装.线材长度, 3);
                    worksheet.Cells[2, 9].Value = 包装.物料码;
                    worksheet.Cells[2, 13].Value = 包装.客户型号;
                    worksheet.Cells[2, 14].Value = 包装.标签显示长度;

                    using (var range = worksheet.Cells[2, 1, 2, 9])
                    {
                        range.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                        range.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                        range.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                        range.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    }

                    包装序号++;
                }

                package.Save();
            }

            // 创建包装方案对象，用于生成流转单
            包装方案 选定方案 = new 包装方案
            {
                多条包装列表 = 方案结果.多条包装列表,
                单条包装列表 = 方案结果.单条包装列表,
                方案名称 = "不同标签码2不能混装方案"
            };

            // 生成包装材料需求流转单
            生成包装材料需求流转单_4(outputDirectory, "不同标签码2不能混装_" + Path.GetFileName(当前处理文件名), 选定方案);

            uiTextBox_状态.AppendText($"不同标签码2不能混装方案已生成，共需要{方案包装数}个包装。" + Environment.NewLine);
            uiTextBox_状态.AppendText("------------------------" + Environment.NewLine);
        }

        private (List<(string 包装名称, string 物料码, List<(string 序号, double 灯带长度, double 线材长度, string 标签码1, string 标签码2, string 标签码3, string 标签码4, string 客户型号, string 标签显示长度)> 包含灯带)> 多条包装列表,
          List<(string 包装名称, string 物料码, string 序号, double 灯带长度, double 线材长度, string 标签码1, string 标签码2, string 标签码3, string 标签码4, string 客户型号, string 标签显示长度)> 单条包装列表)
 执行不同标签码2不能混装方案(List<包装数据> 匹配包装列表, List<(string 序号, double 灯带长度, double 线材长度, string 标签码1, string 标签码2, string 标签码3, string 标签码4, string 客户型号, string 标签显示长度)> 长度列表, string 型号)
        {
            var 非Luminii包装列表 = 匹配包装列表.Where(p => !p.包装名称.Contains("Luminii")).ToList();

            var 多条包装列表 = new List<(string 包装名称, string 物料码, List<(string 序号, double 灯带长度, double 线材长度, string 标签码1, string 标签码2, string 标签码3, string 标签码4, string 客户型号, string 标签显示长度)> 包含灯带)>();
            var 单条包装列表 = new List<(string 包装名称, string 物料码, string 序号, double 灯带长度, double 线材长度, string 标签码1, string 标签码2, string 标签码3, string 标签码4, string 客户型号, string 标签显示长度)>();

            var 四七零包装 = 非Luminii包装列表.FirstOrDefault(p => p.包装名称?.Contains("470") == true &&
                (p.包装类型?.Contains("多条") == true || p.包装类型?.Contains("2条") == true || p.包装类型?.Contains("3条") == true));
            var 六百包装 = 非Luminii包装列表.FirstOrDefault(p => p.包装名称?.Contains("600") == true &&
                (p.包装类型?.Contains("多条") == true || p.包装类型?.Contains("2条") == true || p.包装类型?.Contains("3条") == true));

            var 单条包装列表数据 = 非Luminii包装列表
                .Where(p => p.包装类型?.Contains("单条") == true)
                .OrderBy(p =>
                {
                    if (p.包装名称?.Contains("330") == true) return 1;
                    if (p.包装名称?.Contains("470") == true) return 2;
                    if (p.包装名称?.Contains("600") == true) return 3;
                    return 4;
                })
                .ToList();

            if (四七零包装 == null && 六百包装 == null)
            {
                MessageBox.Show("未找到合适的多条包装方案", "提示");
                return (多条包装列表, 单条包装列表);
            }

            double 权重值 = uiTextBox_权重值.Text.ToDouble();

            var 标签码2分组 = 长度列表.GroupBy(x => x.标签码2).ToList();

            foreach (var 组 in 标签码2分组)
            {
                var 待处理灯带 = 组.ToList();
                string 当前标签码2 = 待处理灯带[0].标签码2;

                while (待处理灯带.Any())
                {
                    var 当前组 = new List<(string 序号, double 灯带长度, double 线材长度, string 标签码1, string 标签码2, string 标签码3, string 标签码4, string 客户型号, string 标签显示长度)>();
                    double current灯带米数 = 0;
                    double current线材长度 = 0;

                    var 第一个灯带 = 待处理灯带[0];
                    当前组.Add(第一个灯带);
                    current灯带米数 = 第一个灯带.灯带长度;
                    current线材长度 = 第一个灯带.线材长度;
                    待处理灯带.RemoveAt(0);

                    int 最大迭代次数 = 1000;
                    int 当前迭代 = 0;

                    while (待处理灯带.Any() && 当前迭代 < 最大迭代次数)
                    {
                        当前迭代++;
                        var 下一个灯带 = 待处理灯带[0];
                        if (下一个灯带.标签码2 != 当前标签码2)
                        {
                            break;
                        }
                        double 预计累计灯带米数 = current灯带米数 + 下一个灯带.灯带长度;
                        double 预计累计线材长度 = current线材长度 + 下一个灯带.线材长度;
                        var 当前灯带组 = 当前组.Concat(new[] { 下一个灯带 }).ToList();

                        bool 可以用470 = 四七零包装 != null && 检查包装是否可用(四七零包装, 型号, 预计累计灯带米数, 预计累计线材长度, 权重值, 当前灯带组);
                        bool 可以用600 = 六百包装 != null && 检查包装是否可用(六百包装, 型号, 预计累计灯带米数, 预计累计线材长度, 权重值, 当前灯带组);

                        if (可以用470 || 可以用600)
                        {
                            当前组.Add(下一个灯带);
                            current灯带米数 = 预计累计灯带米数;
                            current线材长度 = 预计累计线材长度;
                            待处理灯带.RemoveAt(0);
                        }
                        else
                        {
                            break;
                        }
                    }

                    if (当前组.Count == 1)
                    {
                        bool 找到单条包装 = false;
                        foreach (var 包装 in 单条包装列表数据)
                        {
                            var 单条产品信息 = 包装.支持产品型号列表?.FirstOrDefault(p => p.产品型号 == 型号);
                            if (单条产品信息 == null) continue;
                            var 灯带 = 当前组[0];
                            var 单条灯带列表 = new List<(string 序号, double 灯带长度, double 线材长度, string 标签码1, string 标签码2, string 标签码3, string 标签码4, string 客户型号, string 标签显示长度)> { 灯带 };
                            bool 忽略线长限制 = !需要计算线长(单条灯带列表);
                            if (灯带.灯带长度 <= 单条产品信息.灯带最大长度 &&
                                灯带.灯带长度 >= 单条产品信息.最低阈值 &&
                                (忽略线长限制 || 灯带.线材长度 <= 单条产品信息.线材最大长度))
                            {
                                单条包装列表.Add((包装.包装名称, 包装.物料码, 灯带.序号, 灯带.灯带长度, 灯带.线材长度, 灯带.标签码1, 灯带.标签码2, 灯带.标签码3, 灯带.标签码4, 灯带.客户型号, 灯带.标签显示长度));
                                找到单条包装 = true;
                                break;
                            }
                        }
                        if (!找到单条包装)
                        {
                            bool 可以用470单条 = false;
                            if (四七零包装 != null)
                            {
                                var 单条灯带列表 = new List<(string 序号, double 灯带长度, double 线材长度, string 标签码1, string 标签码2, string 标签码3, string 标签码4, string 客户型号, string 标签显示长度)> { 当前组[0] };
                                可以用470单条 = 检查包装是否可用(四七零包装, 型号, 当前组[0].灯带长度, 当前组[0].线材长度, 权重值, 单条灯带列表);
                                if (可以用470单条)
                                {
                                    多条包装列表.Add((四七零包装.包装名称, 四七零包装.物料码, 当前组));
                                }
                            }
                            if (!可以用470单条 && 六百包装 != null)
                            {
                                var 单条灯带列表 = new List<(string 序号, double 灯带长度, double 线材长度, string 标签码1, string 标签码2, string 标签码3, string 标签码4, string 客户型号, string 标签显示长度)> { 当前组[0] };
                                bool 可以用600单条 = 检查包装是否可用(六百包装, 型号, 当前组[0].灯带长度, 当前组[0].线材长度, 权重值, 单条灯带列表);
                                if (可以用600单条)
                                {
                                    多条包装列表.Add((六百包装.包装名称, 六百包装.物料码, 当前组));
                                }
                                else
                                {
                                    throw new Exception($"无法找到合适的包装方案处理灯带 {当前组[0].序号}");
                                }
                            }
                        }
                    }
                    else
                    {
                        bool 可以用470多条 = false;
                        if (四七零包装 != null)
                        {
                            可以用470多条 = 检查包装是否可用(四七零包装, 型号, current灯带米数, current线材长度, 权重值, 当前组);
                        }
                        if (可以用470多条)
                        {
                            多条包装列表.Add((四七零包装.包装名称, 四七零包装.物料码, 当前组));
                        }
                        else if (六百包装 != null)
                        {
                            bool 可以用600多条 = 检查包装是否可用(六百包装, 型号, current灯带米数, current线材长度, 权重值, 当前组);
                            if (可以用600多条)
                            {
                                多条包装列表.Add((六百包装.包装名称, 六百包装.物料码, 当前组));
                            }
                            else
                            {
                                throw new Exception("无法找到合适的多条包装");
                            }
                        }
                        else
                        {
                            throw new Exception("无法找到合适的多条包装");
                        }
                    }
                }
            }
            return (多条包装列表, 单条包装列表);
        }

        private bool 需要计算线长(List<(string 序号, double 灯带长度, double 线材长度, string 标签码1, string 标签码2, string 标签码3, string 标签码4, string 客户型号, string 标签显示长度)> 灯带列表)
        {
            if (灯带列表.Count > 4) return true;  // 超过4条需要计算
            return 灯带列表.Any(x => x.线材长度 > 5.0);  // 有任何一条超过5米需要计算
        }

        private bool 检查包装是否可用(包装数据 包装, string 型号, double 预计累计灯带米数, double 预计累计线材长度, double 权重值, List<(string 序号, double 灯带长度, double 线材长度, string 标签码1, string 标签码2, string 标签码3, string 标签码4, string 客户型号, string 标签显示长度)> 当前灯带组)
        {
            var 产品信息 = 包装.支持产品型号列表?.FirstOrDefault(p => p.产品型号 == 型号);
            if (产品信息 == null) return false;

            double 最大灯带米数 = 产品信息.灯带最大长度;
            double 最大线材长度 = 产品信息.线材最大长度;

            double 累计灯带权重系数 = 预计累计灯带米数 / 最大灯带米数;
            double 累计线材权重系数 = 预计累计线材长度 / 最大线材长度;

            double 调整后最大线材长度 = 最大线材长度 * (1 - 累计灯带权重系数 * 权重值);
            double 调整后最大灯带米数 = 最大灯带米数 * (1 - 累计线材权重系数 * 权重值);

            bool 忽略线长限制 = !需要计算线长(当前灯带组);

            return 预计累计灯带米数 <= 调整后最大灯带米数 &&
                   (忽略线长限制 || 预计累计线材长度 <= 调整后最大线材长度);
        }

        private void 显示方案比较结果_2(
(List<(string 包装名称, string 物料码, List<(string 序号, double 灯带长度, double 线材长度, string 标签码1, string 标签码2, string 标签码3, string 标签码4, string 客户型号, string 标签显示长度)> 包含灯带)> 多条包装列表,
 List<(string 包装名称, string 物料码, string 序号, double 灯带长度, double 线材长度, string 标签码1, string 标签码2, string 标签码3, string 标签码4, string 客户型号, string 标签显示长度)> 单条包装列表) 方案一,
(List<(string 包装名称, string 物料码, List<(string 序号, double 灯带长度, double 线材长度, string 标签码1, string 标签码2, string 标签码3, string 标签码4, string 客户型号, string 标签显示长度)> 包含灯带)> 多条包装列表,
 List<(string 包装名称, string 物料码, string 序号, double 灯带长度, double 线材长度, string 标签码1, string 标签码2, string 标签码3, string 标签码4, string 客户型号, string 标签显示长度)> 单条包装列表) 方案二,
(List<(string 包装名称, string 物料码, List<(string 序号, double 灯带长度, double 线材长度, string 标签码1, string 标签码2, string 标签码3, string 标签码4, string 客户型号, string 标签显示长度)> 包含灯带)> 多条包装列表,
 List<(string 包装名称, string 物料码, string 序号, double 灯带长度, double 线材长度, string 标签码1, string 标签码2, string 标签码3, string 标签码4, string 客户型号, string 标签显示长度)> 单条包装列表) 节约方案)
        {
            int 方案一包装数 = 方案一.多条包装列表.Count + 方案一.单条包装列表.Count;
            int 方案二包装数 = 方案二.多条包装列表.Count + 方案二.单条包装列表.Count;
            int 节约方案包装数 = 节约方案.多条包装列表.Count + 节约方案.单条包装列表.Count;

            StringBuilder 比较结果 = new StringBuilder();
            比较结果.AppendLine("包装方案比较：");
            比较结果.AppendLine($"方案一（优先单条包装）需要包装数：{方案一包装数}");
            比较结果.AppendLine($"  多条包装：{方案一.多条包装列表.Count} 个");
            比较结果.AppendLine($"  单条包装：{方案一.单条包装列表.Count} 个");
            比较结果.AppendLine();

            比较结果.AppendLine($"方案二（优先多条包装）需要包装数：{方案二包装数}");
            比较结果.AppendLine($"  多条包装：{方案二.多条包装列表.Count} 个");
            比较结果.AppendLine($"  单条包装：{方案二.单条包装列表.Count} 个");
            比较结果.AppendLine();

            比较结果.AppendLine($"节约方案需要包装数：{节约方案包装数}");
            比较结果.AppendLine($"  多条包装：{节约方案.多条包装列表.Count} 个");
            比较结果.AppendLine($"  单条包装：{节约方案.单条包装列表.Count} 个");
            比较结果.AppendLine();

            比较结果.AppendLine("----------------------------------------");
            比较结果.AppendLine("最优包装方案详情：");

            // 确定最优方案
            var 最优方案包装数 = Math.Min(Math.Min(方案一包装数, 方案二包装数), 节约方案包装数);
            if (最优方案包装数 == 方案一包装数)
            {
                比较结果.AppendLine("采用方案一（优先单条包装）");
                显示包装方案详情_3(比较结果, 方案一);
            }
            else if (最优方案包装数 == 方案二包装数)
            {
                比较结果.AppendLine("采用方案二（优先多条包装）");
                显示包装方案详情_3(比较结果, 方案二);
            }
            else
            {
                比较结果.AppendLine("采用节约方案");
                显示包装方案详情_3(比较结果, 节约方案);
            }

            // 创建结果窗口
            Form 结果窗口 = new Form
            {
                Width = 800,
                Height = 600,
                Text = "包装方案比较结果"
            };

            RichTextBox 文本框 = new RichTextBox
            {
                Dock = DockStyle.Fill,
                ReadOnly = true,
                Font = new Font("Microsoft YaHei", 10),
                Text = 比较结果.ToString()
            };

            //结果窗口.Controls.Add(文本框);
            //结果窗口.ShowDialog();
            保存日志(比较结果.ToString(), "包装方案比较结果", Path.GetFileName(当前处理文件名));

            // 确定最优方案用于后续处理

            //常规方案
            var 最优方案 = 最优方案包装数 == 方案一包装数 ? 方案一 :
           最优方案包装数 == 方案二包装数 ? 方案二 : 节约方案;

            // 创建Excel文件(最优方案组合结果EXCEL)
            string 文件夹路径 = Path.Combine(Application.StartupPath, "输出结果", 变量.订单编号);
            if (!Directory.Exists(文件夹路径))
            {
                Directory.CreateDirectory(文件夹路径);
            }

            string 文件路径 = Path.Combine(文件夹路径, Path.GetFileName(当前处理文件名));
            using (ExcelPackage package = new ExcelPackage(new FileInfo(文件路径)))
            {
                int 包装序号 = 1;

                // 处理多条包装
                foreach (var 包装 in 最优方案.多条包装列表)
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets.Add($"第{包装序号}盒");

                    // 设置表头
                    worksheet.Cells[1, 1].Value = "序号";
                    worksheet.Cells[1, 2].Value = "条数";
                    worksheet.Cells[1, 3].Value = "米数";
                    worksheet.Cells[1, 4].Value = "标签码1";
                    worksheet.Cells[1, 5].Value = "标签码2";
                    worksheet.Cells[1, 6].Value = "标签码3";
                    worksheet.Cells[1, 7].Value = "标签码4";
                    worksheet.Cells[1, 8].Value = "线长";
                    worksheet.Cells[1, 9].Value = "包装编码";
                    worksheet.Cells[1, 13].Value = "客户型号"; // 新增
                    worksheet.Cells[1, 14].Value = "标签显示长度";

                    // 填充数据
                    int row = 2;
                    foreach (var 灯带 in 包装.包含灯带)
                    {
                        worksheet.Cells[row, 1].Value = 灯带.序号;
                        worksheet.Cells[row, 2].Value = 1;
                        worksheet.Cells[row, 3].Value = Math.Round(灯带.灯带长度, 3);
                        worksheet.Cells[row, 4].Value = 灯带.标签码1;
                        worksheet.Cells[row, 5].Value = 灯带.标签码2;
                        worksheet.Cells[row, 6].Value = 灯带.标签码3;
                        worksheet.Cells[row, 7].Value = 灯带.标签码4;
                        worksheet.Cells[row, 8].Value = Math.Round(灯带.线材长度, 3);
                        worksheet.Cells[row, 9].Value = 包装.物料码;
                        worksheet.Cells[row, 13].Value = 灯带.客户型号; // 新增
                        worksheet.Cells[row, 14].Value = 灯带.标签显示长度;

                        // 设置单元格样式
                        using (var range = worksheet.Cells[row, 1, row, 14])
                        {
                            range.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                            range.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                            range.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                            range.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        }

                        row++;
                    }

                    包装序号++;
                }

                // 处理单条包装
                foreach (var 包装 in 最优方案.单条包装列表)
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets.Add($"第{包装序号}盒");

                    // 设置表头
                    worksheet.Cells[1, 1].Value = "序号";
                    worksheet.Cells[1, 2].Value = "条数";
                    worksheet.Cells[1, 3].Value = "米数";
                    worksheet.Cells[1, 4].Value = "标签码1";
                    worksheet.Cells[1, 5].Value = "标签码2";
                    worksheet.Cells[1, 6].Value = "标签码3";
                    worksheet.Cells[1, 7].Value = "标签码4";
                    worksheet.Cells[1, 8].Value = "线长";
                    worksheet.Cells[1, 9].Value = "包装编码";
                    worksheet.Cells[1, 13].Value = "客户型号";
                    worksheet.Cells[1, 14].Value = "标签显示长度";

                    // 填充单条包装数据
                    worksheet.Cells[2, 1].Value = 包装.序号;
                    worksheet.Cells[2, 2].Value = 1;
                    worksheet.Cells[2, 3].Value = Math.Round(包装.灯带长度, 3);
                    worksheet.Cells[2, 4].Value = 包装.标签码1;
                    worksheet.Cells[2, 5].Value = 包装.标签码2;
                    worksheet.Cells[2, 6].Value = 包装.标签码3;
                    worksheet.Cells[2, 7].Value = 包装.标签码4;
                    worksheet.Cells[2, 8].Value = Math.Round(包装.线材长度, 3);
                    worksheet.Cells[2, 9].Value = 包装.物料码;
                    worksheet.Cells[2, 13].Value = 包装.客户型号;
                    worksheet.Cells[2, 14].Value = 包装.标签显示长度;

                    // 设置单元格样式
                    using (var range = worksheet.Cells[2, 1, 2, 14])
                    {
                        range.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                        range.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                        range.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                        range.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    }

                    包装序号++;
                }

                package.Save();
            }

            包装方案 选定方案 = new 包装方案
            {
                多条包装列表 = 最优方案.多条包装列表,
                单条包装列表 = 最优方案.单条包装列表
            };
            生成包装材料需求流转单_4(文件夹路径, 当前处理文件名, 选定方案);
        }

        private void 显示包装方案详情_3(StringBuilder sb,
(List<(string 包装名称, string 物料码, List<(string 序号, double 灯带长度, double 线材长度, string 标签码1, string 标签码2, string 标签码3, string 标签码4, string 客户型号, string 标签显示长度)> 包含灯带)> 多条包装列表,
 List<(string 包装名称, string 物料码, string 序号, double 灯带长度, double 线材长度, string 标签码1, string 标签码2, string 标签码3, string 标签码4, string 客户型号, string 标签显示长度)> 单条包装列表) 方案)
        {
            // 显示单条包装
            if (方案.单条包装列表.Count > 0)
            {
                sb.AppendLine("\n单条包装：");
                foreach (var (包装名称, 物料码, 序号, 灯带长度, 线材长度, 标签码1, 标签码2, 标签码3, 标签码4, 客户型号, 标签显示长度) in 方案.单条包装列表)
                {
                    sb.AppendLine($"包装名称: {包装名称}");
                    sb.AppendLine($"物料码: {物料码}");
                    sb.AppendLine($"序号: {序号}");
                    sb.AppendLine($"  灯带长度: {灯带长度:F2}m");
                    sb.AppendLine($"  线材长度: {线材长度:F2}m");
                    sb.AppendLine("----------------------------------------");
                }
            }

            // 显示多条包装
            if (方案.多条包装列表.Count > 0)
            {
                sb.AppendLine("\n多条包装：");
                foreach (var (包装名称, 物料码, 包含灯带) in 方案.多条包装列表)
                {
                    sb.AppendLine($"包装名称: {包装名称}");
                    sb.AppendLine($"物料码: {物料码}");
                    sb.AppendLine("包含灯带:");
                    double 包装内总灯带长度 = 0;
                    double 包装内总线材长度 = 0;

                    foreach (var (序号, 灯带长度, 线材长度, 标签码1, 标签码2, 标签码3, 标签码4, 客户型号, 标签显示长度) in 包含灯带)
                    {
                        sb.AppendLine($"  序号: {序号}");
                        sb.AppendLine($"    灯带长度: {灯带长度:F2}m");
                        sb.AppendLine($"    线材长度: {线材长度:F2}m");
                        包装内总灯带长度 += 灯带长度;
                        包装内总线材长度 += 线材长度;
                    }
                    sb.AppendLine($"包装内总灯带长度: {包装内总灯带长度:F2}m");
                    sb.AppendLine($"包装内总线材长度: {包装内总线材长度:F2}m");
                    sb.AppendLine("----------------------------------------");
                }
            }
        }

        private void 生成包装材料需求流转单_4(string 文件夹路径, string 当前处理文件名, 包装方案 最优方案)
        {
            var 仓位字典 = 加载仓位数据();
            string 汇总文件路径 = Path.Combine(文件夹路径, "包装材料需求流转单.xlsx");
            FileInfo 汇总文件信息 = new FileInfo(汇总文件路径);

            // 创建一个字典来统计每个包装物料码的使用数量
            Dictionary<string, int> 包装物料码统计 = new Dictionary<string, int>();

            // 统计多条包装的物料码使用数量
            foreach (var 包装 in 最优方案.多条包装列表)
            {
                if (!string.IsNullOrEmpty(包装.物料码))
                {
                    if (!包装物料码统计.ContainsKey(包装.物料码))
                    {
                        包装物料码统计[包装.物料码] = 0;
                    }
                    包装物料码统计[包装.物料码]++;
                }
            }

            // 统计单条包装的物料码使用数量
            foreach (var 包装 in 最优方案.单条包装列表)
            {
                if (!string.IsNullOrEmpty(包装.物料码))
                {
                    if (!包装物料码统计.ContainsKey(包装.物料码))
                    {
                        包装物料码统计[包装.物料码] = 0;
                    }
                    包装物料码统计[包装.物料码]++;
                }
            }

            using (ExcelPackage 汇总包 = new ExcelPackage(汇总文件信息))
            {
                ExcelWorksheet 汇总表;
                if (汇总包.Workbook.Worksheets.Any(ws => ws.Name == "包装材料需求流转单"))
                {
                    汇总表 = 汇总包.Workbook.Worksheets["包装材料需求流转单"];

                    // 获取已有数据的范围
                    int lastRow = 汇总表.Dimension?.End.Row ?? 6;

                    // 删除匹配的现有数据
                    for (int row = 7; row <= lastRow; row++)
                    {
                        string 现有文件名 = 汇总表.Cells[row, 11].Text;
                        if (现有文件名 == Path.GetFileName(当前处理文件名))
                        {
                            汇总表.DeleteRow(row);
                            row--; // 调整行索引
                            lastRow--; // 调整总行数
                        }
                    }
                }
                else
                {
                    汇总表 = 汇总包.Workbook.Worksheets.Add("包装材料需求流转单");

                    // 设置标题（合并A2-I2）
                    汇总表.Cells["A2:I2"].Merge = true;
                    汇总表.Cells["A2"].Value = "包装材料需求流转单";
                    汇总表.Cells["A2"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                    汇总表.Cells["A2"].Style.Font.Bold = true;

                    // 第3行设置
                    汇总表.Cells["A3"].Value = "订单号:";
                    汇总表.Cells["B3"].Value = 变量.订单编号;
                    汇总表.Cells["C3:D3"].Merge = true;
                    汇总表.Cells["C3"].Value = "客户代码:" + 变量.客户代码;
                    汇总表.Cells["E3:I3"].Merge = true;
                    汇总表.Cells["E3"].Value = "完成时间:" + 变量.完成时间;

                    // 第4行设置
                    汇总表.Cells["A4:B4"].Merge = true;
                    string 制单日期 = DateTime.Now.ToString("yyyy.MM.dd");
                    汇总表.Cells["A4"].Value = "制单日期:" + 制单日期;

                    string MAC地址 = 获取MAC地址();
                    汇总表.Cells["C4:D4"].Merge = true;
                    if (MAC地址 == "98:EE:CB:91:F4:72") { 汇总表.Cells["C4"].Value = "制单人:张水升"; }
                    else if (MAC地址 == "04:D9:C8:BE:16:4F") { 汇总表.Cells["C4"].Value = "制单人:钟珊玲"; }
                    else if (MAC地址 == "04:D9:C8:BE:3F:2A") { 汇总表.Cells["C4"].Value = "制单人:刘丽 "; }
                    else if (MAC地址 == "44:39:C4:55:F7:BF") { 汇总表.Cells["C4"].Value = "制单人:何秀群 "; }
                    else if (MAC地址 == "04:D9:C8:BE:3E:C8") { 汇总表.Cells["C4"].Value = "制单人:巫艳红 "; }
                    else { 汇总表.Cells["C4"].Value = "制单人:未登记"; }

                    汇总表.Cells["E4:G4"].Merge = true;
                    汇总表.Cells["E4"].Value = "业务员:" + 变量.业务员;
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
                    tableRange.Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                    tableRange.Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                    tableRange.Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                    tableRange.Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;

                    // 设置所有单元格的内部边框
                    for (int row = 2; row <= 6; row++)
                    {
                        for (int col = 1; col <= 9; col++)
                        {
                            var cell = 汇总表.Cells[row, col];
                            cell.Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                        }
                    }

                    // 设置字体和对齐方式
                    tableRange.Style.Font.Name = "微软雅黑";
                    tableRange.Style.Font.Size = 10;
                    tableRange.Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;

                    // 设置行高
                    汇总表.Row(2).Height = 33; // 标题行
                    汇总表.Row(3).Height = 24.7;
                    汇总表.Row(4).Height = 24.7;
                    汇总表.Row(5).Height = 24.7;
                    汇总表.Row(6).Height = 34.5;

                    // 设置列宽
                    汇总表.Column(1).Width = 6.86;
                    汇总表.Column(2).Width = 16.29;
                    汇总表.Column(3).Width = 20.43;
                    汇总表.Column(4).Width = 29;
                    汇总表.Column(5).Width = 6.29;
                    汇总表.Column(6).Width = 8.43;
                    汇总表.Column(7).Width = 8.43;
                    汇总表.Column(8).Width = 14.57;
                    汇总表.Column(9).Width = 10.57;

                    // 设置标题字体
                    汇总表.Cells["A2"].Style.Font.Size = 14;
                    汇总表.Cells["A2"].Style.Font.Bold = true;

                    // 设置特定单元格的字体加粗
                    var boldCells = new[] { "A3", "B3", "C3", "E3", "E4", "A4", "H4", "B4", "C4", "F4" };
                    foreach (var cell in boldCells)
                    {
                        汇总表.Cells[cell].Style.Font.Bold = true;
                    }

                    // 设置第5-6行表头的字体加粗
                    汇总表.Cells["A5:I5"].Style.Font.Bold = true;
                    汇总表.Cells["A6:I6"].Style.Font.Bold = true;
                }

                // 获取当前数据的最后一行
                int currentRow = 汇总表.Dimension?.End.Row ?? 6;
                currentRow++;

                // 遍历每个包装物料码，获取对应的包装数据
                foreach (var kvp in 包装物料码统计)
                {
                    string 物料码 = kvp.Key;
                    int 使用数量 = kvp.Value;

                    if (包装数据字典.TryGetValue(物料码, out var 包装数据))
                    {
                        // 添加半成品BOM物料码信息
                        汇总表.Cells[currentRow, 1].Value = "";
                        汇总表.Cells[currentRow, 2].Value = "半成品BOM物料码";
                        汇总表.Cells[currentRow, 3].Value = 物料码;
                        汇总表.Cells[currentRow, 4].Value = 包装数据.包装名称;
                        汇总表.Cells[currentRow, 5].Value = 使用数量;
                        汇总表.Cells[currentRow, 6].Value = 获取仓位(物料码);
                        汇总表.Cells[currentRow, 11].Value = Path.GetFileName(当前处理文件名);

                        // 设置样式
                        var dataRange = 汇总表.Cells[currentRow, 1, currentRow, 9];
                        dataRange.Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                        dataRange.Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                        dataRange.Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                        dataRange.Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                        // 设置淡绿色背景
                        dataRange.Style.Fill.PatternType = ExcelFillStyle.Solid;
                        dataRange.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(198, 239, 206)); // 淡绿色

                        currentRow++;

                        if (变量.客户代码.Contains("13013") || 变量.客户代码.Contains("12056")) { }
                        else
                        {
                            // 添加POF热缩袋信息
                            if (!string.IsNullOrEmpty(包装数据.POF热缩袋))
                            {
                                汇总表.Cells[currentRow, 1].Value = "";
                                汇总表.Cells[currentRow, 2].Value = "POF热缩袋";
                                汇总表.Cells[currentRow, 3].Value = 包装数据.POF热缩袋;
                                汇总表.Cells[currentRow, 4].Value = 包装数据.POF热缩袋_系统尺寸;
                                汇总表.Cells[currentRow, 5].Value = 使用数量;
                                汇总表.Cells[currentRow, 6].Value = 获取仓位(包装数据.POF热缩袋);
                                汇总表.Cells[currentRow, 11].Value = Path.GetFileName(当前处理文件名);

                                dataRange = 汇总表.Cells[currentRow, 1, currentRow, 9];
                                dataRange.Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                                dataRange.Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                                dataRange.Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                                dataRange.Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;

                                currentRow++;
                            }
                        }

                        // 添加全棉织带信息
                        if (!string.IsNullOrEmpty(包装数据.全棉织带))
                        {
                            汇总表.Cells[currentRow, 1].Value = "";
                            汇总表.Cells[currentRow, 2].Value = "全棉织带";
                            汇总表.Cells[currentRow, 3].Value = 包装数据.全棉织带;
                            汇总表.Cells[currentRow, 4].Value = 包装数据.全棉织带_系统尺寸;
                            汇总表.Cells[currentRow, 5].Value = 使用数量;
                            汇总表.Cells[currentRow, 6].Value = 获取仓位(包装数据.全棉织带);
                            汇总表.Cells[currentRow, 11].Value = Path.GetFileName(当前处理文件名);

                            dataRange = 汇总表.Cells[currentRow, 1, currentRow, 9];
                            dataRange.Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            dataRange.Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            dataRange.Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            dataRange.Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;

                            currentRow++;
                        }

                        //添加通口箱信息
                        // 判断是否是Luminii客户方案
                        bool 是Luminii客户 = 当前处理文件名.StartsWith("Luminii_") || 变量.客户代码.Contains("17034");

                        // 如果是Luminii客户，先添加通口箱信息
                        if (是Luminii客户)
                        {   //MessageBox.Show(包装数据.通口箱);
                            汇总表.Cells[currentRow, 1].Value = "";
                            汇总表.Cells[currentRow, 2].Value = "通口箱";
                            汇总表.Cells[currentRow, 3].Value = 包装数据.通口箱;
                            汇总表.Cells[currentRow, 4].Value = 包装数据.通口箱尺寸;
                            汇总表.Cells[currentRow, 5].Value = 使用数量;
                            汇总表.Cells[currentRow, 6].Value = 获取仓位(包装数据.通口箱);
                            汇总表.Cells[currentRow, 11].Value = Path.GetFileName(当前处理文件名);

                            var dataRange2 = 汇总表.Cells[currentRow, 1, currentRow, 9];
                            dataRange2.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                            dataRange2.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                            dataRange2.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                            dataRange2.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                            dataRange2.Style.Fill.PatternType = ExcelFillStyle.Solid;
                            dataRange2.Style.Fill.BackgroundColor.SetColor(Color.LightGreen);

                            currentRow++;
                        }

                        // 从文件名获取当前型号
                        string 当前型号 = "";
                        if (!string.IsNullOrEmpty(当前处理文件名))
                        {
                            var match = Regex.Match(当前处理文件名, @"[FAW]\d+");
                            if (match.Success)
                            {
                                当前型号 = match.Value;
                            }
                        }

                        //// 添加纸箱信息 2025.5.13之前的旧的，可以正常使用
                        //// 计算总盒数（多条包装数量 + 单条包装数量）
                        //int 总盒数 = 使用数量;
                        //// 获取支持的最大盒数
                        //int 最大支持盒数 = 0;
                        //// 从包装数据中查找对应的产品型号信息
                        //var 产品型号信息 = 包装数据.支持产品型号列表.FirstOrDefault(x => x.产品型号 == 当前型号);
                        //if (产品型号信息 != null && !string.IsNullOrEmpty(产品型号信息.支持的盒装外箱))
                        //{
                        //    if (产品型号信息.支持的盒装外箱.Contains("五盒装"))  // 添加五盒装的判断
                        //        最大支持盒数 = 5;
                        //    else if (产品型号信息.支持的盒装外箱.Contains("三盒装"))
                        //        最大支持盒数 = 3;
                        //    else if (产品型号信息.支持的盒装外箱.Contains("二盒装"))
                        //        最大支持盒数 = 2;
                        //    else if (产品型号信息.支持的盒装外箱.Contains("单盒装"))
                        //        最大支持盒数 = 1;
                        //}

                        //int 剩余盒数 = 总盒数;  // 使用计算出的总盒数

                        //// 优先使用单盒装，如果总盒数为1
                        //if (总盒数 == 1 && 最大支持盒数 >= 1 && !string.IsNullOrEmpty(包装数据.单盒装外箱))
                        //{
                        //    汇总表.Cells[currentRow, 1].Value = "";
                        //    汇总表.Cells[currentRow, 2].Value = "纸箱";
                        //    汇总表.Cells[currentRow, 3].Value = 包装数据.单盒装外箱;
                        //    汇总表.Cells[currentRow, 4].Value = 包装数据.单盒装尺寸;
                        //    汇总表.Cells[currentRow, 5].Value = 1;
                        //    汇总表.Cells[currentRow, 6].Value = 获取仓位(包装数据.单盒装外箱);
                        //    汇总表.Cells[currentRow, 9].Value = "1盒装标准";
                        //    汇总表.Cells[currentRow, 11].Value = Path.GetFileName(当前处理文件名);

                        //    var range = 汇总表.Cells[currentRow, 1, currentRow, 9];
                        //    range.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                        //    range.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                        //    range.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                        //    range.Style.Border.Right.Style = ExcelBorderStyle.Thin;

                        //    currentRow++;
                        //    剩余盒数 = 0;
                        //}
                        //// 如果总盒数大于1，则按照最优组合方式处理
                        //else
                        //{
                        //    // 处理五盒装
                        //    if (最大支持盒数 >= 5 && !string.IsNullOrEmpty(包装数据.五盒装外箱) && 剩余盒数 >= 5)
                        //    {
                        //        int 五盒装数量 = 剩余盒数 / 5;

                        //        汇总表.Cells[currentRow, 1].Value = "";
                        //        汇总表.Cells[currentRow, 2].Value = "纸箱";
                        //        汇总表.Cells[currentRow, 3].Value = 包装数据.五盒装外箱;
                        //        汇总表.Cells[currentRow, 4].Value = 包装数据.五盒装尺寸;
                        //        汇总表.Cells[currentRow, 5].Value = 五盒装数量;
                        //        汇总表.Cells[currentRow, 6].Value = 获取仓位(包装数据.五盒装外箱);
                        //        汇总表.Cells[currentRow, 9].Value = "5盒装标准";
                        //        汇总表.Cells[currentRow, 11].Value = Path.GetFileName(当前处理文件名);

                        //        var range = 汇总表.Cells[currentRow, 1, currentRow, 9];
                        //        range.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                        //        range.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                        //        range.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                        //        range.Style.Border.Right.Style = ExcelBorderStyle.Thin;

                        //        currentRow++;
                        //        剩余盒数 %= 5;
                        //    }

                        //    // 处理三盒装
                        //    if (最大支持盒数 >= 3 && !string.IsNullOrEmpty(包装数据.三盒装外箱) && 剩余盒数 >= 3)
                        //    {
                        //        int 三盒装数量 = 剩余盒数 / 3;

                        //        汇总表.Cells[currentRow, 1].Value = "";
                        //        汇总表.Cells[currentRow, 2].Value = "纸箱";
                        //        汇总表.Cells[currentRow, 3].Value = 包装数据.三盒装外箱;
                        //        汇总表.Cells[currentRow, 4].Value = 包装数据.三盒装尺寸;
                        //        汇总表.Cells[currentRow, 5].Value = 三盒装数量;
                        //        汇总表.Cells[currentRow, 6].Value = 获取仓位(包装数据.三盒装外箱);
                        //        汇总表.Cells[currentRow, 9].Value = "3盒装标准";
                        //        汇总表.Cells[currentRow, 11].Value = Path.GetFileName(当前处理文件名);

                        //        var range = 汇总表.Cells[currentRow, 1, currentRow, 9];
                        //        range.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                        //        range.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                        //        range.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                        //        range.Style.Border.Right.Style = ExcelBorderStyle.Thin;

                        //        currentRow++;
                        //        剩余盒数 %= 3;
                        //    }

                        //    // 处理二盒装
                        //    if (最大支持盒数 >= 2 && !string.IsNullOrEmpty(包装数据.二盒装外箱) && 剩余盒数 >= 2)
                        //    {
                        //        int 二盒装数量 = 剩余盒数 / 2;

                        //        汇总表.Cells[currentRow, 1].Value = "";
                        //        汇总表.Cells[currentRow, 2].Value = "纸箱";
                        //        汇总表.Cells[currentRow, 3].Value = 包装数据.二盒装外箱;
                        //        汇总表.Cells[currentRow, 4].Value = 包装数据.二盒装尺寸;
                        //        汇总表.Cells[currentRow, 5].Value = 二盒装数量;
                        //        汇总表.Cells[currentRow, 6].Value = 获取仓位(包装数据.二盒装外箱);
                        //        汇总表.Cells[currentRow, 9].Value = "2盒装标准";
                        //        汇总表.Cells[currentRow, 11].Value = Path.GetFileName(当前处理文件名);

                        //        var range = 汇总表.Cells[currentRow, 1, currentRow, 9];
                        //        range.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                        //        range.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                        //        range.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                        //        range.Style.Border.Right.Style = ExcelBorderStyle.Thin;

                        //        currentRow++;
                        //        剩余盒数 %= 2;
                        //    }

                        //    // 处理单盒装
                        //    if (最大支持盒数 >= 1 && !string.IsNullOrEmpty(包装数据.单盒装外箱) && 剩余盒数 > 0)
                        //    {
                        //        汇总表.Cells[currentRow, 1].Value = "";
                        //        汇总表.Cells[currentRow, 2].Value = "纸箱";
                        //        汇总表.Cells[currentRow, 3].Value = 包装数据.单盒装外箱;
                        //        汇总表.Cells[currentRow, 4].Value = 包装数据.单盒装尺寸;
                        //        汇总表.Cells[currentRow, 5].Value = 剩余盒数;
                        //        汇总表.Cells[currentRow, 6].Value = 获取仓位(包装数据.单盒装外箱);
                        //        汇总表.Cells[currentRow, 9].Value = "1盒装标准";
                        //        汇总表.Cells[currentRow, 11].Value = Path.GetFileName(当前处理文件名);

                        //        var range = 汇总表.Cells[currentRow, 1, currentRow, 9];
                        //        range.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                        //        range.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                        //        range.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                        //        range.Style.Border.Right.Style = ExcelBorderStyle.Thin;

                        //        currentRow++;
                        //    }
                        //}

                        if (算法选择.Text == "不同标签码1不能混装(且外箱不能混装)")
                        {
                            // 1. 收集所有“盒子”对象（多条包装和单条包装），每个盒子是一个整体
                            var allBoxes = new List<(string 物料码, string 标签码1)>();
                            foreach (var 包装 in 最优方案.多条包装列表)
                                foreach (var 明细 in 包装.包含灯带)
                                    allBoxes.Add((包装.物料码, 明细.标签码1));
                            foreach (var 包装 in 最优方案.单条包装列表)
                                allBoxes.Add((包装.物料码, 包装.标签码1));

                            // 2. 只在这里输出一次“半成品BOM物料码”及相关材料
                            foreach (var 物料分组 in allBoxes.GroupBy(x => x.物料码))
                            {
                                物料码 = 物料分组.Key;
                                if (!包装数据字典.TryGetValue(物料码, out 包装数据)) continue;

                                // 只输出一次“半成品BOM物料码”，数量=该物料码下所有盒子的总数
                                int 总盒数 = 物料分组.Count();
                                汇总表.Cells[currentRow, 2].Value = "半成品BOM物料码";
                                汇总表.Cells[currentRow, 3].Value = 物料码;
                                汇总表.Cells[currentRow, 4].Value = 包装数据.包装名称;
                                汇总表.Cells[currentRow, 5].Value = 总盒数;
                                汇总表.Cells[currentRow, 6].Value = 获取仓位(物料码);
                                汇总表.Cells[currentRow, 11].Value = Path.GetFileName(当前处理文件名);
                                currentRow++;

                                // 输出POF热缩袋（如有）
                                if (!string.IsNullOrEmpty(包装数据.POF热缩袋))
                                {
                                    汇总表.Cells[currentRow, 2].Value = "POF热缩袋";
                                    汇总表.Cells[currentRow, 3].Value = 包装数据.POF热缩袋;
                                    汇总表.Cells[currentRow, 4].Value = 包装数据.POF热缩袋_系统尺寸;
                                    汇总表.Cells[currentRow, 5].Value = 总盒数;
                                    汇总表.Cells[currentRow, 6].Value = 获取仓位(包装数据.POF热缩袋);
                                    汇总表.Cells[currentRow, 11].Value = Path.GetFileName(当前处理文件名);
                                    currentRow++;
                                }
                                // ...全棉织带、通口箱同理...

                                // 3. 输出所有纸箱（按标签码1分组、再分配箱型）
                                foreach (var 标签分组 in 物料分组.GroupBy(x => x.标签码1))
                                {
                                    string 当前标签码1 = 标签分组.Key;
                                    int 盒子总数 = 标签分组.Count();
                                    int 最大支持盒数 = 0;
                                    var 产品型号信息 = 包装数据.支持产品型号列表.FirstOrDefault(x => x.产品型号 == 当前型号);
                                    if (产品型号信息 != null && !string.IsNullOrEmpty(产品型号信息.支持的盒装外箱))
                                    {
                                        if (产品型号信息.支持的盒装外箱.Contains("五盒装")) 最大支持盒数 = 5;
                                        else if (产品型号信息.支持的盒装外箱.Contains("三盒装")) 最大支持盒数 = 3;
                                        else if (产品型号信息.支持的盒装外箱.Contains("二盒装")) 最大支持盒数 = 2;
                                        else if (产品型号信息.支持的盒装外箱.Contains("单盒装")) 最大支持盒数 = 1;
                                    }
                                    int 剩余盒数 = 盒子总数;

                                    // 五盒装
                                    while (最大支持盒数 >= 5 && !string.IsNullOrEmpty(包装数据.五盒装外箱) && 剩余盒数 >= 5)
                                    {
                                        汇总表.Cells[currentRow, 2].Value = "纸箱";
                                        汇总表.Cells[currentRow, 3].Value = 包装数据.五盒装外箱;
                                        汇总表.Cells[currentRow, 4].Value = 包装数据.五盒装尺寸;
                                        汇总表.Cells[currentRow, 5].Value = 1;
                                        汇总表.Cells[currentRow, 6].Value = 获取仓位(包装数据.五盒装外箱);
                                        汇总表.Cells[currentRow, 9].Value = "5盒装标准";
                                        汇总表.Cells[currentRow, 11].Value = Path.GetFileName(当前处理文件名);
                                        汇总表.Cells[currentRow, 12].Value = 当前标签码1;
                                        currentRow++;
                                        剩余盒数 -= 5;
                                    }
                                    // 三盒装
                                    while (最大支持盒数 >= 3 && !string.IsNullOrEmpty(包装数据.三盒装外箱) && 剩余盒数 >= 3)
                                    {
                                        汇总表.Cells[currentRow, 2].Value = "纸箱";
                                        汇总表.Cells[currentRow, 3].Value = 包装数据.三盒装外箱;
                                        汇总表.Cells[currentRow, 4].Value = 包装数据.三盒装尺寸;
                                        汇总表.Cells[currentRow, 5].Value = 1;
                                        汇总表.Cells[currentRow, 6].Value = 获取仓位(包装数据.三盒装外箱);
                                        汇总表.Cells[currentRow, 9].Value = "3盒装标准";
                                        汇总表.Cells[currentRow, 11].Value = Path.GetFileName(当前处理文件名);
                                        汇总表.Cells[currentRow, 12].Value = 当前标签码1;
                                        currentRow++;
                                        剩余盒数 -= 3;
                                    }
                                    // 二盒装
                                    while (最大支持盒数 >= 2 && !string.IsNullOrEmpty(包装数据.二盒装外箱) && 剩余盒数 >= 2)
                                    {
                                        汇总表.Cells[currentRow, 2].Value = "纸箱";
                                        汇总表.Cells[currentRow, 3].Value = 包装数据.二盒装外箱;
                                        汇总表.Cells[currentRow, 4].Value = 包装数据.二盒装尺寸;
                                        汇总表.Cells[currentRow, 5].Value = 1;
                                        汇总表.Cells[currentRow, 6].Value = 获取仓位(包装数据.二盒装外箱);
                                        汇总表.Cells[currentRow, 9].Value = "2盒装标准";
                                        汇总表.Cells[currentRow, 11].Value = Path.GetFileName(当前处理文件名);
                                        汇总表.Cells[currentRow, 12].Value = 当前标签码1;
                                        currentRow++;
                                        剩余盒数 -= 2;
                                    }
                                    // 单盒装
                                    while (最大支持盒数 >= 1 && !string.IsNullOrEmpty(包装数据.单盒装外箱) && 剩余盒数 > 0)
                                    {
                                        汇总表.Cells[currentRow, 2].Value = "纸箱";
                                        汇总表.Cells[currentRow, 3].Value = 包装数据.单盒装外箱;
                                        汇总表.Cells[currentRow, 4].Value = 包装数据.单盒装尺寸;
                                        汇总表.Cells[currentRow, 5].Value = 1;
                                        汇总表.Cells[currentRow, 6].Value = 获取仓位(包装数据.单盒装外箱);
                                        汇总表.Cells[currentRow, 9].Value = "1盒装标准";
                                        汇总表.Cells[currentRow, 11].Value = Path.GetFileName(当前处理文件名);
                                        汇总表.Cells[currentRow, 12].Value = 当前标签码1;
                                        currentRow++;
                                        剩余盒数 -= 1;
                                    }
                                }
                            }
                        }
                        else
                        {
                            int 总盒数 = 使用数量;
                            // 获取支持的最大盒数
                            int 最大支持盒数 = 0;
                            // 从包装数据中查找对应的产品型号信息
                            var 产品型号信息 = 包装数据.支持产品型号列表.FirstOrDefault(x => x.产品型号 == 当前型号);
                            if (产品型号信息 != null && !string.IsNullOrEmpty(产品型号信息.支持的盒装外箱))
                            {
                                if (产品型号信息.支持的盒装外箱.Contains("五盒装"))  // 添加五盒装的判断
                                    最大支持盒数 = 5;
                                else if (产品型号信息.支持的盒装外箱.Contains("三盒装"))
                                    最大支持盒数 = 3;
                                else if (产品型号信息.支持的盒装外箱.Contains("二盒装"))
                                    最大支持盒数 = 2;
                                else if (产品型号信息.支持的盒装外箱.Contains("单盒装"))
                                    最大支持盒数 = 1;
                            }

                            int 剩余盒数 = 总盒数;  // 使用计算出的总盒数

                            // 优先使用单盒装，如果总盒数为1
                            if (总盒数 == 1 && 最大支持盒数 >= 1 && !string.IsNullOrEmpty(包装数据.单盒装外箱))
                            {
                                汇总表.Cells[currentRow, 1].Value = "";
                                汇总表.Cells[currentRow, 2].Value = "纸箱";
                                汇总表.Cells[currentRow, 3].Value = 包装数据.单盒装外箱;
                                汇总表.Cells[currentRow, 4].Value = 包装数据.单盒装尺寸;
                                汇总表.Cells[currentRow, 5].Value = 1;
                                汇总表.Cells[currentRow, 6].Value = 获取仓位(包装数据.单盒装外箱);
                                汇总表.Cells[currentRow, 9].Value = "1盒装标准";
                                汇总表.Cells[currentRow, 11].Value = Path.GetFileName(当前处理文件名);

                                var range = 汇总表.Cells[currentRow, 1, currentRow, 9];
                                range.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                                range.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                                range.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                                range.Style.Border.Right.Style = ExcelBorderStyle.Thin;

                                currentRow++;
                                剩余盒数 = 0;
                            }
                            // 如果总盒数大于1，则按照最优组合方式处理
                            else
                            {
                                // 处理五盒装
                                if (最大支持盒数 >= 5 && !string.IsNullOrEmpty(包装数据.五盒装外箱) && 剩余盒数 >= 5)
                                {
                                    int 五盒装数量 = 剩余盒数 / 5;

                                    汇总表.Cells[currentRow, 1].Value = "";
                                    汇总表.Cells[currentRow, 2].Value = "纸箱";
                                    汇总表.Cells[currentRow, 3].Value = 包装数据.五盒装外箱;
                                    汇总表.Cells[currentRow, 4].Value = 包装数据.五盒装尺寸;
                                    汇总表.Cells[currentRow, 5].Value = 五盒装数量;
                                    汇总表.Cells[currentRow, 6].Value = 获取仓位(包装数据.五盒装外箱);
                                    汇总表.Cells[currentRow, 9].Value = "5盒装标准";
                                    汇总表.Cells[currentRow, 11].Value = Path.GetFileName(当前处理文件名);

                                    var range = 汇总表.Cells[currentRow, 1, currentRow, 9];
                                    range.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                                    range.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                                    range.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                                    range.Style.Border.Right.Style = ExcelBorderStyle.Thin;

                                    currentRow++;
                                    剩余盒数 %= 5;
                                }

                                // 处理三盒装
                                if (最大支持盒数 >= 3 && !string.IsNullOrEmpty(包装数据.三盒装外箱) && 剩余盒数 >= 3)
                                {
                                    int 三盒装数量 = 剩余盒数 / 3;

                                    汇总表.Cells[currentRow, 1].Value = "";
                                    汇总表.Cells[currentRow, 2].Value = "纸箱";
                                    汇总表.Cells[currentRow, 3].Value = 包装数据.三盒装外箱;
                                    汇总表.Cells[currentRow, 4].Value = 包装数据.三盒装尺寸;
                                    汇总表.Cells[currentRow, 5].Value = 三盒装数量;
                                    汇总表.Cells[currentRow, 6].Value = 获取仓位(包装数据.三盒装外箱);
                                    汇总表.Cells[currentRow, 9].Value = "3盒装标准";
                                    汇总表.Cells[currentRow, 11].Value = Path.GetFileName(当前处理文件名);

                                    var range = 汇总表.Cells[currentRow, 1, currentRow, 9];
                                    range.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                                    range.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                                    range.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                                    range.Style.Border.Right.Style = ExcelBorderStyle.Thin;

                                    currentRow++;
                                    剩余盒数 %= 3;
                                }

                                // 处理二盒装
                                if (最大支持盒数 >= 2 && !string.IsNullOrEmpty(包装数据.二盒装外箱) && 剩余盒数 >= 2)
                                {
                                    int 二盒装数量 = 剩余盒数 / 2;

                                    汇总表.Cells[currentRow, 1].Value = "";
                                    汇总表.Cells[currentRow, 2].Value = "纸箱";
                                    汇总表.Cells[currentRow, 3].Value = 包装数据.二盒装外箱;
                                    汇总表.Cells[currentRow, 4].Value = 包装数据.二盒装尺寸;
                                    汇总表.Cells[currentRow, 5].Value = 二盒装数量;
                                    汇总表.Cells[currentRow, 6].Value = 获取仓位(包装数据.二盒装外箱);
                                    汇总表.Cells[currentRow, 9].Value = "2盒装标准";
                                    汇总表.Cells[currentRow, 11].Value = Path.GetFileName(当前处理文件名);

                                    var range = 汇总表.Cells[currentRow, 1, currentRow, 9];
                                    range.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                                    range.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                                    range.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                                    range.Style.Border.Right.Style = ExcelBorderStyle.Thin;

                                    currentRow++;
                                    剩余盒数 %= 2;
                                }

                                // 处理单盒装
                                if (最大支持盒数 >= 1 && !string.IsNullOrEmpty(包装数据.单盒装外箱) && 剩余盒数 > 0)
                                {
                                    汇总表.Cells[currentRow, 1].Value = "";
                                    汇总表.Cells[currentRow, 2].Value = "纸箱";
                                    汇总表.Cells[currentRow, 3].Value = 包装数据.单盒装外箱;
                                    汇总表.Cells[currentRow, 4].Value = 包装数据.单盒装尺寸;
                                    汇总表.Cells[currentRow, 5].Value = 剩余盒数;
                                    汇总表.Cells[currentRow, 6].Value = 获取仓位(包装数据.单盒装外箱);
                                    汇总表.Cells[currentRow, 9].Value = "1盒装标准";
                                    汇总表.Cells[currentRow, 11].Value = Path.GetFileName(当前处理文件名);

                                    var range = 汇总表.Cells[currentRow, 1, currentRow, 9];
                                    range.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                                    range.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                                    range.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                                    range.Style.Border.Right.Style = ExcelBorderStyle.Thin;

                                    currentRow++;
                                }
                            }
                        }
                    }
                }

                汇总包.Save();
            }

            uiTextBox_状态.AppendText($"包装汇总已更新到 {汇总文件路径}" + Environment.NewLine);
            uiTextBox_状态.AppendText("------------------------" + Environment.NewLine);
        }

        private void 添加配件说明书_5(string 文件夹路径)
        {
            string 汇总文件路径 = Path.Combine(文件夹路径, "包装材料需求流转单.xlsx");
            FileInfo 汇总文件信息 = new FileInfo(汇总文件路径);

            using (ExcelPackage 汇总包 = new ExcelPackage(汇总文件信息))
            {
                ExcelWorksheet 汇总表 = 汇总包.Workbook.Worksheets["包装材料需求流转单"];
                if (汇总表 == null)
                {
                    MessageBox.Show("找不到'包装材料需求源数据'工作表", "错误");
                    return;
                }

                int lastRow = 6; // 默认值
                if (汇总表.Dimension != null)
                {
                    lastRow = 汇总表.Dimension.End.Row;
                }
                int currentRow = lastRow + 1;

                // 统计纸箱和通口箱数量
                int 纸箱总数 = 0;
                int 通口箱数量 = 0;
                for (int row = 7; row <= lastRow; row++)
                {
                    string 产品类型 = 汇总表.Cells[row, 2].Text;
                    if (产品类型.Contains("通口箱"))
                    {
                        if (int.TryParse(汇总表.Cells[row, 5].Text, out int 数量))
                        {
                            通口箱数量 += 数量;
                        }
                    }
                    else if (产品类型.Contains("纸箱"))
                    {
                        if (int.TryParse(汇总表.Cells[row, 5].Text, out int 数量))
                        {
                            纸箱总数 += 数量;
                        }
                    }
                }

                // 输出调试信息
                //MessageBox.Show($"找到通口箱数量: {通口箱数量}, 纸箱总数: {纸箱总数}");

                // 确定说明书数量
                int 说明书数量 = 变量.客户代码.Contains("17034") ? 通口箱数量 : 纸箱总数;

                // 添加说明书信息
                List<int> 不用下操作指引的客户 = new List<int> { 12573, 12095, 12056, 16066, 12058, 12251, 17100, 12075, 13009, 12233, 18236, 12141, 13086, 12020, 14035, 14029 };
                if (!不用下操作指引的客户.Any(客户代码 => 变量.客户代码.Contains(客户代码.ToString())))
                {
                    bool 已添加说明书 = false;

                    // 根据包装要求添加主要说明书
                    if (变量.包装要求.Contains("水下说明书") || 变量.标签.Contains("水下"))
                    {
                        汇总表.Cells[currentRow, 1].Value = "";
                        汇总表.Cells[currentRow, 2].Value = "配件说明书";
                        汇总表.Cells[currentRow, 3].Value = "1006.010026";
                        汇总表.Cells[currentRow, 4].Value = "UL676-UL2108-CE 水下通用说明书";
                        汇总表.Cells[currentRow, 5].Value = 说明书数量;
                        汇总表.Cells[currentRow, 6].Value = 获取仓位("1006.010026");
                        汇总表.Cells[currentRow, 11].Value = Path.GetFileName(当前处理文件名);

                        var dataRange = 汇总表.Cells[currentRow, 1, currentRow, 9];
                        dataRange.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                        dataRange.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                        dataRange.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        dataRange.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

                        currentRow++;
                        已添加说明书 = true;
                    }
                    else if (变量.包装要求.Contains("中文"))
                    {
                        汇总表.Cells[currentRow, 1].Value = "";
                        汇总表.Cells[currentRow, 2].Value = "配件说明书";
                        汇总表.Cells[currentRow, 3].Value = "1006.010021";
                        汇总表.Cells[currentRow, 4].Value = "中文版中性通用操作指引说明书";
                        汇总表.Cells[currentRow, 5].Value = 说明书数量;
                        汇总表.Cells[currentRow, 6].Value = 获取仓位("1006.010021");
                        汇总表.Cells[currentRow, 11].Value = Path.GetFileName(当前处理文件名);

                        var dataRange = 汇总表.Cells[currentRow, 1, currentRow, 9];
                        dataRange.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                        dataRange.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                        dataRange.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        dataRange.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

                        currentRow++;
                        已添加说明书 = true;
                    }
                    else
                    {
                        汇总表.Cells[currentRow, 1].Value = "";
                        汇总表.Cells[currentRow, 2].Value = "配件说明书";
                        汇总表.Cells[currentRow, 3].Value = "1006.010020";
                        汇总表.Cells[currentRow, 4].Value = "英文版中性通用操作指引说明书";
                        汇总表.Cells[currentRow, 5].Value = 说明书数量;
                        汇总表.Cells[currentRow, 6].Value = 获取仓位("1006.010020");
                        汇总表.Cells[currentRow, 11].Value = Path.GetFileName(当前处理文件名);

                        var dataRange = 汇总表.Cells[currentRow, 1, currentRow, 9];
                        dataRange.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                        dataRange.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                        dataRange.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        dataRange.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

                        currentRow++;
                        已添加说明书 = true;
                    }

                    // 只在未添加过水下说明书时添加额外的说明书
                    if (!已添加说明书)
                    {
                        if (变量.标签.Contains("UL 676"))
                        {
                            汇总表.Cells[currentRow, 1].Value = "";
                            汇总表.Cells[currentRow, 2].Value = "配件说明书";
                            汇总表.Cells[currentRow, 3].Value = "1006.010022";
                            汇总表.Cells[currentRow, 4].Value = "英文版UL676产品系列通用";
                            汇总表.Cells[currentRow, 5].Value = 说明书数量;
                            汇总表.Cells[currentRow, 6].Value = 获取仓位("1006.010022");
                            汇总表.Cells[currentRow, 11].Value = Path.GetFileName(当前处理文件名);

                            var dataRange = 汇总表.Cells[currentRow, 1, currentRow, 9];
                            dataRange.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                            dataRange.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                            dataRange.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                            dataRange.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

                            currentRow++;
                        }

                        if (变量.标签.Contains("水下"))
                        {
                            汇总表.Cells[currentRow, 1].Value = "";
                            汇总表.Cells[currentRow, 2].Value = "配件说明书";
                            汇总表.Cells[currentRow, 3].Value = "1006.010026";
                            汇总表.Cells[currentRow, 4].Value = "UL676-UL2108-CE 水下通用说明书";
                            汇总表.Cells[currentRow, 5].Value = 说明书数量;
                            汇总表.Cells[currentRow, 6].Value = 获取仓位("1006.010023");
                            汇总表.Cells[currentRow, 11].Value = Path.GetFileName(当前处理文件名);

                            var dataRange = 汇总表.Cells[currentRow, 1, currentRow, 9];
                            dataRange.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                            dataRange.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                            dataRange.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                            dataRange.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

                            currentRow++;
                        }
                    }
                }

                汇总包.Save();
            }
        }

        private Dictionary<string, string> 加载仓位数据()
        {
            Dictionary<string, string> 仓位字典 = new Dictionary<string, string>();
            try
            {
                string 仓位表路径 = Path.Combine(Application.StartupPath, "新包装资料", "仓位.xlsx");
                if (!File.Exists(仓位表路径))
                {
                    MessageBox.Show($"仓位表文件不存在: {仓位表路径}", "错误");
                    return 仓位字典;
                }

                using (ExcelPackage package = new ExcelPackage(new FileInfo(仓位表路径)))
                {
                    var worksheet = package.Workbook.Worksheets[0];
                    int rowCount = worksheet.Dimension.End.Row;

                    for (int row = 2; row <= rowCount; row++)
                    {
                        string 物料码 = worksheet.Cells[row, 1].Text?.Trim();
                        string 仓位 = worksheet.Cells[row, 2].Text?.Trim();
                        if (!string.IsNullOrEmpty(物料码))
                        {
                            仓位字典[物料码] = 仓位;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"加载仓位数据出错: {ex.Message}", "错误");
            }
            return 仓位字典;
        }

        private string 获取仓位(string 物料码)
        {
            string 仓位文件路径 = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "新包装资料", "仓位.xlsx");

            if (File.Exists(仓位文件路径))
            {
                using (ExcelPackage 仓位包 = new ExcelPackage(new FileInfo(仓位文件路径)))
                {
                    ExcelWorksheet 仓位表 = 仓位包.Workbook.Worksheets[0];
                    int 行数 = 仓位表.Dimension?.End.Row ?? 0;

                    // 直接查找匹配的物料码
                    for (int row = 1; row <= 行数; row++)
                    {
                        string 当前物料码 = 仓位表.Cells[row, 1].Text?.Trim() ?? "";
                        if (当前物料码 == 物料码)
                        {
                            return 仓位表.Cells[row, 2].Text?.Trim() ?? "#N/A";
                        }
                    }
                }
            }

            return "#N/A";
        }
    }
}