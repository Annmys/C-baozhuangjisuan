using OfficeOpenXml; // EPPlus的命名空间.
using Sunny.UI;
using System.Text;
using OfficeOpenXml.Style;
using System.Drawing;
using System.Text.RegularExpressions;
using System.Net.Http;
using System.Text;
using Newtonsoft.Json;
using System.Threading.Tasks;
using System.Management;
using System.ComponentModel;
using Timer = System.Windows.Forms.Timer; 

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

            下载新包装资料();

            // 软件启动加载包装数据
            包装数据字典 = 加载包装数据();

            算法选择.SelectedIndexChanged += 算法选择_SelectedIndexChanged;

        }


        // 添加选择变更事件处理函数
        private void 算法选择_SelectedIndexChanged(object sender, EventArgs e)
        {
            当前选择算法 = 算法选择.Text;
        }

        /// <summary>
        /// 从网络路径下载文件到本地目录
        /// </summary>
        private void 下载新包装资料()
        {
            try
            {
                // 定义源路径和目标路径
                string 源路径 = @"\\192.168.1.33\Annmy\包装计算\新包装资料";
                string 目标路径 = Path.Combine(Application.StartupPath, "新包装资料");

                // 确保目标目录存在
                if (!Directory.Exists(目标路径))
                {
                    Directory.CreateDirectory(目标路径);
                }

                // 检查网络路径是否可访问
                if (!Directory.Exists(源路径))
                {
                    MessageBox.Show($"无法访问网络路径: {源路径}\n请检查网络连接或权限设置。", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                // 获取源目录中的所有文件
                string[] 文件列表 = Directory.GetFiles(源路径, "*.*", SearchOption.TopDirectoryOnly);
                int 总文件数 = 文件列表.Length;
                int 已复制文件数 = 0;
                int 跳过文件数 = 0;

                // 创建进度条窗口
                Form 进度窗口 = new Form();
                进度窗口.Text = "正在下载新包装资料";
                进度窗口.Size = new Size(400, 120);
                进度窗口.StartPosition = FormStartPosition.CenterScreen;
                进度窗口.FormBorderStyle = FormBorderStyle.FixedDialog;
                进度窗口.MaximizeBox = false;
                进度窗口.MinimizeBox = false;

                ProgressBar 进度条 = new ProgressBar();
                进度条.Minimum = 0;
                进度条.Maximum = 总文件数;
                进度条.Value = 0;
                进度条.Width = 360;
                进度条.Height = 20;
                进度条.Location = new Point(20, 20);

                Label 状态标签 = new Label();
                状态标签.AutoSize = true;
                状态标签.Location = new Point(20, 50);
                状态标签.Text = "准备下载...";

                进度窗口.Controls.Add(进度条);
                进度窗口.Controls.Add(状态标签);

                // 使用后台线程执行复制操作
                BackgroundWorker 后台工作线程 = new BackgroundWorker();
                后台工作线程.WorkerReportsProgress = true;
                后台工作线程.DoWork += (s, e) =>
                {
                    foreach (string 源文件 in 文件列表)
                    {
                        string 文件名 = Path.GetFileName(源文件);
                        string 目标文件 = Path.Combine(目标路径, 文件名);

                        // 检查目标文件是否已存在，如果存在则比较修改时间
                        bool 需要复制 = true;
                        if (File.Exists(目标文件))
                        {
                            DateTime 源文件时间 = File.GetLastWriteTime(源文件);
                            DateTime 目标文件时间 = File.GetLastWriteTime(目标文件);

                            // 如果目标文件比源文件新或相同，则跳过
                            if (目标文件时间 >= 源文件时间)
                            {
                                需要复制 = false;
                                跳过文件数++;
                            }
                        }

                        if (需要复制)
                        {
                            // 复制文件，覆盖已存在的文件
                            File.Copy(源文件, 目标文件, true);
                            已复制文件数++;
                        }

                        // 报告进度
                        后台工作线程.ReportProgress(已复制文件数 + 跳过文件数, 文件名);
                    }
                };

                后台工作线程.ProgressChanged += (s, e) =>
                {
                    进度条.Value = e.ProgressPercentage;
                    状态标签.Text = $"正在处理: {e.UserState}\n已复制: {已复制文件数}, 已跳过: {跳过文件数}, 总计: {总文件数}";
                };

                后台工作线程.RunWorkerCompleted += (s, e) =>
                {
                    if (e.Error != null)
                    {
                        MessageBox.Show($"下载过程中发生错误: {e.Error.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                    else
                    {
                        状态标签.Text = $"下载完成! 已复制: {已复制文件数}, 已跳过: {跳过文件数}, 总计: {总文件数}";
                        // 延迟关闭进度窗口

                        Timer 关闭计时器 = new Timer();
                        关闭计时器.Interval = 2000; // 2秒后关闭
                        关闭计时器.Tick += (sender, args) =>
                        {
                            关闭计时器.Stop();
                            进度窗口.Close();
                        };
                        关闭计时器.Start();
                    }
                };

                // 显示进度窗口并开始下载
                //进度窗口.Show();
                后台工作线程.RunWorkerAsync();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"下载新包装资料时发生错误: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private string 获取MAC地址()
        {
            try
            {
                ManagementObjectSearcher searcher = new ManagementObjectSearcher("SELECT MACAddress FROM Win32_NetworkAdapterConfiguration WHERE IPEnabled = 'TRUE'");
                ManagementObjectCollection information = searcher.Get();

                foreach (ManagementObject obj in information)
                {
                    if (obj["MACAddress"] != null)
                    {
                        return obj["MACAddress"].ToString().Trim();
                    }
                }

                return "未知";
            }
            catch (Exception ex)
            {
                Console.WriteLine($"获取MAC地址时出错: {ex.Message}");
                return "未知";
            }
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
                    int 剪切长度列 = -1;
                    int 客户代码列 = -1;
                    int 完工日期列 = -1;
                    int 业务员列 = -1;
                    int 标签列 = -1;
                    int 包装要求列 = -1;
                    int 客户型号列 = -1; // 新增客户型号列
                    int 标签要求列 = -1;

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

                            case "剪切长度":  // 新增剪切长度列的判断
                                剪切长度列 = col;
                                break;
                            case "客户代码":
                                客户代码列 = col;
                                break;
                            case "表头要求完工日期":
                                完工日期列 = col;
                                break;
                            case "客户业务员":
                                业务员列 = col;
                                break;
                            case "标签":
                                标签列 = col;
                                break;
                            case "包装要求":
                                包装要求列 = col;
                                break;
                                
                            case "客户型号1":  //可能被的表格名字又不叫这个了，测试如果改成客户型号想通用的话不行。
                                客户型号列 = col;
                                break;
                            case "标签要求":
                                标签要求列 = col;
                                break;
                        }
                    }

                    // 验证必要的列是否都找到
                    if (订单编号列 == -1 || 规格型号列 == -1 || 销售数量列 == -1)
                    {
                        throw new Exception("未找到必要的列标题（单据编号、规格型号或销售数量）");
                    }

                    变量.订单出线字典 = new Dictionary<string, List<(string 型号, HashSet<string> 出线方式, string F列内容, double 销售数量, string 客户型号)>>();
                    变量.订单编号 = worksheet.Cells[2, 订单编号列].Text;
                    变量.订单出线字典[变量.订单编号] = new List<(string, HashSet<string>, string, double, string)>();

                    变量.客户代码 = worksheet.Cells[2, 客户代码列].Text;
                    变量.完成时间 = worksheet.Cells[2, 完工日期列].Text;
                    变量.业务员 = worksheet.Cells[2, 业务员列].Text;
                    变量.标签 = worksheet.Cells[2, 标签列].Text;
                    变量.包装要求 = worksheet.Cells[2, 包装要求列].Text;
                    //MessageBox.Show(变量.标签);
                    //MessageBox.Show(变量.包装要求);

                    int startRow = -1;
                    string 当前型号 = "";
                    var 当前出线方式 = new HashSet<string>();
                    string 当前F列内容 = "";
                    double 当前销售数量 = 0;
                    string 当前客户型号 = ""; // 新增当前客户型号变量
                    //MessageBox.Show(客户型号列.ToString());

                    string 文件夹路径 = Path.Combine(Application.StartupPath, "输出结果", 变量.订单编号);
                    // 如果文件夹已存在，先删除
                    if (Directory.Exists(文件夹路径))
                    {
                        Directory.Delete(文件夹路径, true); // true表示递归删除所有内容
                    }

                    string 当前标签要求 = "";

                    // 从第2行开始处理数据
                    for (int row = 2; row <= worksheet.Dimension.End.Row; row++)
                    {
                        var materialCell = worksheet.Cells[row, F列];
                        var specCell = worksheet.Cells[row, 规格型号列];
                        var salesCell = worksheet.Cells[row, 销售数量列];

                        
                        if (标签要求列 > 0)
                        {
                            当前标签要求 = worksheet.Cells[2, 标签要求列].Text ?? "";
                            //MessageBox.Show(当前标签要求); 
                        }

                        if (materialCell.Value != null && materialCell.Text.StartsWith("80."))
                        {
                            // 如果已经有一个开始行，说明找到了下一个80.开头的行，需要处理之前的数据
                            if (startRow != -1)
                            {
                                
                                
                                
                                // 获取客户型号信息(如果存在)
                                当前客户型号 = "";
                                if (客户型号列 > 0)
                                {
                                    当前客户型号 = worksheet.Cells[startRow, 客户型号列].Text ?? "";
                                    //MessageBox.Show(当前客户型号);
                                }
                                // 处理之前的区间
                                获取订单数据(worksheet, startRow, row - 1, 规格型号列, F列, 销售数量列, 剪切长度列, ref 当前型号, ref 当前出线方式, ref 当前F列内容, ref 当前销售数量, ref 当前客户型号,ref 当前标签要求);

                                变量.订单出线字典[变量.订单编号].Add((当前型号, new HashSet<string>(当前出线方式), 当前F列内容, 当前销售数量, 当前客户型号));
                            }

                            // 更新新的开始行
                            startRow = row;
                            当前出线方式 = new HashSet<string>();
                            
                            // 获取当前行的客户型号(如果有)
                            当前客户型号 = "";
                            if (客户型号列 > 0)
                            {
                                当前客户型号 = worksheet.Cells[row, 客户型号列].Text ?? "";
                            }
                        }
                    }

                    // 处理最后一个区间
                    if (startRow != -1)
                    {
                        
                        
                        // 获取客户型号信息(如果存在)
                        当前客户型号 = "";
                        if (客户型号列 > 0)
                        {
                            当前客户型号 = worksheet.Cells[startRow, 客户型号列].Text ?? "";
                            
                            //MessageBox.Show(当前客户型号);
                        }

                        获取订单数据(worksheet, startRow, worksheet.Dimension.End.Row, 规格型号列, F列, 销售数量列, 剪切长度列, ref 当前型号, ref 当前出线方式, ref 当前F列内容, ref 当前销售数量, ref 当前客户型号,ref 当前标签要求);
                        
                        变量.订单出线字典[变量.订单编号].Add((当前型号, new HashSet<string>(当前出线方式), 当前F列内容, 当前销售数量, 当前客户型号));
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



        /// <summary>
        /// 获取备注列的索引，通过查询第一行，找到包含"备注"的列
        /// </summary>
        /// <param name="worksheet">Excel工作表</param>
        /// <returns>备注列的列索引，如果未找到，返回-1</returns>
        private int 获取备注列索引(ExcelWorksheet worksheet)
        {
            //MessageBox.Show($"worksheet{worksheet}");
            int maxColumn = worksheet.Dimension.End.Column;
            for (int col = 1; col <= maxColumn; col++)
            {
                var cellValue = worksheet.Cells[1, col].Text;
                if (cellValue.Contains("备注"))
                {
                    return col;
                }
            }
            return -1; // 未找到备注列
        }

        /// <summary>
        /// 从规格型号中提取线长列表，可能有一个或两个
        /// </summary>
        /// <param name="规格型号">规格型号字符串</param>
        /// <returns>线长数值列表</returns>
        private List<double> 提取线长列表(string 规格型号)
        {
            List<double> 线长列表 = new List<double>();
            // 使用正则表达式匹配 "-数字M"
            var matches = Regex.Matches(规格型号, @"-(\d+(\.\d+)?)M");
            foreach (Match match in matches)
            {
                if (match.Success && match.Groups.Count > 1)
                {
                    string 米数字符串 = match.Groups[1].Value; // 返回数字部分
                    if (double.TryParse(米数字符串, out double 米数))
                    {
                        线长列表.Add(米数);
                    }
                }
            }
            return 线长列表;
        }

        /// <summary>
        /// 从备注中获取替代的线长，如果有的话
        /// </summary>
        /// <param name="备注">备注字符串</param>
        /// <returns>线长数值，如果未找到返回0</returns>
        private double 从备注中获取线长(string 备注)
        {
            // 使用正则表达式匹配 "线长X米" 或 "线长X.XX米"
            var match = Regex.Match(备注, @"线长(\d+(\.\d+)?)米");
            if (match.Success && match.Groups.Count > 1)
            {
                string 线长字符串 = match.Groups[1].Value;
                if (double.TryParse(线长字符串, out double 线长))
                {
                    return 线长;
                }
            }
            return 0;
        }

        private void 输出订单信息(string 订单编号, List<(string 型号, HashSet<string> 出线方式, string 物料编码, double 销售数量, string 客户型号)> 型号出线方式列表)
        {

            uiTextBox_状态.AppendText($"订单编号: {订单编号}" + Environment.NewLine);
            foreach (var (型号, 出线方式, 物料编码, 销售数量, 客户型号) in 型号出线方式列表)
            {
                string 出线方式字符串 = 出线方式.Count > 0 ? string.Join("，", 出线方式) : "无";
                uiTextBox_状态.AppendText($"型号: {型号}" + Environment.NewLine);
                uiTextBox_状态.AppendText($"出线方式: {出线方式字符串}" + Environment.NewLine);
                uiTextBox_状态.AppendText($"物料编码: {物料编码}" + Environment.NewLine);
                uiTextBox_状态.AppendText($"销售数量: {销售数量}" + Environment.NewLine);
                if (!string.IsNullOrEmpty(客户型号))
                {
                    uiTextBox_状态.AppendText($"客户型号: {客户型号}" + Environment.NewLine);
                }
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
                foreach (var (型号, 出线方式, F列内容, 销售数量, 客户型号) in 订单明细列表)
                {
                    // 从匹配信息列表中查找对应型号的工作表名称
                    var 匹配信息 = 变量.订单附件匹配列表.FirstOrDefault(x =>
                        x.订单编号 == 订单编号 &&
                        x.产品型号 == 型号);

                    if (匹配信息 != null)
                    {
                        处理订单包装(型号, 出线方式.ToList(), F列内容, 销售数量, 匹配信息.工作表名称, 客户型号);
                    }
                    else
                    {
                        // 如果找不到匹配信息，可以使用一个默认值或者记录错误
                        uiTextBox_状态.Invoke((MethodInvoker)(() =>
                        {
                            uiTextBox_状态.AppendText($"警告：未找到型号 {型号} 的匹配工作表信息" + Environment.NewLine);
                        }));
                        处理订单包装(型号, 出线方式.ToList(), F列内容, 销售数量, "", 客户型号);
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

        private void 处理订单包装(string 型号, List<string> 出线方式列表, string F列内容, double 销售数量, string 工作表名称, string 客户型号)  // 添加工作表名称参数
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
                        if (匹配信息 != null)  // 添加检查，确保不是附件型号
                        {
                            // 检查是否是需要附件处理的型号
                            if (变量.需要附件处理的型号.ContainsKey(匹配信息.产品型号))
                            {
                                // 如果是见附件的情况，先不设置序号前缀，等待附件导入后再设置
                                return;
                            }
                            else if (变量.附件表数据.ContainsKey(工作表名称))
                            {
                                // 如果不是见附件，且附件表数据存在，则设置序号前缀
                                匹配信息.设置Sheet序号前缀(变量.附件表数据[工作表名称]);
                            }
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

                    // 完成序号前缀匹配
                    if (变量.附件表数据 != null)  // 先检查附件表数据字典是否存在
                    {
                        var 匹配信息 = 变量.订单附件匹配列表.FirstOrDefault(x => x.工作表名称 == 工作表名称);
                        if (匹配信息 != null && 变量.附件表数据.ContainsKey(工作表名称))
                        {
                            匹配信息.设置Sheet序号前缀(变量.附件表数据[工作表名称]);
                        }
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
        private async void button_附件导入_Click(object sender, EventArgs e)
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

                        string 工作表名称 = worksheet.Name;


                        //MessageBox.Show(工作表名称, "线长列号识别");

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
                        int 标签码3列号 = -1;
                        int 标签码4列号 = -1;

                        int 实际剪切长度毫米列号 = -1;
                        int 实际剪切长度米列号 = -1;
                        int 线长列号 = -1;  // 新增线长列号变量
                        int 一端线长列号 = -1;
                        int 二端线长列号 = -1;
                        int 总线长列号 = -1;


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
                                //MessageBox.Show(总长度列号.ToString());
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
                                //MessageBox.Show($"实际剪切长度毫米列号：\n列= {实际剪切长度毫米列号}\n");
                            }
                            if (cellValue.Contains("实际剪切长度(米)"))
                            {
                                实际剪切长度米列号 = col;
                            }
                            if (cellValue.Contains("线长"))
                            {
                                // 检查下一行是否包含米数
                                var nextRowValue = worksheet.Cells[序号行号 + 1, col].Text?.Trim() ?? "";
                                if (nextRowValue.EndsWith("m", StringComparison.OrdinalIgnoreCase))  // 检查是否以"m"结尾
                                {
                                    线长列号 = col;
                                    //MessageBox.Show($"找到线长列：\n列标题 = {cellValue}\n线长列号 = {线长列号}\n示例值 = {nextRowValue}", "线长列号识别");
                                    break;  // 找到符合条件的列后退出循环
                                }

                            }
                            if (Regex.IsMatch(cellValue, @"01.*端.*线.*长"))
                            {
                                一端线长列号 = col;
                                //MessageBox.Show($"正则匹配到一端线长列：\n列号 = {col}\n内容 = '{cellValue}'", "一端线长列号识别");
                            }

                            if (cellValue.Contains("02") && cellValue.Contains("端") && cellValue.Contains("线") && cellValue.Contains("长"))
                            {
                                二端线长列号 = col;
                                //MessageBox.Show($"找到二端线长列：\n列号 = {col}\n内容 = '{cellValue}'", "二端线长列号识别");
                            }
                            if (cellValue.Contains("总线长"))
                            {
                                总线长列号 = col;
                                //MessageBox.Show($"总线长列号：\n列标题 = {cellValue}\n总线长列号 = {总线长列号}", "总线长列号识别");
                            }

                            //MessageBox.Show($"一端线长列号：\n列标题 = {cellValue}\n一端线长列号 = {一端线长列号}", "一端线长列号识别");

                        }


                        // 处理数据（从标题行的下一行开始，确保不超出工作表范围）
                        for (int row = 序号行号 + 1; row <= worksheet.Dimension.End.Row; row++)
                        {
                            try
                            {
                                var cell序号 = worksheet.Cells[row, 序号列号];
                                var 空白位 = worksheet.Cells[row, 3];
                                var cell条数 = worksheet.Cells[row, 条数列号];
                                var cell总米数 = worksheet.Cells[row, 总长度列号];

                                //有问题
                                //836订单可以看到文本，889订单看不到文本
                                StringBuilder debug = new StringBuilder();
                                debug.AppendLine($"行号: {row}, 列号: {总长度列号}");
                                debug.AppendLine($"Text: {cell总米数.Text}");
                                debug.AppendLine($"Value: {cell总米数.Value}");
                                debug.AppendLine($"Formula: {cell总米数.Formula}");
                                debug.AppendLine($"单元格地址: {cell总米数.Address}");
                                //MessageBox.Show(debug.ToString());

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

                                //// 然后在数据读取时使用
                                //string 线长值 = 线长列号 != -1 ? worksheet.Cells[row, 线长列号].Text?.Trim() ?? "" : "";
                                //// 提取数字部分（去掉"m"）
                                //string 线长数字 = 线长值.ToLower().Replace("m", "").Trim();
                                //double 线长 = 1.0; // 默认值
                                //double.TryParse(线长数字, out 线长);
                                ////MessageBox.Show(线长.ToString(), "线长1");
                                ////如果没有线长，则寻找总线长
                                //if(线长 == 0) 
                                //{
                                //    线长值 = 总线长列号 != -1 ? worksheet.Cells[row, 总线长列号].Text?.Trim() ?? "" : "";
                                //    MessageBox.Show(线长值, "线长2");
                                //}

                                // 获取线长值
                                string 线长值 = 线长列号 != -1 ? worksheet.Cells[row, 线长列号].Text?.Trim() ?? "" : "";
                                double 线长 = 0.0;
                                // 尝试从02端线长获取
                                if (!string.IsNullOrEmpty(线长值))
                                {
                                    string 线长数字 = 线长值.ToLower().Replace("m", "").Trim();
                                    double.TryParse(线长数字, out 线长);
                                }

                                string 一端线长值 = 一端线长列号 != -1 ? worksheet.Cells[row, 一端线长列号].Text?.Trim() ?? "" : "";
                                double 一端线长 = 0.0;
                                if (!string.IsNullOrEmpty(一端线长值))
                                {
                                    string 一端线长数字 = 一端线长值.ToLower().Replace("m", "").Trim();
                                    double.TryParse(一端线长数字, out 一端线长);
                                }

                                string 二端线长值 = 二端线长列号 != -1 ? worksheet.Cells[row, 二端线长列号].Text?.Trim() ?? "" : "";
                                double 二端线长 = 0.0;
                                if (!string.IsNullOrEmpty(二端线长值))
                                {
                                    string 二端线长数字 = 二端线长值.ToLower().Replace("m", "").Trim();
                                    double.TryParse(二端线长数字, out 二端线长);
                                }

                                double 合并线长 = 一端线长 + 二端线长;

                                //MessageBox.Show($"行 {row} 的线长: {一端线长}", "线长信息");

                                // 如果02端线长为0，尝试从总线长获取
                                if (线长 == 0 && 总线长列号 != -1)
                                {
                                    int 总线长实际列号 = 总线长列号;
                                    ExcelRange 当前单元格 = worksheet.Cells[row, 总线长实际列号];

                                    // 检查当前单元格是否在合并区域内
                                    bool 是合并单元格 = false;
                                    string 合并区域 = "";

                                    // 正确判断是否为合并单元格
                                    var mergedAddress = worksheet.MergedCells?.FirstOrDefault(x =>
                                    {
                                        var addr = new ExcelAddress(x);
                                        return row >= addr.Start.Row &&
                                               row <= addr.End.Row &&
                                               总线长实际列号 >= addr.Start.Column &&
                                               总线长实际列号 <= addr.End.Column;
                                    });

                                    if (mergedAddress != null)
                                    {
                                        是合并单元格 = true;
                                        合并区域 = mergedAddress;
                                    }

                                    // 调试信息
                                    //MessageBox.Show(
                                    //    $"单元格检查:\n" +
                                    //    $"当前行号: {row}\n" +
                                    //    $"当前列号: {总线长实际列号}\n" +
                                    //    $"当前单元格地址: {当前单元格.Address}\n" +
                                    //    $"是否合并单元格: {是合并单元格}\n" +
                                    //    $"合并区域: {合并区域}",
                                    //    "单元格状态"
                                    //);
                                }

                                //MessageBox.Show($"行 {row} 的线长: {线长}", "线长信息");

                                // 如果两种方式都没有获取到线长，使用订单导入时的区间总线长
                                if (线长 == 0)
                                {
                                    // 获取匹配信息中的区间总线长
                                    var 匹配信息 = 变量.订单附件匹配列表.FirstOrDefault(x => x.工作表名称 == worksheet.Name);
                                    if (匹配信息 != null && 匹配信息.线材长度 > 0)
                                    {
                                        线长 = 匹配信息.线材长度;
                                    }
                                    else
                                    {
                                        线长 = 1.0; // 如果实在获取不到，才使用默认值
                                    }
                                }

                                string 实际值 = cell总米数.Text?.Trim() ?? "0";
                                

                                //MessageBox.Show($"总长度 {cell总米数.Value.ToString()}", "");
                                if (空白位.Value != null && cell条数.Value != null && cell总米数.Value != null)
                                {
                                    if (空白位.Value is ExcelErrorValue || cell条数.Value is ExcelErrorValue || cell总米数.Value is ExcelErrorValue)
                                    {
                                        continue;
                                    }

                                    string 序号 = cell序号.Value?.ToString() ?? "";
                                    //string 序号 = cell序号.Value.ToString();
                                    int 条数;
                                    double 总米数;

                                    if (int.TryParse(cell条数.Value.ToString(), out 条数) &&
                                        double.TryParse(cell总米数.Text?.Trim() ?? "0", out 总米数))
                                    {
                                        //if (double.TryParse(cell总米数.Text?.Trim() ?? "0", out double 解析米数))
                                        //{
                                        //    总米数 = 解析米数;
                                        //}

                                        if (条数 > 1)
                                        {
                                            double 分割后的米数 = 总米数 / 条数;
                                            for (int i = 0; i < 条数; i++)
                                            {
                                                string 实际剪切长度毫米 = 实际剪切长度毫米列号 != -1 ? worksheet.Cells[row, 实际剪切长度毫米列号].Text?.Trim() ?? "" : "";
                                                string 实际剪切长度米 = 实际剪切长度米列号 != -1 ? worksheet.Cells[row, 实际剪切长度米列号].Text?.Trim() ?? "" : "";
                                                string 添加的数据 = $"{序号}, {1}, {分割后的米数}, {标签码}, {标签码1}, {标签码2}, {实际剪切长度毫米}, {实际剪切长度米}, {线长}";
                                                变量.附件表数据[worksheet.Name].Add(添加的数据);
                                            }
                                        }
                                        else
                                        {
                                            string 实际剪切长度毫米 = 实际剪切长度毫米列号 != -1 ? worksheet.Cells[row, 实际剪切长度毫米列号].Text?.Trim() ?? "" : "";
                                            string 实际剪切长度米 = 实际剪切长度米列号 != -1 ? worksheet.Cells[row, 实际剪切长度米列号].Text?.Trim() ?? "" : "";
                                            string 添加的数据 = $"{序号}, {条数}, {总米数}, {标签码}, {标签码1}, {标签码2}, {实际剪切长度毫米}, {实际剪切长度米}, {线长}";
                                            变量.附件表数据[worksheet.Name].Add(添加的数据);
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
                                // 检查序号不是空的，且不包含 "Grand Total"
                                if (数据项.Length >= 3 &&
                                    !string.IsNullOrEmpty(数据项[0].Trim()) &&
                                    !数据项[0].Contains("Grand Total") &&
                                    double.TryParse(数据项[2].Trim(), out double 总米数))
                                {
                                    //MessageBox.Show($"长度 {数据项[2].Trim()}", "");
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
                            //if (!string.IsNullOrEmpty(uiTextBox_订单地址.Text) && 变量.订单出线字典 != null && 变量.订单出线字典.Any())
                            //{
                            //    foreach (var 订单 in 变量.订单出线字典)
                            //    {
                            //        foreach (var (型号, 出线方式, F列内容, 销售数量) in 订单.Value)
                            //        {
                            //            // 将销售数量四舍五入到三位小数进行比较
                            //            double 订单数量 = Math.Round(销售数量, 3);
                            //            if (Math.Abs(订单数量 - 总米数和) < 0.001) // 使用小于0.001的差值来判断相等
                            //            {
                            //                // 查找所有匹配的订单附件信息
                            //                var 匹配信息列表 = 变量.订单附件匹配列表
                            //                    .Where(x => x.产品型号 == 型号 && Math.Abs(x.销售数量 - 订单数量) < 0.001)
                            //                    .ToList();

                            //                // 根据匹配数量进行不同处理
                            //                if (匹配信息列表.Count > 0)
                            //                {
                            //                    // 如果有多个匹配，需要分别处理
                            //                    if (匹配信息列表.Count > 1)
                            //                    {
                            //                        // 获取所有匹配的工作表名称
                            //                        var 工作表名称列表 = 匹配信息列表.Select(x => x.工作表名称).Distinct().ToList();

                            //                        // 确保有足够的文件来处理多个匹配
                            //                        string 文件夹路径 = Path.Combine(Application.StartupPath, "输出结果", 变量.订单编号);
                            //                        string 处理好订单文件夹路径 = Path.Combine(文件夹路径, "订单资料");

                            //                        // 查找所有匹配的文件（包括带数字后缀的）
                            //                        var 匹配文件列表 = Directory.GetFiles(处理好订单文件夹路径, $"{型号}-{销售数量}-附件*.xlsx")
                            //                            .OrderBy(f => f)
                            //                            .ToList();

                            //                        if (匹配文件列表.Count >= 工作表名称列表.Count)
                            //                        {
                            //                            // 为每个匹配的工作表更新对应的文件
                            //                            for (int i = 0; i < 工作表名称列表.Count; i++)
                            //                            {
                            //                                if (i < 匹配文件列表.Count)
                            //                                {
                            //                                    MessageBox.Show("多个型号的同样销售数量！！！");
                            //                                }
                            //                            }
                            //                        }
                            //                        else
                            //                        {
                            //                            MessageBox.Show($"警告：文件数量不足，需要{工作表名称列表.Count}个文件，但只找到{匹配文件列表.Count}个文件", "文件数量不足");
                            //                        }
                            //                    }
                            //                    else
                            //                    {
                            //                        // 只有一个匹配，正常处理
                            //                        var 匹配信息 = 匹配信息列表.First();
                            //                        匹配信息.工作表名称 = worksheet.Name;  // 更新为实际的工作表名称
                            //                        更新订单_附件型号Excel文件添加附件数据(型号, 销售数量, worksheet.Name);
                            //                    }
                            //                }
                            //            }
                            //        }
                            //    }
                            //}

                            // 如果订单已导入，进行数量比对
                            if (!string.IsNullOrEmpty(uiTextBox_订单地址.Text) && 变量.订单出线字典 != null && 变量.订单出线字典.Any())
                            {
                                
                                foreach (var 订单 in 变量.订单出线字典)
                                {
                                    // 创建一个集合来跟踪已经使用过的匹配信息
                                    HashSet<匹配信息> 已使用的匹配信息 = new HashSet<匹配信息>();

                                    foreach (var (型号, 出线方式, F列内容, 销售数量, 客户型号) in 订单.Value)
                                    {
                                        // 将销售数量四舍五入到三位小数进行比较
                                        double 订单数量 = Math.Round(销售数量, 2);
                                        //MessageBox.Show(订单数量.ToString(), 总米数和.ToString());

                                        if (Math.Abs(订单数量 - 总米数和) < 0.01) // 使用小于0.01的差值来判断相等
                                        {
                                            // 查找所有相同型号的匹配信息，排除已使用的
                                            var 匹配信息列表 = 变量.订单附件匹配列表
                                                .Where(x => x.产品型号 == 型号 && !已使用的匹配信息.Contains(x))
                                                .ToList();

                                            if (匹配信息列表.Any())
                                            {
                                                // 设置可接受的最大差异（例如5%）
                                                double 可接受差异比例 = 0.05;
                                                double 可接受差异值 = 订单数量 * 可接受差异比例;

                                                // 查找最接近的匹配信息
                                                var 最接近的匹配信息 = 匹配信息列表
                                                    .OrderBy(x => Math.Abs(x.销售数量 - 订单数量))
                                                    .FirstOrDefault();

                                                // 检查是否找到匹配信息
                                                if (最接近的匹配信息 != null)
                                                {
                                                    // 计算差异
                                                    double 差异 = Math.Abs(最接近的匹配信息.销售数量 - 订单数量);

                                                    // 检查差异是否在可接受范围内
                                                    if (差异 <= 可接受差异值)
                                                    {
                                                        // 标记该匹配信息为已使用
                                                        已使用的匹配信息.Add(最接近的匹配信息);

                                                        最接近的匹配信息.工作表名称 = worksheet.Name;  // 更新为实际的工作表名称

                                                        // 更新已存在的Excel文件，添加附件数据
                                                        更新订单_附件型号Excel文件添加附件数据(型号, 最接近的匹配信息.销售数量, worksheet.Name);

                                                        // 可选：显示匹配信息
                                                        //MessageBox.Show($"找到最接近的匹配：\n型号：{型号}\n订单数量：{订单数量}\n匹配数量：{最接近的匹配信息.销售数量}\n差异：{差异}\n可接受差异：{可接受差异值}");
                                                    }
                                                    else
                                                    {
                                                        // 差异超出可接受范围
                                                        MessageBox.Show($"找到的匹配差异过大：\n型号：{型号}\n订单数量：{订单数量}\n最接近匹配数量：{最接近的匹配信息.销售数量}\n差异：{差异}\n可接受差异：{可接受差异值}");
                                                    }
                                                }
                                                else
                                                {
                                                    // 没有找到任何匹配
                                                    MessageBox.Show($"未找到型号 {型号} 的任何匹配信息");
                                                }
                                            }
                                            else
                                            {
                                                // 没有找到任何匹配
                                                MessageBox.Show($"未找到型号 {型号} 的任何未使用的匹配信息");
                                            }
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

            //if (uiCheckBox_RU客户.Checked)
            //{
            //    MessageBox.Show("RU客户");
            //}
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
                List<新包装资料> 候选包装列表 = new List<新包装资料>();

                候选包装列表 = 包装查询.获取候选包装列表(产品型号, 包装类型);


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

                    string 使用的表名;
                    // 检查是否有组合结果，确保至少有一个工作表
                    var 当前工作表匹配 = 变量.订单附件匹配列表.FirstOrDefault(x =>
                        x.产品型号 == 产品型号 &&
                        Math.Abs(x.销售数量 - 匹配信息.销售数量) < 0.001);

                    if (当前工作表匹配 != null)
                    {
                        使用的表名 = 当前工作表匹配.工作表名称;

                        // 添加调试信息
                        //MessageBox.Show($"匹配信息:\n" +
                        //               $"产品型号: {产品型号}\n" +
                        //               $"销售数量: {当前工作表匹配.销售数量}\n" +
                        //               $"工作表名称: {使用的表名}",
                        //               "工作表名称调试",
                        //               MessageBoxButtons.OK,
                        //               MessageBoxIcon.Information);
                    }
                    else
                    {
                        使用的表名 = 产品型号;
                        MessageBox.Show($"未找到匹配信息，使用产品型号作为表名: {使用的表名}",
                                       "工作表名称调试",
                                       MessageBoxButtons.OK,
                                       MessageBoxIcon.Warning);
                    }

                    // 验证使用的表名是否存在于附件表数据中
                    if (!变量.附件表数据.ContainsKey(使用的表名))
                    {
                        MessageBox.Show($"错误：找不到1表名 {使用的表名} 的数据！\n" +
                                       $"可用的表名：{string.Join(", ", 变量.附件表数据.Keys)}",
                                       "错误",
                                       MessageBoxButtons.OK,
                                       MessageBoxIcon.Error);
                        return;
                    }

                    int 数据条数 = 变量.附件表数据[使用的表名].Count - 1;

                    //MessageBox.Show(数据条数.ToString(), "数据条数", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    // 选择最佳包装   测试V0.9
                    //var 包装资料 = 包装查询.选择最佳包装(候选包装列表, 数据条数);

                    string 灯带型号 = 获取灯带型号(); // 从订单数据中获取灯带型号
                    string 出线方式 = 获取出线方式(); // 从订单数据中获取出线方式


                    var 包装资料 = 包装查询.选择最佳包装(候选包装列表, 灯带型号, 出线方式, 数据条数);

                    // 方案2：启动新的任务进行AI分析
                    //if (包装资料 != null)
                    //{
                    //    //MessageBox.Show($"开始AI分析", "AI分析", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    //    Task.Run(async () =>
                    //    {
                    //        try
                    //        {
                    //            var AI建议 = await 获取AI包装建议(灯带型号, 出线方式, 数据条数, 包装资料);
                    //            // 使用 Invoke 在UI线程上显示消息框
                    //            this.Invoke((MethodInvoker)delegate
                    //            {
                    //                MessageBox.Show($"AI包装方案分析：\n\n{AI建议}", "AI分析结果", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    //            });
                    //        }
                    //        catch (Exception ex)
                    //        {
                    //            this.Invoke((MethodInvoker)delegate
                    //            {
                    //                MessageBox.Show($"AI分析发生错误：{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    //            });
                    //        }
                    //    });
                    //}

                    //MessageBox.Show(包装资料.半成品BOM物料码, "数据条数", MessageBoxButtons.OK, MessageBoxIcon.Information);

                    // 保存完整的包装资料到匹配信息中
                    匹配信息.选中包装资料 = 包装资料;  // 这一行是关键

                    if (包装资料 != null)
                    {
                        // 计算实际可用面积
                        double 实际面积 = (包装资料.总有效面积 - 包装资料.内撑纸卡面积) * 0.7;

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

        private string 获取灯带型号()
        {
            // 从订单数据中获取灯带型号
            if (变量.订单出线字典 != null && 变量.订单出线字典.ContainsKey(变量.订单编号))
            {
                var 订单信息 = 变量.订单出线字典[变量.订单编号];
                if (订单信息.Any())
                {
                    return 订单信息.First().型号; // 返回第一个订单的灯带型号
                }
            }
            return string.Empty; // 如果未找到，返回空字符串
        }

        private string 获取出线方式()
        {
            // 从订单数据中获取出线方式
            if (变量.订单出线字典 != null && 变量.订单出线字典.ContainsKey(变量.订单编号))
            {
                var 订单信息 = 变量.订单出线字典[变量.订单编号];
                if (订单信息.Any())
                {
                    return 订单信息.First().出线方式.FirstOrDefault() ?? string.Empty; // 返回第一个订单的第一个出线方式
                }
            }
            return string.Empty; // 如果未找到，返回空字符串
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
            List<double> 灯带长度列表 = new List<double>();
            List<double> 线长列表 = new List<double>();
            List<数据项> 数据项列表 = new List<数据项>();
            //MessageBox.Show($"产品型号: {产品型号}", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);

            // 获取匹配信息中的包装资料
            var 匹配信息 = 变量.订单附件匹配列表.FirstOrDefault(x => x.工作表名称 == 型号SHEET);
            if (匹配信息 == null || 匹配信息.选中包装资料 == null)
            {
                MessageBox.Show($"未找到工作表 {型号SHEET} 的包装资料", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // 确保包装资料类型正确
            var 包装资料 = 匹配信息.选中包装资料 as 新包装资料;
            if (包装资料 == null)
            {
                MessageBox.Show($"包装资料类型转换失败", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }



            if (变量.附件表数据.ContainsKey(型号SHEET))
            {
                var 有效数据 = 变量.附件表数据[型号SHEET].Take(变量.附件表数据[型号SHEET].Count - 1);

                foreach (var item in 有效数据)
                {
                    string[] 数据 = item.Split(',');

                    StringBuilder sb = new StringBuilder();
                    for (int i = 0; i < 数据.Length; i++)
                    {
                        sb.AppendLine($"数据[{i}] = {数据[i].Trim()}");
                    }

                    //MessageBox.Show(sb.ToString(), "数据内容检查");

                    if (double.TryParse(数据[2].Trim(), out double 灯带长度))
                    {
                        string 内容A = 数据[0].Trim();
                        string 内容O = 数据[1].Trim();

                        // 先检查数据项长度是否足够
                        if (数据.Length >= 9 && double.TryParse(数据[8].Trim(), out double 线材长度1))
                        {
                            // 有附件且能获取到线材长度
                            double 灯带面积 = 灯带尺寸.每厘米面积 * (灯带长度 * 100);
                            灯带长度列表.Add(灯带长度);
                            线长列表.Add(线材长度1);
                            //下面的标志很重要，会影响EXCEL保存报错
                            string 标志 = Guid.NewGuid().ToString();
                            数据项列表.Add(new 数据项(内容A, 内容O, 灯带面积, 线材长度1, 标志));
                        }
                        else
                        {
                            // 无附件或获取不到线材长度，使用匹配信息中的默认线材长度
                            double 单条线长 = 匹配信息.线材长度;
                            double 灯带面积 = 灯带尺寸.每厘米面积 * (灯带长度 * 100);
                            灯带长度列表.Add(灯带长度);
                            线长列表.Add(单条线长);
                            //下面的标志很重要，会影响EXCEL保存报错
                            string 标志 = Guid.NewGuid().ToString();
                            数据项列表.Add(new 数据项(内容A, 内容O, 灯带面积, 单条线长, 标志));
                        }

                    }

                }

                // 在调用Calculate带线长Combinations之前
                //MessageBox.Show($"输入数据检查:\n" +
                //    $"灯带长度列表数量: {灯带长度列表.Count}\n" +
                //    $"线长列表数量: {线长列表.Count}\n" +
                //    $"灯带长度: {string.Join(", ", 灯带长度列表)}\n" +
                //    $"线长: {string.Join(", ", 线长列表)}",
                //    "输入数据");

                Solution带线长 s = new Solution带线长();

                if (变量.是否RU客户)
                {
                    var 组合结果 = s.CalculateRU客户(
                    灯带长度列表,
                    线长列表,
                    产品型号,
                    变量.查找组合_基数,
                    包装资料,
                    变量.灯带尺寸列表  // 传入灯带尺寸列表
                    );
                    // 保存结果
                    保存组合结果到Excel(组合结果, 数据项列表, 订单编号, 型号SHEET, 灯带尺寸);
                }
                else
                {
                    var 组合结果1 = s.Calculate带线长Combinations(
                    灯带长度列表,
                    线长列表,
                    产品型号,
                    变量.查找组合_基数,
                    包装资料,
                    变量.灯带尺寸列表  // 传入灯带尺寸列表
                    );
                    // 保存结果
                    保存组合结果到Excel(组合结果1, 数据项列表, 订单编号, 型号SHEET, 灯带尺寸);
                }






            }
        }





        private void 保存组合结果到Excel(List<List<double>> 组合结果, List<数据项> 数据项原列表, string 订单编号, string 工作表名称, 灯带尺寸 灯带尺寸对象)
        {
            // 在方法开始处添加
            Dictionary<string, int> 序号使用计数 = new Dictionary<string, int>();

            string 文件夹路径 = Path.Combine(Application.StartupPath, "输出结果", 订单编号);

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
                // 检查是否有组合结果，确保至少有一个工作表
                if (组合结果 == null || 组合结果.Count == 0)
                {
                    // 检查是否存在同名工作表，如果存在则删除
                    var existingWorksheet = excelPackage.Workbook.Worksheets["无组合结果"];
                    if (existingWorksheet != null)
                    {
                        excelPackage.Workbook.Worksheets.Delete("无组合结果");
                    }

                    // 创建新的工作表
                    var worksheet = excelPackage.Workbook.Worksheets.Add("无组合结果");
                    worksheet.Cells[1, 1].Value = "未找到有效的组合方案";
                }


                for (int i = 0; i < 组合结果.Count; i++)
                {
                    //12.24增加修改部分,
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
                    //MessageBox.Show($"工作表名: {工作表名}", "", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    // 如果已存在同名工作表，先删除它
                    var 已存在工作表 = excelPackage.Workbook.Worksheets.FirstOrDefault(ws => ws.Name == 工作表名);
                    if (已存在工作表 != null)
                    {
                        excelPackage.Workbook.Worksheets.Delete(已存在工作表);
                    }

                    ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add(工作表名);


                    //12.24原来可以用的部分
                    //var combination = 组合结果[i];
                    //List<数据项> 可用数据项 = 数据项原列表
                    //    .Where(x => !全局已使用序号.Contains(x.内容A))
                    //    .ToList();

                    //// 如果可用数据项不足，跳过此组合
                    //if (可用数据项.Count < combination.Count)
                    //{
                    //    continue;
                    //}

                    //string 工作表名 = $"第 {i + 1}盒";
                    //if (excelPackage.Workbook.Worksheets.Any(ws => ws.Name == 工作表名))
                    //{
                    //    工作表名 = $"第 {i + 1}盒_{Guid.NewGuid().ToString().Substring(0, 4)}";
                    //}

                    //ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add(工作表名);

                    worksheet.Cells[1, 1].Value = "序号";
                    worksheet.Cells[1, 2].Value = "条数";
                    worksheet.Cells[1, 3].Value = "米数";
                    worksheet.Cells[1, 8].Value = "线长"; // 新增 "线长" 列
                    worksheet.Cells[1, 9].Value = "包装编码";

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
                        var 项 = 可用数据项
                        .Where(d => Math.Abs(d.内容R - combination[j]) < 0.001)  // 首先匹配面积
                        .Where(d => !序号使用计数.ContainsKey(d.内容A) ||
                        序号使用计数[d.内容A] < 数据项原列表.Count(x => x.内容A == d.内容A))  // 检查使用次数
                        .OrderBy(d => d.内容A.Length)  // 先按序号长度排序（A1在A10之前）
                        .ThenBy(d => d.内容A)  // 再按序号字符串排序
                        .FirstOrDefault();

                        if (项 != null)
                        {
                            // 更新使用计数
                            if (!序号使用计数.ContainsKey(项.内容A))
                            {
                                序号使用计数[项.内容A] = 1;
                            }
                            else
                            {
                                序号使用计数[项.内容A]++;
                            }

                            worksheet.Cells[j + 2, 1].Value = 项.内容A;
                            worksheet.Cells[j + 2, 2].Value = 项.内容O;

                            // 使用传入的灯带尺寸参数
                            double 米数 = 面积值 / (灯带尺寸对象.每厘米面积 * 100);
                            worksheet.Cells[j + 2, 3].Value = Math.Round(米数, 3);

                            double 线长 = 项.线长; // 直接获取传入的 "线长"
                            worksheet.Cells[j + 2, 8].Value = Math.Round(线长, 3); // 填充 "线长"

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


                    // 在这里添加填充最后一盒
                    if (i == 组合结果.Count - 1)
                    {
                        // 获取工作表的行数
                        int endRow = worksheet.Dimension.End.Row;

                        // 填充最后一盒的数据行（从第2行到endRow，第1行为标题行）
                        using (var range = worksheet.Cells[2, 1, endRow, 7])
                        {
                            range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                            range.Style.Fill.BackgroundColor.SetColor(Color.Red);
                        }
                    }



                }

                //2025.1.22问题点：无附件订单保存正常，有附件订单保存报错。已经解决
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

            // 在开始处添加调试
            // MessageBox.Show("开始获取BOM物料码");

            // 获取BOM物料码 - 直接使用之前保存的包装信息
            var 匹配信息 = 变量.订单附件匹配列表.FirstOrDefault(x =>
                x.工作表名称 == 工作表名称);

            // 添加调试信息
            if (匹配信息 == null)
            {
                MessageBox.Show("未找到匹配信息");
                return;
            }

            if (匹配信息.选中包装资料 == null)
            {
                MessageBox.Show("未找到选中包装资料");
                return;
            }

            新包装资料? 包装资料 = (新包装资料)匹配信息.选中包装资料;
            string? BOM物料码 = 包装资料?.半成品BOM物料码;

            if (string.IsNullOrEmpty(BOM物料码))
            {
                MessageBox.Show("BOM物料码为空");
                return;
            }

            //MessageBox.Show($"找到BOM物料码: {BOM物料码}");

            //新包装资料? 包装资料 = null;
            //string? BOM物料码 = null;

            if (匹配信息?.选中包装资料 != null)
            {
                包装资料 = (新包装资料)匹配信息.选中包装资料;
                BOM物料码 = 包装资料.半成品BOM物料码;

                uiTextBox_状态.AppendText($"使用已选择的包装 - BOM物料码: {BOM物料码}" + Environment.NewLine);

                // 添加更多包装信息的输出，方便调试
                uiTextBox_状态.AppendText($"包装名称: {包装资料.包装名称}" + Environment.NewLine);
                uiTextBox_状态.AppendText($"总有效面积: {包装资料.总有效面积}" + Environment.NewLine);
                uiTextBox_状态.AppendText("------------------------" + Environment.NewLine);
            }
            else
            {
                uiTextBox_状态.AppendText($"警告：未找到之前选择的包装信息 - 工作表: {工作表名称}" + Environment.NewLine);
                return; // 如果没有包装信息，直接返回
            }


            // 创建或更新汇总Excel
            string 汇总文件路径 = Path.Combine(Application.StartupPath, "输出结果", 订单编号, "包装材料需求流转单.xlsx");
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


                    // 删除匹配的配件说明书行
                    for (int row = 7; row <= lastRow; row++)
                    {
                        string 物料类型 = 汇总表.Cells[row, 2].Text;

                        if (物料类型 == "配件说明书")
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
                    汇总表.Cells["C3:D3"].Merge = true;
                    汇总表.Cells["C3"].Value = "客户代码:" + 变量.客户代码;
                    汇总表.Cells["E3:I3"].Merge = true;
                    string 转换后时间 = 变量.完成时间.Replace("-", ".");
                    汇总表.Cells["E3"].Value = "完成时间:" + 变量.完成时间;
                    // D3-I3保持空白

                    // 第4行设置
                    汇总表.Cells["A4:B4"].Merge = true;
                    string 制单日期 = DateTime.Now.ToString("yyyy.MM.dd");
                    汇总表.Cells["A4"].Value = "制单日期:" + 制单日期;

                    //连接名:     以太网
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
                    汇总表.Column(1).Width = 6.86; // 根据需要调整数值
                    汇总表.Column(2).Width = 16.29;
                    汇总表.Column(3).Width = 20.43;
                    汇总表.Column(4).Width = 29;
                    汇总表.Column(5).Width = 6.29;
                    汇总表.Column(6).Width = 8.43;
                    汇总表.Column(7).Width = 8.43;
                    汇总表.Column(8).Width = 14.57;
                    汇总表.Column(9).Width = 10.57;
                    //for (int col = 1; col <= 9; col++)
                    //{
                    //    汇总表.Column(col).AutoFit();
                    //}

                    // 设置标题字体
                    汇总表.Cells["A2"].Style.Font.Size = 14;
                    汇总表.Cells["A2"].Style.Font.Bold = true;

                    // 设置特定单元格的字体加粗
                    var boldCells = new[] { "A3", "B3", "C3", "E3", "E4", "A4", "H4", "B4", "C4", "F4" };
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

                // 用于跟踪纸箱总数量
                int 纸箱总数量 = 0;

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
                        汇总表.Cells[currentRow, 1].Value = 序号前缀;
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
                        半成品BOM.Style.Font.Bold = true;
                        汇总表.Row(currentRow).Height = 24.75;

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
                            pof范围.Style.Font.Bold = true;
                            汇总表.Row(currentRow).Height = 24.75;


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
                            织带范围.Style.Font.Bold = true;
                            汇总表.Row(currentRow).Height = 24.75;


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

                            // 累加纸箱数量
                            纸箱总数量 += 纸箱信息.Item2;

                            // 设置纸箱行的样式
                            var range = 汇总表.Cells[currentRow, 1, currentRow, 9];
                            range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                            range.Style.Fill.BackgroundColor.SetColor(Color.LightGreen);
                            range.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                            range.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                            range.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                            range.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                            range.Style.Font.Bold = true;
                            汇总表.Row(currentRow).Height = 24.75;


                            currentRow++;
                        }


                        // 处理配件说明书信息
                        List<int> 不用下操作指引的客户 = new List<int> { 12573, 12095, 12056, 16066, 12058, 12251, 17100, 12075, 13009, 12233, 18236, 12141, 13086, 12020, 14035, 14029 };

                        // 检查客户代码是否在目标列表中
                        if (!不用下操作指引的客户.Any(客户代码 => 变量.客户代码.Contains(客户代码.ToString())))
                        {
                            if (变量.包装要求.Contains("中文"))
                            {
                                汇总表.Cells[currentRow, 2].Value = "配件说明书";
                                汇总表.Cells[currentRow, 3].Value = "1006.010021";
                                汇总表.Cells[currentRow, 4].Value = "中文版中性通用操作指引说明书";
                                汇总表.Cells[currentRow, 5].Value = 纸箱总数量;  // 使用纸箱总数量

                                // 设置配件说明书行的样式
                                var 说明书范围 = 汇总表.Cells[currentRow, 1, currentRow, 9];
                                说明书范围.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                                说明书范围.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                                说明书范围.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                                说明书范围.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                                说明书范围.Style.Font.Bold = true;
                                汇总表.Row(currentRow).Height = 24.75;


                                currentRow++;
                            }
                            else
                            {
                                汇总表.Cells[currentRow, 2].Value = "配件说明书";
                                汇总表.Cells[currentRow, 3].Value = "1006.010020";
                                汇总表.Cells[currentRow, 4].Value = "英文版中性通用操作指引说明书";
                                汇总表.Cells[currentRow, 5].Value = 纸箱总数量;  // 使用纸箱总数量

                                // 设置配件说明书行的样式
                                var 说明书范围 = 汇总表.Cells[currentRow, 1, currentRow, 9];
                                说明书范围.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                                说明书范围.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                                说明书范围.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                                说明书范围.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                                说明书范围.Style.Font.Bold = true;
                                汇总表.Row(currentRow).Height = 24.75;


                                currentRow++;
                            }

                            if (变量.标签.Contains("UL 676"))
                            {
                                汇总表.Cells[currentRow, 2].Value = "配件说明书";
                                汇总表.Cells[currentRow, 3].Value = "1006.010022";
                                汇总表.Cells[currentRow, 4].Value = "英文版UL676产品系列通用";
                                汇总表.Cells[currentRow, 5].Value = 纸箱总数量;  // 使用纸箱总数量

                                // 设置配件说明书行的样式
                                var 说明书范围 = 汇总表.Cells[currentRow, 1, currentRow, 9];
                                说明书范围.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                                说明书范围.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                                说明书范围.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                                说明书范围.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                                说明书范围.Style.Font.Bold = true;
                                汇总表.Row(currentRow).Height = 24.75;


                                currentRow++;
                            }

                            if (变量.标签.Contains("水下"))
                            {
                                汇总表.Cells[currentRow, 2].Value = "配件说明书";
                                汇总表.Cells[currentRow, 3].Value = "1006.010023";
                                汇总表.Cells[currentRow, 4].Value = "英文版UL2108/CE产品系列水下方案说明书通用";
                                汇总表.Cells[currentRow, 5].Value = 纸箱总数量;  // 使用纸箱总数量

                                // 设置配件说明书行的样式
                                var 说明书范围 = 汇总表.Cells[currentRow, 1, currentRow, 9];
                                说明书范围.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                                说明书范围.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                                说明书范围.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                                说明书范围.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                                说明书范围.Style.Font.Bold = true;
                                汇总表.Row(currentRow).Height = 24.75;


                                currentRow++;
                            }
                        }


                    }



                    // 自动调整列宽
                    //for (int col = 1; col <= 9; col++)
                    //{
                    //    汇总表.Column(col).AutoFit();
                    //}



                }

                // 在保存汇总包之前，查询并更新仓位信息
                string 仓位文件路径 = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "新包装资料", "仓位.xlsx");

                // 检查仓位文件是否存在
                if (File.Exists(仓位文件路径))
                {
                    // 创建一个字典来存储物料编码和对应的仓位
                    Dictionary<string, string> 仓位字典 = new Dictionary<string, string>();

                    // 读取仓位文件
                    using (ExcelPackage 仓位包 = new ExcelPackage(new FileInfo(仓位文件路径)))
                    {
                        ExcelWorksheet 仓位表 = 仓位包.Workbook.Worksheets[0]; // 假设仓位信息在第一个工作表

                        // 获取已使用的行数
                        int 行数 = 仓位表.Dimension?.End.Row ?? 0;

                        // 读取所有物料编码和对应的仓位
                        for (int row = 1; row <= 行数; row++)
                        {
                            string 物料编码 = 仓位表.Cells[row, 1].Text?.Trim() ?? "";
                            string 仓位 = 仓位表.Cells[row, 2].Text?.Trim() ?? "";

                            if (!string.IsNullOrEmpty(物料编码))
                            {
                                仓位字典[物料编码] = 仓位;
                            }
                        }
                    }

                    // 遍历汇总表中的所有行，查找并更新仓位信息
                    ExcelWorksheet 汇总工作表 = 汇总包.Workbook.Worksheets["包装材料需求流转单"];
                    if (汇总工作表 != null)
                    {
                        int 最大行 = 汇总工作表.Dimension?.End.Row ?? 0;

                        // 从第7行开始查找物料编码（根据图片，第7行开始是数据行）
                        for (int row = 7; row <= 最大行; row++)
                        {
                            // 获取物料编码（C列）
                            string 物料编码 = 汇总工作表.Cells[row, 3].Text?.Trim() ?? "";

                            if (!string.IsNullOrEmpty(物料编码))
                            {
                                // 查找对应的仓位
                                string 仓位 = "#N/A";
                                if (仓位字典.ContainsKey(物料编码))
                                {
                                    仓位 = 仓位字典[物料编码];
                                }

                                // 更新仓位信息（G列）
                                汇总工作表.Cells[row, 6].Value = 仓位;
                            }
                        }
                    }
                }

                汇总包.Save();
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

            if (变量.是否RU客户)
            {
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
            else
            {
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

        private async void uiButton2_Click(object sender, EventArgs e)
        {
            try
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
                    List<新包装资料> 候选包装列表 = 包装查询.获取候选包装列表(产品型号, 包装类型);

                    if (候选包装列表.Any())
                    {
                        string 使用的表名;
                        var 当前工作表匹配 = 变量.订单附件匹配列表.FirstOrDefault(x =>
                            x.产品型号 == 产品型号 &&
                            Math.Abs(x.销售数量 - 匹配信息.销售数量) < 0.001);

                        if (当前工作表匹配 != null)
                        {
                            使用的表名 = 当前工作表匹配.工作表名称;
                        }
                        else
                        {
                            使用的表名 = 产品型号;
                            MessageBox.Show($"未找到匹配信息，使用产品型号作为表名: {使用的表名}",
                                           "工作表名称调试",
                                           MessageBoxButtons.OK,
                                           MessageBoxIcon.Warning);
                        }

                        if (!变量.附件表数据.ContainsKey(使用的表名))
                        {
                            MessageBox.Show($"错误：找不到表名 {使用的表名} 的数据！\n" +
                                           $"可用的表名：{string.Join(", ", 变量.附件表数据.Keys)}",
                                           "错误",
                                           MessageBoxButtons.OK,
                                           MessageBoxIcon.Error);
                            continue;
                        }

                        int 数据条数 = 变量.附件表数据[使用的表名].Count - 1;

                        // 对每个候选包装进行AI分析
                        List<(新包装资料 包装, double 评分)> 包装评分列表 = new List<(新包装资料, double)>();

                        foreach (var 包装 in 候选包装列表)
                        {
                            var AI建议 = await 获取AI包装建议(产品型号, 包装类型, 数据条数, 包装);
                            double 评分 = 解析AI评分(AI建议);
                            包装评分列表.Add((包装, 评分));
                        }

                        // 选择评分最高的包装
                        var 最佳包装 = 包装评分列表.OrderByDescending(x => x.评分).First();
                        var 包装资料 = 最佳包装.包装;

                        // 保存完整的包装资料到匹配信息中
                        匹配信息.选中包装资料 = 包装资料;

                        if (包装资料 != null)
                        {
                            // 计算实际可用面积
                            double 实际面积 = (包装资料.总有效面积 - 包装资料.内撑纸卡面积) * 0.7;
                            变量.查找组合_基数 = 实际面积;

                            uiTextBox_状态.AppendText($"型号 {产品型号} 使用包装: {包装资料.包装名称}" + Environment.NewLine);
                            uiTextBox_状态.AppendText($"总面积: {包装资料.总有效面积}" + Environment.NewLine);
                            uiTextBox_状态.AppendText($"实际可用面积: {实际面积}" + Environment.NewLine);
                            uiTextBox_状态.AppendText($"灯带数量: {数据条数}条" + Environment.NewLine);

                            string 包装类型说明 = 包装资料.是多条短条专用包装 ? "多条专用包装" :
                                              包装资料.允许多条包装 ? "多条包装" : "普通包装";
                            uiTextBox_状态.AppendText($"包装类型: {包装类型说明}" + Environment.NewLine);
                            uiTextBox_状态.AppendText("------------------------" + Environment.NewLine);
                        }
                        else
                        {
                            MessageBox.Show($"未找到型号 {产品型号} 出线方式 {包装类型} 的匹配包装", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            变量.查找组合_基数 = 1420 - 706;
                            uiTextBox_状态.AppendText($"型号 {产品型号} 使用默认包装容积: {变量.查找组合_基数}" + Environment.NewLine);
                            uiTextBox_状态.AppendText("------------------------" + Environment.NewLine);
                        }

                        // 调用开始组合方法
                        开始组合(型号SHEET, 产品型号, 匹配信息.订单编号);
                    }
                    else
                    {
                        MessageBox.Show($"未找到型号 {产品型号} 出线方式 {包装类型} 的候选包装", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        变量.查找组合_基数 = 1420 - 706;
                        uiTextBox_状态.AppendText($"型号 {产品型号} 使用默认包装容积: {变量.查找组合_基数}" + Environment.NewLine);
                        uiTextBox_状态.AppendText("------------------------" + Environment.NewLine);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"处理过程中出错：{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }


        private double 解析AI评分(string ai建议)
        {
            try
            {
                // 查找"适用度评分："后面的数字
                var 评分行 = ai建议.Split('\n')
                    .FirstOrDefault(line => line.Trim().StartsWith("适用度评分："));

                if (评分行 != null)
                {
                    var 评分文本 = 评分行.Split('：')[1].Trim();
                    if (double.TryParse(评分文本, out double 评分))
                    {
                        return 评分;
                    }
                }

                // 如果没有找到评分，返回默认值
                return 50; // 默认中等评分
            }
            catch
            {
                return 50; // 出错时返回默认中等评分
            }
        }



        private async Task<string> 获取AI包装建议(string 灯带型号, string 出线方式, int 数据条数, 新包装资料 选中包装)
        {
            const string apiKey = "sk-9S1Znc8NehWPQflzDqESpNL1xYykJNVcFCAko0RoeOuBhgzu";
            const string apiUrl = "https://www.dmxapi.com/v1/chat/completions";  // 移除末尾的斜杠

            try
            {
                // 构造包装信息描述
                var 包装描述 = $"包装名称: {选中包装.包装名称}\n" +
                            $"适用产品: {string.Join(", ", 选中包装.装箱产品型号)}\n" +
                            $"包装尺寸: {选中包装.总有效面积}*{选中包装.高度}\n" +
                            $"单盒装选用: {选中包装.单盒装选用}\n" +
                            $"二盒装选用: {选中包装.二盒装选用}\n" +
                            $"三盒装选用: {选中包装.三盒装选用}";

                // 构造请求数据
                var requestData = new
                {
                    model = "claude-3-5-sonnet-20240620",
                    messages = new[]
                    {
                new
                {
                    role = "user",
                    content = $@"请分析以下包装方案是否适合这个产品：
                    产品信息：
                    - 灯带型号：{灯带型号}
                    - 出线方式：{出线方式}
                    - 数据条数：{数据条数}

                    选择的包装方案：
                    {包装描述}

                     请从以下几个方面分析并给出评分：
                    1. 产品型号匹配度（0-30分）
                    2. 包装空间利用率（0-40分）
                    3. 操作便利性（0-30分）

                    请按以下格式返回：
                    适用度评分：XX
                    详细分析：
                    1. 产品型号匹配度：XX分
                    原因：...
                    2. 包装空间利用率：XX分
                    原因：...
                    3. 操作便利性：XX分
                    原因：...
                    总体建议：..."
                }
            },
                    temperature = 0.7,
                    max_tokens = 1000
                };

                using (var client = new HttpClient())
                {
                    // 修改认证头的设置
                    client.DefaultRequestHeaders.Clear();
                    client.DefaultRequestHeaders.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", apiKey);
                    client.DefaultRequestHeaders.Accept.Add(new System.Net.Http.Headers.MediaTypeWithQualityHeaderValue("application/json"));

                    var jsonContent = System.Text.Json.JsonSerializer.Serialize(requestData);
                    var httpContent = new StringContent(jsonContent, Encoding.UTF8, "application/json");

                    var response = await client.PostAsync(apiUrl, httpContent);

                    if (response.IsSuccessStatusCode)
                    {
                        var responseJson = await response.Content.ReadAsStringAsync();
                        var responseObj = System.Text.Json.JsonDocument.Parse(responseJson);

                        // 修改响应解析逻辑
                        if (responseObj.RootElement.TryGetProperty("choices", out var choices) &&
                            choices.GetArrayLength() > 0)
                        {
                            return choices[0].GetProperty("message")
                                           .GetProperty("content")
                                           .GetString();
                        }
                    }
                    else
                    {
                        var errorResponse = await response.Content.ReadAsStringAsync();
                        return $"API调用失败: {response.StatusCode}\n{errorResponse}";
                    }
                }
                return "无法获取AI分析结果";
            }
            catch (Exception ex)
            {
                return $"AI分析出错: {ex.Message}";
            }
        }

        private void uiCheckBox_RU客户_CheckedChanged(object sender, EventArgs e)
        {
            变量.是否RU客户 = uiCheckBox_RU客户.Checked;
        }

        private void uiTextBox_订单地址_TextChanged(object sender, EventArgs e)
        {

        }

        private void uiButton3_Click_1(object sender, EventArgs e)
        {


            混合处理();

            try
            {
                string 文件夹路径 = Path.Combine(Application.StartupPath, "输出结果", 变量.订单编号);
                string 处理好订单文件夹路径 = Path.Combine(文件夹路径, "订单资料");

                // 获取订单资料文件夹中的所有Excel文件
                var Excel文件列表 = Directory.GetFiles(处理好订单文件夹路径, "*.xlsx");

                foreach (var 文件路径 in Excel文件列表)
                {

                    using (ExcelPackage package = new ExcelPackage(new FileInfo(文件路径)))
                    {
                        var worksheet = package.Workbook.Worksheets[0];

                        // 读取基本信息
                        string 型号 = worksheet.Cells[1, 2].Text;
                        string 出线方式 = worksheet.Cells[5, 2].Text;

                        // 创建长度列表
                        var 长度列表 = new List<(string 序号, double 灯带长度, double 线材长度, string 标签码1, string 标签码2, string 标签码3, string 标签码4,string 客户型号,string 标签显示长度)>();

                        // 从第8行开始读取数据
                        int currentRow = 8;
                        while (!string.IsNullOrEmpty(worksheet.Cells[currentRow, 2].Text))
                        {
                            int 数量;
                            if (int.TryParse(worksheet.Cells[currentRow, 2].Text, out 数量))
                            {
                                string 序号 = worksheet.Cells[currentRow, 1].Text?.Trim(); // 从第1列读取序号
                                double 灯带长度;
                                double 线材长度;

                                if (double.TryParse(worksheet.Cells[currentRow, 3].Text, out 灯带长度) &&
                                    double.TryParse(worksheet.Cells[currentRow, 4].Text, out 线材长度))
                                {
                                    // 根据数量添加多条记录
                                    for (int i = 0; i < 数量; i++)
                                    {
                                        string 标签码1 = worksheet.Cells[currentRow, 5].Text;
                                        string 标签码2 = worksheet.Cells[currentRow, 6].Text;
                                        string 标签码3 = worksheet.Cells[currentRow, 7].Text;
                                        string 标签码4 = worksheet.Cells[currentRow, 8].Text;
                                        string 客户型号 = worksheet.Cells[6, 2].Text?.Trim() ?? "";
                                        string 标签显示长度 = currentRow != -1 ? worksheet.Cells[currentRow, 9].Text?.Trim() ?? "0" : "0"; 
                                        长度列表.Add((序号, 灯带长度, 线材长度, 标签码1, 标签码2, 标签码3, 标签码4,客户型号,标签显示长度));
                                    }
                                }
                            }

                            currentRow++;
                        }

                        if (长度列表.Count > 0)
                        {
                            // 构建提示信息
                            StringBuilder sb = new StringBuilder();
                            sb.AppendLine($"文件名: {Path.GetFileName(文件路径)}");
                            sb.AppendLine($"型号: {型号}");
                            sb.AppendLine($"出线方式: {出线方式}");
                            sb.AppendLine($"总条数: {长度列表.Count}");
                            sb.AppendLine("\n长度明细:");

                            // 按长度分组并统计数量
                            var 分组统计 = 长度列表
                                .GroupBy(x => (x.灯带长度, x.线材长度))
                                .Select(g => (g.Key.灯带长度, g.Key.线材长度, 数量: g.Count()))
                                .OrderByDescending(x => x.灯带长度);

                            foreach (var 组 in 分组统计)
                            {
                                sb.AppendLine($"灯带长度: {组.灯带长度:F2}m, 线材长度: {组.线材长度:F2}m, 数量: {组.数量}条");
                            }

                            // 显示提示信息
                            //MessageBox.Show(sb.ToString(), "读取到的数据");
                            当前处理文件名 = 文件路径;
                            保存日志(sb.ToString(), "读取到的数据", Path.GetFileName(当前处理文件名));


                            // 调用查找包装方法
                            查找合适包装_0(长度列表, 型号, 出线方式);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"处理Excel文件出错: {ex.Message}\n\n{ex.StackTrace}", "错误");
            }



            string 文件夹路径1 = Path.Combine(Application.StartupPath, "输出结果", 变量.订单编号);
            添加配件说明书_5(文件夹路径1);

        }




        private void 保存日志(string 内容, string 日志类型, string excel文件名)
        {
            try
            {
                // 创建日志文件路径
                string 日志文件名 = Path.GetFileNameWithoutExtension(excel文件名) + $"_{日志类型}.txt";
                string 日志文件路径 = Path.Combine(
                    Application.StartupPath,
                    "输出结果",
                    变量.订单编号,
                    "计算日志",
                    日志文件名
                );

                // 确保日志文件夹存在
                Directory.CreateDirectory(Path.GetDirectoryName(日志文件路径));

                // 保存日志文件
                File.WriteAllText(日志文件路径, 内容, Encoding.UTF8);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"保存日志出错: {ex.Message}", "错误");
            }
        }

        private void uiButton_整合多个附件_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "Excel Files|*.xlsx";
                openFileDialog.Multiselect = true;

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    string[] selectedFiles = openFileDialog.FileNames;
                    string outputDirectory = Path.GetDirectoryName(selectedFiles[0]);
                    string outputFilePath = Path.Combine(outputDirectory, "合并结果.xlsx");

                    using (ExcelPackage package = new ExcelPackage())
                    {
                        foreach (string filePath in selectedFiles)
                        {
                            using (ExcelPackage inputPackage = new ExcelPackage(new FileInfo(filePath)))
                            {
                                foreach (var worksheet in inputPackage.Workbook.Worksheets)
                                {
                                    // 确保工作表名称唯一
                                    string sheetName = worksheet.Name;
                                    int suffix = 1;
                                    while (package.Workbook.Worksheets[sheetName] != null)
                                    {
                                        sheetName = $"{worksheet.Name}_{suffix++}";
                                    }

                                    // 创建新的工作表
                                    var newWorksheet = package.Workbook.Worksheets.Add(sheetName);

                                    // 复制内容
                                    for (int row = 1; row <= worksheet.Dimension.End.Row; row++)
                                    {
                                        for (int col = 1; col <= worksheet.Dimension.End.Column; col++)
                                        {
                                            newWorksheet.Cells[row, col].Value = worksheet.Cells[row, col].Value;
                                        }
                                    }
                                }
                            }
                        }

                        // 保存合并后的文件
                        package.SaveAs(new FileInfo(outputFilePath));
                    }

                    MessageBox.Show($"合并完成，文件已保存到: {outputFilePath}", "成功");
                }
            }
        }

    }


}