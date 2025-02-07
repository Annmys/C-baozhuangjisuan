using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace 包装计算
{






    //按照顺序版算法（简单）
    public class Solution
    {
        public List<List<double>> CalculateCombinations(List<double> numbers, double target)
        {
            List<List<double>> ans = new List<List<double>> { new List<double>() };

            double currentSum = 0;

            for (int i = 0; i < numbers.Count; i++)
            {
                if (currentSum + numbers[i] >= target)
                {
                    currentSum = 0;
                    ans.Add(new List<double>());
                }

                currentSum += numbers[i];
                ans.Last().Add(numbers[i]);
            }

            return ans;
        }
    }



    //节约包装版本算法（复杂）
    public class Solution1
    {
        public List<List<double>> CalculateCombinations(List<double> numbers, double target)
        {
            List<List<double>> tmp = new List<List<double>> { new List<double>() }; //存储所有可能的组合
            List<Tuple<int, double>> sum = new List<Tuple<int, double>> { new Tuple<int, double>(0, 0) }; // 存储每个组合的索引和当前组合的和
            List<int> numberLoc = new List<int>(new int[numbers.Count]); // 记录数字在哪个组合

            numbers.Sort(); // 对数字排序
            double currentSum = 0;
            int combinationIndex = 0;

            for (int i = 0; i < numbers.Count; i++)
            {
                if (currentSum + numbers[i] >= target)
                {
                    combinationIndex++;
                    currentSum = 0;
                    sum.Add(new Tuple<int, double>(combinationIndex, 0));
                    tmp.Add(new List<double>());
                }

                currentSum += numbers[i];
                tmp[combinationIndex].Add(numbers[i]);
                sum[combinationIndex] = new Tuple<int, double>(combinationIndex, currentSum);
                numberLoc[i] = combinationIndex;
            }

            // 找到组合和最小的组合
            var minP = sum.MinBy(s => s.Item2);

            int k = minP.Item1;
            double v = minP.Item2;
            int idx = -1;

            for (int i = 0; i < numbers.Count; i++)
            {
                if (v + numbers[i] > target) break;
                if (numberLoc[i] != k && sum[numberLoc[i]].Item2 - numbers[i] >= v + numbers[i])
                {
                    idx = i;
                }
            }

            if (idx != -1)
            {
                tmp[k].Add(numbers[idx]);
                tmp[numberLoc[idx]].Remove(numbers[idx]);
            }

            return tmp;
        }
    }



    // 在 suanfa.cs 中添加新的类
    public class Solution带线长
    {
        public class 组合结果项
        {
            public string 序号 { get; set; }
            public string 条数 { get; set; }
            public double 灯带长度 { get; set; }
            public double 线长 { get; set; }

            public 组合结果项(string 序号, string 条数, double 灯带长度, double 线长)
            {
                this.序号 = 序号;
                this.条数 = 条数;
                this.灯带长度 = 灯带长度;
                this.线长 = 线长;
            }
        }

        private const double 默认线径 = 6.5; // 默认线材直径 6.5mm
        private const double F23线径 = 5.4;  // F23型号线材直径 5.4mm

        public List<List<double>> Calculate带线长Combinations(
        List<double> 灯带长度列表,
        List<double> 线长列表,
        string 产品型号,
        double 查找组合_基数,
        新包装资料 包装规格,
        List<灯带尺寸> 灯带尺寸列表)
        {
           
            var 灯带尺寸 = 灯带尺寸列表.FirstOrDefault(x => x.型号 == 产品型号.Replace("B", ""));
            if (灯带尺寸 == null)
            {
                throw new Exception($"未找到型号 {产品型号} 的尺寸数据");
            }

            List<List<double>> ans = new List<List<double>> { new List<double>() };
            double currentSum = 0;
            double current线材总长 = 0;
            double current灯带米数 = 0;

            // 获取包装限制信息
            新包装 包装信息 = new 新包装();
            var 盘灯信息 = 包装信息.获取盘灯信息(包装规格.包装名称, 产品型号);
            if (盘灯信息 == null)
            {
                throw new Exception($"未找到包装 {包装规格.包装名称} 对应型号 {产品型号} 的限制信息");
            }

            double 最大灯带米数 = double.Parse(盘灯信息.盘灯米数);
            double 最大线材长度 = double.Parse(盘灯信息.盘线材);

            //显示获取到的限制信息
            //MessageBox.Show(
            //    $"包装名称: {包装规格.包装名称}\n" +
            //    $"产品型号: {产品型号}\n" +
            //    $"最大灯带米数: {最大灯带米数} 米\n" +
            //    $"最大线材长度: {最大线材长度} 米",
            //    "包装限制信息"
            //);


            List<List<double>> ans1 = new List<List<double>> { new List<double>() };  // 初始化第一个组合
            double current灯带米数1 = 0;
            double current线材长度1 = 0;

            //MessageBox.Show($"开始组合计算:\n" +
            //    $"灯带数量: {灯带长度列表.Count}\n" +
            //    $"最大灯带米数: {最大灯带米数}\n" +
            //    $"灯带列表: {string.Join(", ", 灯带长度列表)}",
            //    "初始数据");

            //2025.01.15可以用
            //for (int i = 0; i < 灯带长度列表.Count; i++)
            //{
            //    double 当前灯带米数 = 灯带长度列表[i];
            //    double 当前线材长度 = 线长列表[i];

            //    // 检查灯带米数和线材长度是否超出限制
            //    bool 超出灯带限制 = current灯带米数1 + 当前灯带米数 > 最大灯带米数;
            //    bool 超出线材限制 = Math.Max(current线材长度1, 当前线材长度) > 最大线材长度;

            //    if (超出灯带限制 || 超出线材限制)
            //    {
            //        current灯带米数1 = 0;  // 重置灯带累计值
            //        current线材长度1 = 0;  // 重置线材累计值
            //        ans1.Add(new List<double>());  // 创建新组合
            //    }

            //    // 添加到当前组合并更新累计值
            //    current灯带米数1 += 当前灯带米数;
            //    current线材长度1 = Math.Max(current线材长度1, 当前线材长度);  // 更新最大线材长度
            //    //ans1.Last().Add(当前灯带米数);
            //    double 当前灯带面积 = 当前灯带米数 * 灯带尺寸.每厘米面积 * 100;  // 转换为面积
            //    ans1.Last().Add(当前灯带面积);

            //    //MessageBox.Show($"处理第 {i + 1} 条灯带:\n" +
            //    //    $"当前灯带米数: {当前灯带米数}\n" +
            //    //    $"当前线材长度: {当前线材长度}\n" +
            //    //    $"累计灯带米数: {current灯带米数1}\n" +
            //    //    $"当前组合最大线材长度: {current线材长度1}\n" +
            //    //    $"当前组合序号: {ans1.Count}\n" +
            //    //    $"当前组合内容: {string.Join(", ", ans1.Last())}",
            //    //    "处理过程");
            //}

            for (int i = 0; i < 灯带长度列表.Count; i++)
            {
                double 当前灯带米数 = 灯带长度列表[i];
                double 当前线材长度 = 线长列表[i];
                double 权重值 = 1;

                // 计算如果添加当前灯带后的累计值
                double 预计累计灯带米数 = current灯带米数1 + 当前灯带米数;
                double 预计累计线材长度 = current线材长度1 + 当前线材长度;  // 修改这里，改为累加

                // 计算累计后的权重系数
                double 累计灯带权重系数 = 预计累计灯带米数 / 最大灯带米数;
                double 累计线材权重系数 = 预计累计线材长度 / 最大线材长度;  // 使用累加后的线材长度

                // 根据累计值调整最大限制
                double 调整后最大线材长度 = 最大线材长度 * (1 - 累计灯带权重系数 * 权重值);
                double 调整后最大灯带米数 = 最大灯带米数 * (1 - 累计线材权重系数 * 权重值);

                // 检查是否超出调整后的限制
                bool 超出灯带限制 = 预计累计灯带米数 > 调整后最大灯带米数;
                bool 超出线材限制 = 预计累计线材长度 > 调整后最大线材长度;

                // 添加调试信息
                //MessageBox.Show($"处理第 {i + 1} 条数据:\n" +
                //                $"当前灯带米数: {当前灯带米数:F2}\n" +
                //                $"累计灯带米数: {预计累计灯带米数:F2}\n" +
                //                $"当前线材长度: {当前线材长度:F2}\n" +
                //                $"累计线材长度: {预计累计线材长度:F2}\n" +
                //                $"累计灯带权重系数: {累计灯带权重系数:F2}\n" +
                //                $"累计线材权重系数: {累计线材权重系数:F2}\n" +
                //                $"调整后最大灯带米数: {调整后最大灯带米数:F2}\n" +
                //                $"调整后最大线材长度: {调整后最大线材长度:F2}\n" +
                //                $"是否超出灯带限制: {超出灯带限制}\n" +
                //                $"是否超出线材限制: {超出线材限制}",
                //                "权重调整信息");

                if (超出灯带限制 || 超出线材限制)
                {
                    current灯带米数1 = 0;  // 重置灯带累计值
                    current线材长度1 = 0;  // 重置线材累计值
                    ans1.Add(new List<double>());  // 创建新组合
                }

                // 添加到当前组合并更新累计值
                current灯带米数1 += 当前灯带米数;
                current线材长度1 += 当前线材长度;  // 修改这里，改为累加
                double 当前灯带面积 = 当前灯带米数 * 灯带尺寸.每厘米面积 * 100;
                ans1.Last().Add(当前灯带面积);
            }

            // 移除空组合
            ans1.RemoveAll(x => x.Count == 0);

            // 显示最终结果
            //StringBuilder sb = new StringBuilder();
            //sb.AppendLine($"组合完成，共 {ans1.Count} 个组合:");
            //for (int i = 0; i < ans1.Count; i++)
            //{
            //    sb.AppendLine($"\n组合 {i + 1}:");
            //    sb.AppendLine($"灯带数量: {ans1[i].Count}");
            //    sb.AppendLine($"总米数: {ans1[i].Sum():F2}");
            //    sb.AppendLine($"内容: {string.Join(", ", ans1[i])}");
            //}
            //MessageBox.Show(sb.ToString(), "组合结果");

            return ans1;

        }

        private double 计算线材占用空间(double 线长, double 线径)
        {
            // 计算线材的横截面积
            double 线材横截面积 = Math.PI * Math.Pow(线径 / 2, 2);
            // 计算线材体积（近似空间占用）
            return 线材横截面积 * 线长;
        }

        // 添加一个方法来计算组合的利用率
        public double 计算组合利用率(List<组合结果项> 组合, double target容积)
        {
            double 灯带容积 = target容积 * 0.5;
            double 线材容积 = target容积 * 0.5;

            double 总灯带长度 = 组合.Sum(x => x.灯带长度);
            double 总线长 = 组合.Sum(x => x.线长);

            double 灯带利用率 = 总灯带长度 / 灯带容积;
            double 线材利用率 = 总线长 / 线材容积;

            // 返回两者的平均值
            return (灯带利用率 + 线材利用率) / 2;
        }

        // 添加一个方法来验证组合是否有效
        
    }





}
