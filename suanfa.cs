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



    





}
