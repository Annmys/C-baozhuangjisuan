using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace 包装计算
{
    public class 变量
    {
        public string? 订单excel地址;
        public string? 附件excel地址;

        public List<string> 订单型号列表 = new List<string>();
        public Dictionary<string, List<string>> 附件表数据 = new Dictionary<string, List<string>>();

        public List<double> 测试 = new List<double>();

        public List<string> 公司型号列表 = new List<string>
        { "F10","F15","F16","F21","F22","F23","F1212","F2008","F2010","F2012","F2019","F2222" };

        public List<List<double>> 订单附件 = new List<List<double>> { new List<double>() };

        public int 查找组合_基数 = 1900;  // 查找组合的基数

        public List<灯带尺寸> 灯带尺寸列表 = new List<灯带尺寸>   // 灯带尺寸列表
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


    }

    public class 灯带尺寸
    {
        public string 型号 { get; set; }
        public double 宽度 { get; set; }
        public double 高度 { get; set; }
        public double 每米面积 { get; set; }

        public 灯带尺寸(string 型号, double 宽度, double 高度)
        {
            this.型号 = 型号;
            this.宽度 = 宽度;
            this.高度 = 高度;
            this.每米面积 = 宽度 * 10; //单位是CM
        }

        // 重写 ToString 方法以便打印
        public override string ToString()
        {
            return $"{型号} - 宽度:{宽度} - 高度:{高度} - 面积:{每米面积}";
        }
    }
}