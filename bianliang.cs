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
        public string? 订单编号;

        public List<string> 订单型号列表 = new List<string>();
        public Dictionary<string, List<string>> 附件表数据 = new Dictionary<string, List<string>>();

        // 在变量类中更新订单出线字典的定义
        public Dictionary<string, List<(string 型号, HashSet<string> 出线方式, string F列内容, double 销售数量)>> 订单出线字典;

        public List<double> 测试 = new List<double>();

        public List<string> 公司型号列表 = new List<string>
        { "F10","F15","F16","F21","F22","F23","F1212","F2008","F2010","F2012","F2019","F2222" };

        public List<List<double>> 订单附件 = new List<List<double>> { new List<double>() };

        public double 查找组合_基数 = 1900;  // 查找组合的基数

        public List<灯带尺寸> 灯带尺寸列表 = new List<灯带尺寸>   // 灯带尺寸列表
        {
            //灯带尺寸名称,横截面宽度,横截面高度，单位是CM
            new 灯带尺寸("F10",0.9,1.8),
            new 灯带尺寸("F15",1.15,2.1),
            new 灯带尺寸("F16",1.55,0.6),
            new 灯带尺寸("F21",1.15,2.9),
            new 灯带尺寸("F22",1.6,1.7),
            new 灯带尺寸("F23",1.0,1.0),
            new 灯带尺寸("F1212",1.2,1.2),
            new 灯带尺寸("F2008",2.0,0.8),
            new 灯带尺寸("F2010",2.0,1.0),
            new 灯带尺寸("F2012",2.0,1.2),
            new 灯带尺寸("F2219",2.2,1.9),
            new 灯带尺寸("F2222",2.2,2.2)
        };


    }

    public class 灯带尺寸
    {
        public string 型号 { get; set; }
        public double 横截面宽度 { get; set; }
        public double 横截面高度 { get; set; }
        public double 每厘米面积 { get; set; }

        public 灯带尺寸(string 型号, double 宽度, double 高度)
        {
            this.型号 = 型号;
            this.横截面宽度 = 宽度;
            this.横截面高度 = 高度;
            this.每厘米面积 = 宽度 * 1; //单位是CM
        }

        // 重写 ToString 方法以便打印
        public override string ToString()
        {
            return $"{型号} - 宽度:{横截面宽度} - 高度:{横截面高度} - 面积:{每厘米面积}";
        }
    }
}

public class 数据项
{
    public string 内容A { get; set; }
    public string 内容O { get; set; }
    public double 内容R { get; set; }
    public string 标志 { get; set; }

    public 数据项(string a, string o, double r, string id)
    {
        内容A = a;
        内容O = o;
        内容R = r;
        标志 = id;
    }
}

