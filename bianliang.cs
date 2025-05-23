﻿using System;
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

        public List<匹配信息> 订单附件匹配列表 = new List<匹配信息>();

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
    public int 出现次数 { get; set; }  // 新增属性

    // 原有的构造函数
    public 数据项(string a, string o, double r, string id)
    {
        内容A = a;
        内容O = o;
        内容R = r;
        标志 = id;
        出现次数 = 1;  // 默认值
    }

    // 新增带出现次数的构造函数
    public 数据项(string a, string o, double r, string id, int count)
    {
        内容A = a;
        内容O = o;
        内容R = r;
        标志 = id;
        出现次数 = count;
    }
}



public class 匹配信息
{
    public string 订单编号 { get; set; }
    public string 产品型号 { get; set; }
    public HashSet<string> 出线方式 { get; set; }
    public double 销售数量 { get; set; }
    public string 工作表名称 { get; set; }
    public double 工作表总米数 { get; set; }
    public Dictionary<string, string> A列序号字母映射 { get; set; } = new Dictionary<string, string>();  // 新增属性，用于存储A列序号与字母的映射

    public 匹配信息(string 订单编号, string 产品型号, HashSet<string> 出线方式, double 销售数量, string 工作表名称, double 工作表总米数)
    {
        this.订单编号 = 订单编号;
        this.产品型号 = 产品型号;
        this.出线方式 = 出线方式;
        this.销售数量 = 销售数量;
        this.工作表名称 = 工作表名称;
        this.工作表总米数 = 工作表总米数;
    }

    // 新增方法：添加A列序号映射
    

    public string Sheet序号前缀 { get; private set; } = "";  // 新增属性，用于存储工作表中A列的序号前缀（如"R"）

    public void 设置Sheet序号前缀(List<string> 表数据)
    {
        if (表数据.Count > 0)
        {
            // 获取第一行数据的序号部分
            string 首行数据 = 表数据[0];
            string 序号 = 首行数据.Split(',')[0].Trim();  // 获取逗号前的序号部分
            // 提取序号中的字母部分
            Sheet序号前缀 = new string(序号.TakeWhile(c => !char.IsDigit(c)).ToArray());
        }
    }


    public override string ToString()
    {
        string 出线方式字符串 = 出线方式.Count > 0 ? string.Join("，", 出线方式) : "无";
        string 序号映射字符串 = string.Join(", ", A列序号字母映射.Select(x => $"{x.Key}: {x.Value}"));

        return $"找到匹配的订单和附件：\n" +
               $"订单编号：{订单编号}\n" +
               $"产品型号：{产品型号}\n" +
               $"出线方式：{出线方式字符串}\n" +
               $"销售数量：{销售数量:F3}\n" +
               $"匹配工作表：{工作表名称}\n" +
               $"工作表总米数：{工作表总米数:F3}\n" +
               $"A列序号映射：{序号映射字符串}";
    }
}