namespace 包装计算
{
    public class 新包装资料
    {
        public string 装箱产品型号 { get; set; }
        public string 装箱产品类型 { get; set; }
        public string 包材名称 { get; set; }
        public string 系统编码 { get; set; }
        public int 装箱数量 { get; set; }
        public string 半成品BOM物料码 { get; set; }
        public string 有效容积_ { get; set; }
        public string 系统尺寸 { get; set; }
        public double 总有效容积 { get; set; }
        public double 内撑纸卡容积 { get; set; }
        public double 圆形平卡容积 { get; set; }

        // 其他属性...
    }

    public class 新包装
    {
        private List<新包装资料> 装箱产品列表 = new List<新包装资料>
        {
            new 新包装资料
            {
                装箱产品型号 = "F22",
                装箱产品类型 = "单层/双层",
                包材名称 = "圆盘式天盒",
                系统编码 = "1006.045137",
                装箱数量=1,
                系统尺寸 = "1K8K纸质",
                半成品BOM物料码 = "30.23.010101",
                总有效容积 = 68 * 63 * 4.3,
                内撑纸卡容积 = 0,  // 假设值
                圆形平卡容积 = 0   // 假设值
            },
            new 新包装资料
            {
                装箱产品型号 = "F22",
                装箱产品类型 = "单层/双层",
                包材名称 = "圆盘式天盒",
                系统编码 = "1006.045137",
                装箱数量=1,
                系统尺寸 = "1K8K纸质",
                半成品BOM物料码 = "30.23.010101",
                总有效容积 = 68 * 63 * 4.3,
                内撑纸卡容积 = 0,  // 假设值
                圆形平卡容积 = 0   // 假设值
             }

        };
    }
}