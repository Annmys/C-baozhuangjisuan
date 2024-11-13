namespace 包装计算
{
    public class 新包装资料
    {
        public string? 包装名称 { get; set; }
        public List<string>? 装箱产品型号 { get; set; }
        public List<string>? 装箱产品类型 { get; set; }
        public List<string>? 包材名称 { get; set; }  // 修改为列表
        public string? 系统编码 { get; set; }
        public int 装箱数量 { get; set; }
        public string? 圆盘内全部盘灯米数 { get; set; }
        public string? 圆盘内全部盘线材 { get; set; }
        public string? 半成品BOM物料码 { get; set; }
        public string? 系统尺寸 { get; set; }
        public double 总有效面积 { get; set; }
        public double 内撑纸卡面积 { get; set; }
        public double 圆形平卡面积 { get; set; }
        public double 圆垫板面积 { get; set; }
        public double 垫高纸卡面积 { get; set; }
        public double 高度 { get; set; }
        public double 总有效容积 { get; set; }
        public double 内撑纸卡容积 { get; set; }
        public double 圆形平卡容积 { get; set; }
        public double 圆垫板容积 { get; set; }
        public double 垫高纸卡容积 { get; set; }
        public string? 五盒装选用 { get; set; }
        public string? 三盒装选用 { get; set; }
        public string? 二盒装选用 { get; set; }
        public string? 单盒装选用 { get; set; }

        // 其他属性...
    }

    public class 新包装
    {
        private List<新包装资料> 装箱产品列表 = new List<新包装资料>
    {
            ///开始470mm包装-------------470mm包装------------------470mm包装------------470mm包装-----------------470mm包装-------------------` 

        new 新包装资料
        {
            
            包装名称="圆盘式包装470mm圆盘-01",
            装箱产品型号 = new List<string> { "F16", "F22", "F10", "F23", "F1212" },
            装箱产品类型 = new List<string>
            {
                "F16、单层注塑端部出线、单层注塑底部出线、双层注塑端部出线、双层注塑底部出线",
                "F22、单层注塑正弯端部出线、单层注塑侧弯端部出线、单层注塑侧弯侧面出线、单层注塑正弯底部出线、双层注塑正弯端部出线、双层注塑侧弯端部出线、双层注塑侧弯侧面出线、双层注塑正弯底部出线",
                "F10、单层注塑端部出线、单层注塑侧面出线、双层注塑端部出线、双层注塑侧面出线",
                "F23、单层注塑正弯端部出线、单层注塑侧弯端部出线、单层注塑侧弯侧面出线、单层注塑正弯底部出线、双层注塑正弯端部出线、双层注塑侧弯端部出线、双层注塑侧弯侧面出线、双层注塑正弯底部出线",
                "F1212、单层注塑正弯端部出线、单层注塑侧弯端部出线、单层注塑侧弯侧面出线、单层注塑正弯底部出线"
            },
            包材名称 = new List<string>
            {
                "天盒", "地盒", "加强平卡", "上圆纸板", "下圆纸板",
                "圆纸管", "内撑纸卡", "内五角平卡", "圆形平卡"
            },
            系统编码 = "1006.045137",
            装箱数量 = 1,
            系统尺寸 = "1K8K纸质",
            半成品BOM物料码 = "30.23.010101",
            总有效面积 = 1734,    //1
            内撑纸卡面积 = 314,  //2
            圆形平卡面积 = 706.5,   //3
            高度=1.8,

            总有效容积 = 1734 * 1.8,    //1
            内撑纸卡容积 = 314 * 1.8,  //2
            圆形平卡容积 = 706.5 * 1.8 ,  //3

            单盒装选用="1006.036037",
            三盒装选用="1006.036045",
            五盒装选用="1006.036041",
                //注：
                //1、线槽内可放置10米电线
                //2、产品长度米数小于6米的，请选用多条包装方式
        },

        new 新包装资料
        {
            包装名称="圆盘式包装470mm圆盘-02",
            装箱产品型号 = new List<string> { "F16", "F22", "F10", "F23", "F1212", "F15", "F21" },
            装箱产品类型 = new List<string>
            {
                "F16、单层注塑侧部出线",
                "F22、单层注塑侧弯底部出线、单层注塑正弯侧部出线、双层注塑侧弯底部出线、双层注塑正弯侧部出线",
                "F10、单层注塑底部出线、双层注塑底部出线",
                "F23、单层注塑侧弯底部出线、单层注塑正弯侧部出线、双层注塑侧弯底部出线、双层注塑正弯侧部出线",
                "F1212、单层注塑侧弯底部出线、单层注塑正弯侧部出线",
                "F15、双层注塑底部出线、多条或短条包装",
                "F21、双层注塑底部出线、多条或短条包装"
            },
            包材名称 = new List<string>
            {
                "天盒", "地盒", "加强平卡", "上圆纸板", "下圆纸板", "圆纸管",
                "内撑纸卡", "内五角平卡", "圆形平卡"
            },
            系统编码 = "1006.045143",
            装箱数量 = 1,
            系统尺寸 = "1K8K纸质",
            半成品BOM物料码 = "30.23.010104",
            总有效面积 = 1734,    //1
            内撑纸卡面积 = 314,  //2
            圆形平卡面积 = 706.5,
            垫高纸卡面积=1734,
            高度=4.3,

            总有效容积 = 1734 * 4.3,  //1
            内撑纸卡容积 = 314 * 4.3,  //2
            圆形平卡容积 = 706.5 * 4.3,   //3
            垫高纸卡容积=1734*2.5, //5

            单盒装选用="1006.036038",
            三盒装选用="1006.036041",
            五盒装选用="1006.036042",
                //注：
                //1、线槽内可放置10米电线
                //2、产品长度米数小于6米的，请选用多条包装方式。
                //3.F15/F21双层底部注塑出线产品（多条或短条包装）可选用此包装。【注意：此多条包装是不需要上/下圆纸板、圆纸管、内撑纸卡、内五角平卡及全棉织带物料】
        },

        new 新包装资料
        {
            包装名称 = "圆盘式包装470mm圆盘-03",
            装箱产品型号 = new List<string> { "F16", "F22", "F10", "F23", "F1212" },
            装箱产品类型 = new List<string>
            {
                "F16、单层注塑端部出线、单层注塑底部出线、单层注塑侧部出线、双层注塑端部出线、双层注塑底部出线、双层注塑侧部出线、（多条或短条包装）",
                "F22、单层注塑正弯端部出线、单层注塑侧弯端部出线、单层注塑正弯底部出线、单层注塑侧弯底部出线、单层注塑正弯侧部出线、单层注塑侧弯侧部出线、双层注塑正弯端部出线、双层注塑侧弯端部出线、双层注塑正弯底部出线、双层注塑侧弯侧部出线、（多条或短条包装）",
                "F10、单层注塑端部出线、单层注塑侧部出线、双层注塑端部出线、双层注塑侧部出线、（多条或短条包装）",
                "F23、单层注塑正弯端部出线、单层注塑侧弯端部出线、单层注塑正弯底部出线、单层注塑侧弯底部出线、单层注塑正弯侧部出线、单层注塑侧弯侧部出线、双层注塑正弯端部出线、双层注塑侧弯端部出线、双层注塑正弯底部出线、双层注塑侧弯底部出线、双层注塑正弯侧部出线、双层注塑侧弯侧部出线、（多条或短条包装）",
                "F1212、单层注塑端部出线、单层注塑底部出线、单层注塑侧部出线、双层注塑端部出线、双层注塑底部出线、双层注塑侧部出线、（多条或短条包装）"
            },
            包材名称 = new List<string>
            {
                "天盒", "地盒", "加强平卡", "圆垫板",
            },
            系统编码 = "1006.045141",
            装箱数量 = 1,
            系统尺寸 = "K8K纸质",
            半成品BOM物料码 = "30.23.010105",
            总有效面积 = 1734,
            圆形平卡面积 = 706.5, //???
            圆垫板面积 = 1734,
            高度=3.4,

            总有效容积 = 1734 * 3.4,  //1
            圆垫板容积 = 1734*0.6*2,   //4

            单盒装选用="1006.036037",
            三盒装选用="1006.036045",
            五盒装选用="1006.036041",
            // 注：
            // 1、此套包装为多条灯带或短条灯带包装使用。
            // 2、线槽内可放置10米电线
            // 3、多条灯与短条灯包装优先选用此包装。
            // 4、F22双层注塑侧弯底部出线、正弯侧出线成品/F10双层注塑底部出线成品多条或短条包装从此包装内删除。
        },
        new 新包装资料
        {
            包装名称 = "圆盘式包装470mm圆盘-04",
            装箱产品型号 = new List<string>
            {
                "F15", "F21", "F2008", "F2010", "F2012",
            },
            装箱产品类型 = new List<string>
            {
                "F15、单层注塑端部出线、单层注塑侧部出线、双层注塑端部出线、双层注塑侧部出线",
                "F21、单层注塑端部出线、单层注塑侧面出线、双层注塑端部出线、双层注塑侧部出线",
                "F2008、单层注塑端部出线、单层注塑底部出线、双层注塑端部出线、双层注塑底部出线",
                "F2010、单层注塑端部出线、单层注塑底部出线、双层注塑端部出线、双层注塑底部出线",
                "F2012、单层注塑端部出线、单层注塑底部出线、双层注塑端部出线、双层注塑底部出线"
            },
            包材名称 = new List<string>
            {
                "圆盘式天盒", "地盒", "加强平卡", "上圆纸板", "下圆纸板",
                "圆纸管", "内撑纸卡","内五角平卡", "圆形平卡",
            },
            系统编码 = "1006.045146",
            装箱数量 = 1,
            系统尺寸 = "K8K纸质",
            半成品BOM物料码 = "30.23.010109",
            总有效面积 = 1734,    //1
            内撑纸卡面积 = 314,  //2
            圆形平卡面积 = 706.5,   //3
            高度=2.9,

            总有效容积 = 1734 * 2.9,  //1
            内撑纸卡容积 = 314*2.9,  //2
            圆形平卡容积 = 706.5*2.9,  //3

            单盒装选用="1006.036046",
            三盒装选用="1006.036051",
            五盒装选用="1006.036050",
            //注：
            //1、线槽内可放置10米电线
            //2、产品长度米数小于6米的，请选用多条包装方式。
            },

        new 新包装资料

        {
    包装名称 = "圆盘式包装470mm圆盘-05",
    装箱产品型号 = new List<string>
    {
        "F15", "F21", "F2008", "F2010", "F2012"
    },
    装箱产品类型 = new List<string>
    {
    "F15、单层注塑端部出线、双层注塑端部出线",
    "F21、单层注塑端部出线、双层注塑端部出线",
    "F2008、单层注塑端部出线、双层注塑端部出线",
    "F2010、单层注塑端部出线、双层注塑端部出线",
    "F2012、单层注塑端部出线、双层注塑端部出线"
    },
    包材名称 = new List<string>
    {
        "天盒", "地盒", "加强平卡", "上圆纸板", "下圆纸板",
        "圆纸管", "内撑纸卡", "内五角平卡", "垫高纸卡", "圆形平卡",
    },
    系统编码 = "1006.045148",
    装箱数量 = 1,
    系统尺寸 = "K8K纸质",
    半成品BOM物料码 = "30.23.010110",
    总有效面积 = 1734,    //1
    内撑纸卡面积 = 314,  //2
    圆形平卡面积 = 706.5,   //3
    垫高纸卡面积 = 1734,    //4
    高度 = 5.4,

    总有效容积 = 1734 * 5.4,    //1
    内撑纸卡容积 = 314 * 5.4,  //2
    圆形平卡容积 = 706.5 * 5.4,   //3
    垫高纸卡容积 = 1734 * 2.5,    //4

    单盒装选用 = "1006.036047",
    三盒装选用 = "1006.036056",
    五盒装选用 = "1006.036052",

    //注：
    //1、线槽内可放置10米电线
    //2、产品长度米数小于6米的，请选用多条包装方式。
        },


        new 新包装资料
{
    包装名称 = "圆盘式包装470mm圆盘-06",
    装箱产品型号 = new List<string>
    {
        "F15", "F21", "F2008", "F2010", "F2012", "F22", "F10"
    },
    装箱产品类型 = new List<string>
    {
        "F15、单层注塑端部出线、单层注塑底部出线、单层注塑侧部出线、双层注塑端部出线、双层注塑底部出线、双层注塑侧部出线、（多米数短条包装）",
        "F21、单层注塑端部出线、单层注塑底部出线、单层注塑侧部出线、双层注塑端部出线、双层注塑底部出线、双层注塑侧部出线、（多米数短条包装）",
        "F2008、单层注塑端部出线、单层注塑底部出线、单层注塑侧部出线、双层注塑端部出线、双层注塑底部出线、双层注塑侧部出线、（多米数短条包装）",
        "F2010、单层注塑端部出线、单层注塑底部出线、单层注塑侧部出线、双层注塑端部出线、双层注塑底部出线、双层注塑侧部出线、（多米数短条包装）",
        "F2012、单层注塑端部出线、单层注塑底部出线、单层注塑侧部出线、双层注塑端部出线、双层注塑底部出线、双层注塑侧部出线、（多米数短条包装）",
        "F22、双层注塑侧弯底部出线、双层注塑正弯侧部出线、（多米数短条包装）",
        "F10、双层主型底部出线、（多米数短条包装）"
    },
    包材名称 = new List<string>
    {
        "天盒", "地盒", "加强平卡", "圆垫板", "FOF热缩袋",
        "通口箱", "纸箱", "封箱胶纸-高水平皮胶纸"
    },
    系统编码 = "1006.045146",
    装箱数量 = 1,
    系统尺寸 = "K8K纸质",
    半成品BOM物料码 = "30.23.010111",
    总有效面积 = 1734,    //1
    圆垫板面积 = 1734,    //2
    高度 = 4.5,

    总有效容积 = 1734 * 4.5,    //1
    圆垫板容积 = 1734 * 0.6 * 2,    //2

    单盒装选用 = "1006.036046",
    三盒装选用 = "1006.036051",
    五盒装选用 = "1006.036050",

//注：
//1、此套包装为多条灯带或短条灯带包装使用。
//2、线槽内可放置10米电线
//3、多条灯与短条灯包装优先选用此包装。
//4.新增F22双层注塑侧弯底部出线、正弯侧出线成品/F10双层注塑底部出线成品多条或短条包装。
},


        ///开始600mm包装-------------600mm包装------------------600mm包装------------600mm包装-----------------600mm包装-------------------
        new 新包装资料
{
    包装名称 = "圆盘式包装600mm圆盘-01",
    装箱产品型号 = new List<string>
    {
        "F16", "F22", "F10", "F23", "F1212"
    },
    装箱产品类型 = new List<string>
{
    "F16、单层注塑端部出线、单层注塑底部出线、双层注塑端部出线、双层注塑底部出线",
    "F22、单层注塑正弯端部出线、单层注塑侧弯端部出线、单层注塑侧弯底部出线、单层注塑正弯底部出线、双层注塑正弯端部出线、双层注塑侧弯端部出线、双层注塑侧弯底部出线、双层注塑正弯底部出线",
    "F10、单层注塑端部出线、单层注塑侧部出线、双层注塑端部出线、双层注塑侧部出线",
    "F23、单层注塑正弯端部出线、单层注塑侧弯端部出线、单层注塑侧弯侧部出线、单层注塑正弯底部出线、双层注塑正弯端部出线、双层注塑侧弯端部出线、双层注塑侧弯侧部出线、双层注塑正弯底部出线",
    "F1212、单层注塑端部出线、单层注塑侧弯侧部出线、单层注塑正弯底部出线成品"
},
    包材名称 = new List<string>
    {
        "天盒", "地盒", "加强平卡-Y1", "上圆纸板", "下圆纸板",
        "圆纸管", "内撑纸卡", "内五角平卡", "圆形平卡",
    },
    系统编码 = "1006.045137",
    装箱数量 = 1,
    系统尺寸 = "K8K纸质",
    半成品BOM物料码 = "30.23.010101",
    总有效面积 = 2826,    //1
    内撑纸卡面积 = 314,  //2
    圆形平卡面积 = 706.5,   //3
    高度 = 1.8,

    总有效容积 = 2826 * 1.8,    //1
    内撑纸卡容积 = 314 * 1.8,  //2
    圆形平卡容积 = 706.5 * 1.8,   //3

    单盒装选用 = "1006.036035",
    二盒装选用 = "1006.036062",
    三盒装选用 = "1006.036044",

    //注：
    //1、线槽内可放置10米电线
    //2、产品长度米数小于6米的，请选用多条包装方式。
    //3、新单内含/箱包装。
    //4、删除5盒/箱包装的纸箱。
},

        new 新包装资料
{
    包装名称 = "圆盘式包装600mm圆盘-02",
    装箱产品型号 = new List<string>
    {
        "F16", "F22", "F10", "F23", "F1212", "F15", "F21"
    },
    装箱产品类型 = new List<string>
{
    "F16、单层注塑侧部出线",
    "F22、单层注塑侧弯底部出线、单层注塑正弯侧部出线、双层注塑侧弯底部出线、双层注塑正弯侧部出线",
    "F10、单层注塑底部出线、双层注塑底部出线",
    "F23、单层注塑侧弯底部出线、单层注塑正弯侧部出线、双层注塑侧弯底部出线、双层注塑正弯侧部出线",
    "F1212、单层注塑侧弯底部出线、单层注塑正弯侧部出线",
    "F15、双层注塑底部出线产品、（多米或短条包装）",
    "F21、双层注塑底部出线产品、（多米或短条包装）"
},
    包材名称 = new List<string>
    {
        "天盒", "地盒", "加强平卡-Y1", "上圆纸板", "下圆纸板",
        "垫高纸卡", "圆纸管", "内撑纸卡", "内五角平卡", "圆形平卡",
    },
    系统编码 = "1006.045139",
    装箱数量 = 1,
    系统尺寸 = "K8K纸质",
    半成品BOM物料码 = "30.23.010102",
    总有效面积 = 2826,    //1
    内撑纸卡面积 = 314,  //2
    圆形平卡面积 = 706.5,   //3
    垫高纸卡面积 = 2826,    //4
    高度 = 4.3,

    总有效容积 = 2826 * 4.3,    //1
    内撑纸卡容积 = 314 * 4.3,  //2
    圆形平卡容积 = 706.5 * 4.3,   //3
    垫高纸卡容积 = 2826 * 2.5,    //4

    单盒装选用 = "1006.036036",
    二盒装选用 = "1006.036063",
    三盒装选用 = "1006.036039",
    五盒装选用 = "1006.036040",

//注：
//1、线槽内可放置10米电线
//2、产品长度米数小于6米的，请选用多条包装方式。
//3.F15/F21双层底部注塑出线产品（多条或短条包装）可选用此包装。【注意：此多条包装是不需要上/下圆纸板、圆纸管、内撑纸卡、内五角平卡及全棉织带物料】
//4、新增两盒/箱包装。
},

        new 新包装资料
{
    包装名称 = "圆盘式包装600mm圆盘-03",
    装箱产品型号 = new List<string>
    {
        "F16", "F22", "F10", "F23", "F1212"
    },
    装箱产品类型 = new List<string>
{
    "F16、单层注塑端部出线、单层注塑底部出线、单层注塑侧部出线、双层注塑端部出线、双层注塑底部出线、双层注塑侧部出线、（多条或短条包装）",
    "F22、单层注塑正弯端部出线、单层注塑侧弯端部出线、单层注塑正弯底部出线、单层注塑侧弯侧部出线、双层注塑正弯端部出线、双层注塑侧弯端部出线、双层注塑正弯底部出线、双层注塑侧弯侧部出线、（多条或短条包装）",
    "F10、单层注塑端部出线、单层注塑侧部出线、双层注塑端部出线、双层注塑侧部出线、（多条或短条包装）",
    "F23、单层注塑正弯端部出线、单层注塑侧弯端部出线、单层注塑正弯底部出线、单层注塑侧弯底部出线、单层注塑正弯侧部出线、单层注塑侧弯侧部出线、双层注塑正弯端部出线、双层注塑侧弯端部出线、双层注塑正弯底部出线、双层注塑侧弯底部出线、双层注塑正弯侧部出线、双层注塑侧弯侧部出线、（多条或短条包装）",
    "F1212、单层注塑端部出线、单层注塑底部出线、单层注塑侧部出线、双层注塑端部出线、双层注塑底部出线、双层注塑侧部出线、（多条或短条包装）"
},
    包材名称 = new List<string>
    {
        "天盒", "地盒", "加强平卡-Y1", "圆垫板", 
    },
    系统编码 = "1006.045137",
    装箱数量 = 1,
    系统尺寸 = "K8K纸质",
    半成品BOM物料码 = "30.23.010106",
    总有效面积 = 2826,    //1
    圆垫板面积 = 2826,    //2
    高度 = 3.4,

    总有效容积 = 2826 * 3.4,    //1
    圆垫板容积 = 2826 * 0.6 * 2,    //2

    单盒装选用 = "1006.036035",
    二盒装选用 = "1006.036062",
    三盒装选用 = "1006.036044",

//注：
//1.此套包装为多条灯带或短条灯带包装使用。  
//2.线槽内可放置10米电线
//3.多条灯与短条灯包装优先选用此包装。
//4.F22双层注塑侧弯底部出线、正弯侧出线成品/F10双层注塑底部出线成品多条或短条包装从此包装内删除。
//5、新增两盒/箱包装。
//6、删除5盒/箱包装的纸箱。
},

        new 新包装资料
{
    包装名称 = "圆盘式包装600mm圆盘-04",
    装箱产品型号 = new List<string>
    {
        "F15", "F21", "F2008", "F2010", "F2012", "F2219", "F2222"
    },
    装箱产品类型 = new List<string>
    {
        "F15、单层注塑端部出线、单层注塑侧部出线、双层注塑端部出线、双层注塑侧部出线",
        "F21、单层注塑端部出线、单层注塑侧部出线、双层注塑端部出线、双层注塑侧部出线",
        "F2008、单层注塑端部出线、单层注塑底部出线、双层注塑端部出线、双层注塑底部出线",
        "F2010、单层注塑端部出线、单层注塑底部出线、双层注塑端部出线、双层注塑底部出线",
        "F2012、单层注塑端部出线、单层注塑底部出线、双层注塑端部出线、双层注塑底部出线",
        "F2219、单层注塑正弯端部出线、单层注塑正弯底部出线、双层注塑正弯端部出线、双层注塑正弯底部出线、单层注塑侧弯端部出线、单层注塑侧弯底部出线、双层注塑侧弯端部出线、双层注塑侧弯底部出线",
        "F2222、单层注塑正弯端部出线、单层注塑正弯底部出线、双层注塑正弯端部出线、双层注塑正弯底部出线、单层注塑侧弯端部出线、单层注塑侧弯底部出线、双层注塑侧弯端部出线、双层注塑侧弯底部出线"
    },
    包材名称 = new List<string>
    {
        "天盒", "地盒", "加强平卡", "上圆纸板", "下圆纸板",
        "圆纸管", "内撑纸卡", "内五角平卡", "圆形平卡",
    },
    系统编码 = "1006.045150",
    装箱数量 = 1,
    系统尺寸 = "K8K纸质",
    半成品BOM物料码 = "30.23.010107",
    总有效面积 = 2826,    //1
    内撑纸卡面积 = 314,  //2
    圆形平卡面积 = 706.5,   //3
    高度 = 2.9,

    总有效容积 = 2826 * 2.9,    //1
    内撑纸卡容积 = 314 * 2.9,  //2
    圆形平卡容积 = 706.5 * 2.9,   //3

    单盒装选用 = "1006.036048",
    二盒装选用 = "1006.036064",
    三盒装选用 = "1006.036054",
    五盒装选用 = "1006.036053",

//注：
//1、线槽内可放置10米电线
//2、产品长度米数小于6米的，请选用多条包装方式。
//3、新增两盒/箱包装。
},

        new 新包装资料
{
    包装名称 = "圆盘式包装600mm圆盘-05",
    装箱产品型号 = new List<string>
    {
        "F15", "F21", "F2008", "F2010", "F2012", "F2219", "F2222"
    },
    装箱产品类型 = new List<string>
    {
        "F15、单层注塑底部出线、双层注塑底部出线",
        "F21、单层注塑底部出线、双层注塑底部出线",
        "F2008、单层注塑侧部出线、双层注塑侧部出线",
        "F2010、单层注塑侧部出线、双层注塑侧部出线",
        "F2012、单层注塑侧部出线、双层注塑侧部出线",
        "F2219、单层注塑正弯侧部出线、双层注塑正弯侧部出线、单层注塑侧弯侧部出线、双层注塑侧弯侧部出线",
        "F2222、单层注塑正弯底部出线、双层注塑正弯底部出线、单层注塑侧弯底部出线、双层注塑侧弯底部出线"
    },
    包材名称 = new List<string>
    {
        "地盒", "加强平卡", "上圆纸板", "下圆纸板", "垫高纸卡",
        "圆纸管", "内撑纸卡", "内五角平卡", "圆形平卡", 
    },
    系统编码 = "1006.045152",
    装箱数量 = 1,
    系统尺寸 = "K8K纸质",
    半成品BOM物料码 = "30.23.010108",
    总有效面积 = 2826,    //1
    内撑纸卡面积 = 314,  //2
    圆形平卡面积 = 706.5,   //3
    垫高纸卡面积 = 2826,    //4
    高度 = 5.4,

    总有效容积 = 2826 * 5.4,    //1
    内撑纸卡容积 = 314 * 5.4,  //2
    圆形平卡容积 = 706.5 * 5.4,   //3
    垫高纸卡容积 = 2826 * 2.5,    //4

    单盒装选用 = "1006.036049",
    二盒装选用 = "1006.036065",
    三盒装选用 = "1006.036057",
    五盒装选用 = "1006.036055",
//    注：
//1、线槽内可放置10米电线
//2、产品长度米数小于6米的，请选用多条包装方式。
//3、新增两盒/箱包装。
},

        new 新包装资料
{
    包装名称 = "圆盘式包装600mm圆盘-06",
    装箱产品型号 = new List<string>
    {
        "F15", "F21", "F2008", "F2010", "F2012", "F2219", "F2222", "F22", "F10"
    },
    装箱产品类型 = new List<string>
    {
        "F15、单层注塑端部出线、单层注塑侧部出线、双层注塑端部出线、双层注塑侧部出线、（多米数短条包装）",
        "F21、单层注塑端部出线、单层注塑侧部出线、双层注塑端部出线、双层注塑侧部出线、（多米数短条包装）",
        "F2008、单层注塑端部出线、单层注塑底部出线、单层注塑侧部出线、双层注塑端部出线、双层注塑底部出线、双层注塑侧部出线、（多米数短条包装）",
        "F2010、单层注塑端部出线、单层注塑底部出线、单层注塑侧部出线、双层注塑端部出线、双层注塑底部出线、双层注塑侧部出线、（多米数短条包装）",
        "F2012、单层注塑端部出线、单层注塑底部出线、单层注塑侧部出线、双层注塑端部出线、双层注塑底部出线、双层注塑侧部出线、（多米数短条包装）",
        "F2219、单层注塑正弯端部出线、单层注塑正弯底部出线、单层注塑正弯侧部出线、双层注塑正弯端部出线、双层注塑正弯底部出线、双层注塑正弯侧部出线、、单层注塑侧弯端部出线、单层注塑侧弯底部出线、单层注塑侧弯侧部出线、双层注塑侧弯端部出线、双层注塑侧弯底部出线、双层注塑侧弯侧部出线、（多米数短条包装）",
        "F2222、单层注塑正弯端部出线、单层注塑正弯底部出线、单层注塑正弯侧部出线、双层注塑正弯端部出线、双层注塑正弯底部出线、双层注塑正弯侧部出线、、单层注塑侧弯端部出线、单层注塑侧弯底部出线、单层注塑侧弯侧部出线、双层注塑侧弯端部出线、双层注塑侧弯底部出线、双层注塑侧弯侧部出线、（多米数短条包装）",
        "F22、双层注塑侧弯底部出线、双层注塑正弯侧部出线、（多米数短条包装）",
        "F10、双层注塑底部出线、（多米数短条包装）"
    },
    包材名称 = new List<string>
    {
        "天盒", "地盒", "加强平卡", "圆垫板", 
    },
    系统编码 = "1006.045150",
    装箱数量 = 1,
    系统尺寸 = "K8K纸质",
    半成品BOM物料码 = "30.23.010112",
    总有效面积 = 2826,    //1
    圆垫板面积 = 2826,    //2
    高度 = 4.5,

    总有效容积 = 2826 * 4.5,    //1
    圆垫板容积 = 2826 * 0.6 * 2,    //2

    单盒装选用 = "1006.036048",
    二盒装选用 = "1006.036064",
    三盒装选用 = "1006.036054",
    五盒装选用 = "1006.036053",

    //注：
    //1、此套包装为多条灯带或短灯带包装使用。
    //2、线槽内可放置10米电线
    //3、多条灯与短条灯包装优先选用此包装。
    //4、新增22双层注塑侧弯底部出线、正弯侧出线成品/F10双层注塑底部出线成品多条或短条包装。
},

            // 其他包装资料...
        };

        public 新包装资料? 查找包装资料(string 产品型号, string 产品类型)
        {
            var 包装资料 = 装箱产品列表
                .FirstOrDefault(包装 =>
                    包装.装箱产品型号 != null && 包装.装箱产品型号.Contains(产品型号) &&
                    包装.装箱产品类型 != null && 包装.装箱产品类型.Any(类型 => 类型.Contains(产品类型)));

            return 包装资料 ?? null; // 明确返回 null，避免警告
        }
    }



}





    //包装清单_表格转格式_算法 使用方法：
    //var converter = new TableConverter();
    //DataTable table = // 从Excel或其他来源获取的表格数据
    //string code = converter.ConvertTableToCode(table);


//包装清单_表格转格式_算法 算法代码：
//   public class TableConverter
//{
//    public class PackageInfo
//    {
//        public string Name { get; set; }
//        public List<string> Models { get; set; }
//        public List<string> Types { get; set; }
//        public List<string> Materials { get; set; }
//        public Dictionary<string, decimal> Dimensions { get; set; }
//        public string SystemCode { get; set; }
//        public string BomCode { get; set; }
//    }

//    public string ConvertTableToCode(DataTable table)
//    {
//        try
//        {
//            var packageInfo = ExtractPackageInfo(table);
//            return FormatToCode(packageInfo);
//        }
//        catch (Exception ex)
//        {
//            return $"转换失败: {ex.Message}";
//        }
//    }

//    private PackageInfo ExtractPackageInfo(DataTable table)
//    {
//        return new PackageInfo
//        {
//            Name = ExtractBasicInfo(table),
//            Models = ExtractProductModels(table),
//            Types = ExtractProductTypes(table),
//            Materials = ExtractPackageMaterials(table),
//            Dimensions = ExtractDimensions(table),
//            SystemCode = GetCellValue(table, "系统编码"),
//            BomCode = GetCellValue(table, "半成品BOM物料码")
//        };
//    }

//    private string ExtractBasicInfo(DataTable table)
//    {
//        var diameter = table.TableName.Contains("600") ? "600" : "470";
//        var number = GetPackageNumber(table);
//        return $"圆盘式包装{diameter}mm圆盘-{number}";
//    }

//    private List<string> ExtractProductModels(DataTable table)
//    {
//        var models = new HashSet<string>();
//        foreach (DataRow row in table.Rows)
//        {
//            var model = row["装箱产品型号"]?.ToString();
//            if (!string.IsNullOrEmpty(model))
//            {
//                models.Add(model);
//            }
//        }
//        return models.ToList();
//    }

//    private List<string> ExtractProductTypes(DataTable table)
//    {
//        var types = new List<string>();
//        foreach (DataRow row in table.Rows)
//        {
//            var type = FormatProductType(row["装箱产品类型"]?.ToString());
//            if (!string.IsNullOrEmpty(type))
//            {
//                types.Add(type);
//            }
//        }
//        return types;
//    }

//    private string FormatProductType(string type)
//    {
//        if (string.IsNullOrEmpty(type)) return string.Empty;

//        // 格式化产品类型字符串
//        var parts = type.Split('、');
//        var model = parts[0];
//        var details = string.Join("、", parts.Skip(1));
        
//        return $"{model}、{details}";
//    }

//    private List<string> ExtractPackageMaterials(DataTable table)
//    {
//        var materials = new HashSet<string>();
//        foreach (DataRow row in table.Rows)
//        {
//            var material = row["包材名称"]?.ToString();
//            if (!string.IsNullOrEmpty(material))
//            {
//                materials.Add(material);
//            }
//        }
//        return materials.ToList();
//    }

//    private Dictionary<string, decimal> ExtractDimensions(DataTable table)
//    {
//        var dimensions = new Dictionary<string, decimal>();
        
//        // 提取基本尺寸
//        dimensions["总有效面积"] = GetDecimalValue(table, "总有效面积");
//        dimensions["高度"] = GetDecimalValue(table, "高度");
        
//        // 计算容积
//        dimensions["总有效容积"] = dimensions["总有效面积"] * dimensions["高度"];

//        // 提取其他尺寸
//        var dimensionTypes = new[] { "内撑纸卡", "圆形平卡", "垫高纸卡", "圆垫板" };
//        foreach (var dimType in dimensionTypes)
//        {
//            var area = GetDecimalValue(table, $"{dimType}面积");
//            if (area > 0)
//            {
//                dimensions[$"{dimType}面积"] = area;
//                dimensions[$"{dimType}容积"] = area * dimensions["高度"];
//            }
//        }

//        return dimensions;
//    }

//    private string FormatToCode(PackageInfo info)
//    {
//        var sb = new StringBuilder();
//        sb.AppendLine("new 新包装资料");
//        sb.AppendLine("{");
//        sb.AppendLine($"    包装名称 = \"{info.Name}\",");
        
//        // 格式化产品型号列表
//        sb.AppendLine("    装箱产品型号 = new List<string>");
//        sb.AppendLine("    {");
//        sb.AppendLine($"        {string.Join(",\r\n        ", info.Models.Select(m => $"\"{m}\""))}");
//        sb.AppendLine("    },");
        
//        // 格式化产品类型列表
//        sb.AppendLine("    装箱产品类型 = new List<string>");
//        sb.AppendLine("    {");
//        sb.AppendLine($"        {string.Join(",\r\n        ", info.Types.Select(t => $"\"{t}\""))}");
//        sb.AppendLine("    },");
        
//        // 格式化包材名称列表
//        sb.AppendLine("    包材名称 = new List<string>");
//        sb.AppendLine("    {");
//        sb.AppendLine($"        {string.Join(",\r\n        ", info.Materials.Select(m => $"\"{m}\""))}");
//        sb.AppendLine("    },");
        
//        // 添加基本信息
//        sb.AppendLine($"    系统编码 = \"{info.SystemCode}\",");
//        sb.AppendLine("    装箱数量 = 1,");
//        sb.AppendLine("    系统尺寸 = \"K8K纸质\",");
//        sb.AppendLine($"    半成品BOM物料码 = \"{info.BomCode}\",");
        
//        // 添加尺寸信息
//        foreach (var dim in info.Dimensions)
//        {
//            sb.AppendLine($"    {dim.Key} = {dim.Value},");
//        }

//        sb.AppendLine("},");
        
//        return sb.ToString();
//    }

//    private string GetCellValue(DataTable table, string columnName)
//    {
//        foreach (DataRow row in table.Rows)
//        {
//            if (row[columnName] != null && row[columnName] != DBNull.Value)
//            {
//                return row[columnName].ToString();
//            }
//        }
//        return string.Empty;
//    }

//    private decimal GetDecimalValue(DataTable table, string columnName)
//    {
//        var value = GetCellValue(table, columnName);
//        return decimal.TryParse(value, out decimal result) ? result : 0;
//    }
//}

