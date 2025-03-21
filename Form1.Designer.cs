namespace 包装计算
{
    partial class Form1
    {
        /// <summary>
        ///  Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        ///  Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        ///  Required method for Designer support - do not modify
        ///  the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            components = new System.ComponentModel.Container();
            Sunny.UI.UIStyleManager uiStyleManager2;
            uiStyleManager1 = new Sunny.UI.UIStyleManager(components);
            button_订单导入 = new Sunny.UI.UIButton();
            button_附件导入 = new Sunny.UI.UIButton();
            button_开始处理 = new Sunny.UI.UIButton();
            uiTextBox_订单地址 = new Sunny.UI.UITextBox();
            uiButton1 = new Sunny.UI.UIButton();
            uiTextBox_附件地址 = new Sunny.UI.UITextBox();
            uiTextBox_状态 = new Sunny.UI.UITextBox();
            uiButton2 = new Sunny.UI.UIButton();
            uiCheckBox_RU客户 = new Sunny.UI.UICheckBox();
            uiButton3 = new Sunny.UI.UIButton();
            uiStyleManager2 = new Sunny.UI.UIStyleManager(components);
            SuspendLayout();
            // 
            // button_订单导入
            // 
            button_订单导入.Font = new Font("宋体", 12F, FontStyle.Regular, GraphicsUnit.Point, 134);
            button_订单导入.Location = new Point(268, 40);
            button_订单导入.MinimumSize = new Size(1, 1);
            button_订单导入.Name = "button_订单导入";
            button_订单导入.Size = new Size(71, 26);
            button_订单导入.TabIndex = 9;
            button_订单导入.Text = "订单导入";
            button_订单导入.TipsFont = new Font("宋体", 9F, FontStyle.Regular, GraphicsUnit.Point, 134);
            button_订单导入.Click += button_订单导入_Click;
            // 
            // button_附件导入
            // 
            button_附件导入.Font = new Font("宋体", 12F, FontStyle.Regular, GraphicsUnit.Point, 134);
            button_附件导入.Location = new Point(268, 76);
            button_附件导入.MinimumSize = new Size(1, 1);
            button_附件导入.Name = "button_附件导入";
            button_附件导入.Size = new Size(71, 26);
            button_附件导入.TabIndex = 11;
            button_附件导入.Text = "附件导入";
            button_附件导入.TipsFont = new Font("宋体", 9F, FontStyle.Regular, GraphicsUnit.Point, 134);
            button_附件导入.Click += button_附件导入_Click_1;
            // 
            // button_开始处理
            // 
            button_开始处理.Font = new Font("宋体", 12F, FontStyle.Regular, GraphicsUnit.Point, 134);
            button_开始处理.Location = new Point(313, 176);
            button_开始处理.MinimumSize = new Size(1, 1);
            button_开始处理.Name = "button_开始处理";
            button_开始处理.Size = new Size(26, 26);
            button_开始处理.TabIndex = 13;
            button_开始处理.Text = "V1";
            button_开始处理.TipsFont = new Font("宋体", 9F, FontStyle.Regular, GraphicsUnit.Point, 134);
            button_开始处理.Click += button_开始处理_Click;
            // 
            // uiTextBox_订单地址
            // 
            uiTextBox_订单地址.Font = new Font("宋体", 12F);
            uiTextBox_订单地址.Location = new Point(4, 40);
            uiTextBox_订单地址.Margin = new Padding(4, 5, 4, 5);
            uiTextBox_订单地址.MinimumSize = new Size(1, 16);
            uiTextBox_订单地址.Name = "uiTextBox_订单地址";
            uiTextBox_订单地址.Padding = new Padding(5);
            uiTextBox_订单地址.ShowText = false;
            uiTextBox_订单地址.Size = new Size(257, 29);
            uiTextBox_订单地址.TabIndex = 15;
            uiTextBox_订单地址.TextAlignment = ContentAlignment.MiddleLeft;
            uiTextBox_订单地址.Watermark = "";
            uiTextBox_订单地址.TextChanged += uiTextBox_订单地址_TextChanged;
            // 
            // uiButton1
            // 
            uiButton1.Font = new Font("宋体", 12F, FontStyle.Regular, GraphicsUnit.Point, 134);
            uiButton1.Location = new Point(373, 79);
            uiButton1.MinimumSize = new Size(1, 1);
            uiButton1.Name = "uiButton1";
            uiButton1.Size = new Size(55, 26);
            uiButton1.TabIndex = 18;
            uiButton1.Text = "uiButton1";
            uiButton1.TipsFont = new Font("宋体", 9F, FontStyle.Regular, GraphicsUnit.Point, 134);
            uiButton1.Visible = false;
            uiButton1.Click += uiButton1_Click;
            // 
            // uiTextBox_附件地址
            // 
            uiTextBox_附件地址.Font = new Font("宋体", 12F);
            uiTextBox_附件地址.Location = new Point(4, 76);
            uiTextBox_附件地址.Margin = new Padding(4, 5, 4, 5);
            uiTextBox_附件地址.MinimumSize = new Size(1, 16);
            uiTextBox_附件地址.Name = "uiTextBox_附件地址";
            uiTextBox_附件地址.Padding = new Padding(5);
            uiTextBox_附件地址.ShowText = false;
            uiTextBox_附件地址.Size = new Size(257, 29);
            uiTextBox_附件地址.TabIndex = 16;
            uiTextBox_附件地址.TextAlignment = ContentAlignment.MiddleLeft;
            uiTextBox_附件地址.Watermark = "";
            // 
            // uiTextBox_状态
            // 
            uiTextBox_状态.BackColor = Color.Silver;
            uiTextBox_状态.Font = new Font("宋体", 12F);
            uiTextBox_状态.Location = new Point(4, 115);
            uiTextBox_状态.Margin = new Padding(4, 5, 4, 5);
            uiTextBox_状态.MinimumSize = new Size(1, 16);
            uiTextBox_状态.Multiline = true;
            uiTextBox_状态.Name = "uiTextBox_状态";
            uiTextBox_状态.Padding = new Padding(5);
            uiTextBox_状态.ShowText = false;
            uiTextBox_状态.Size = new Size(257, 94);
            uiTextBox_状态.TabIndex = 17;
            uiTextBox_状态.TextAlignment = ContentAlignment.MiddleLeft;
            uiTextBox_状态.Watermark = "";
            // 
            // uiButton2
            // 
            uiButton2.Font = new Font("宋体", 12F, FontStyle.Regular, GraphicsUnit.Point, 134);
            uiButton2.Location = new Point(265, 176);
            uiButton2.MinimumSize = new Size(1, 1);
            uiButton2.Name = "uiButton2";
            uiButton2.Size = new Size(42, 26);
            uiButton2.TabIndex = 19;
            uiButton2.Text = "AI";
            uiButton2.TipsFont = new Font("宋体", 9F, FontStyle.Regular, GraphicsUnit.Point, 134);
            uiButton2.Click += uiButton2_Click;
            // 
            // uiCheckBox_RU客户
            // 
            uiCheckBox_RU客户.Font = new Font("宋体", 12F, FontStyle.Regular, GraphicsUnit.Point, 134);
            uiCheckBox_RU客户.ForeColor = Color.FromArgb(48, 48, 48);
            uiCheckBox_RU客户.Location = new Point(362, 115);
            uiCheckBox_RU客户.MinimumSize = new Size(1, 1);
            uiCheckBox_RU客户.Name = "uiCheckBox_RU客户";
            uiCheckBox_RU客户.Size = new Size(81, 28);
            uiCheckBox_RU客户.TabIndex = 20;
            uiCheckBox_RU客户.Text = "RU客户";
            uiCheckBox_RU客户.CheckedChanged += uiCheckBox_RU客户_CheckedChanged;
            // 
            // uiButton3
            // 
            uiButton3.Font = new Font("宋体", 12F, FontStyle.Regular, GraphicsUnit.Point, 134);
            uiButton3.Location = new Point(268, 142);
            uiButton3.MinimumSize = new Size(1, 1);
            uiButton3.Name = "uiButton3";
            uiButton3.Size = new Size(71, 26);
            uiButton3.TabIndex = 21;
            uiButton3.Text = "处理-V2";
            uiButton3.TipsFont = new Font("宋体", 9F, FontStyle.Regular, GraphicsUnit.Point, 134);
            uiButton3.Click += uiButton3_Click;
            // 
            // Form1
            // 
            AutoScaleMode = AutoScaleMode.None;
            BackgroundImageLayout = ImageLayout.None;
            ClientSize = new Size(346, 214);
            Controls.Add(uiButton3);
            Controls.Add(uiCheckBox_RU客户);
            Controls.Add(uiButton2);
            Controls.Add(uiButton1);
            Controls.Add(uiTextBox_状态);
            Controls.Add(uiTextBox_附件地址);
            Controls.Add(uiTextBox_订单地址);
            Controls.Add(button_开始处理);
            Controls.Add(button_附件导入);
            Controls.Add(button_订单导入);
            Name = "Form1";
            Text = "包装资料生成(C#版)";
            ZoomScaleRect = new Rectangle(15, 15, 433, 220);
            ResumeLayout(false);
        }

        #endregion
        private TextBox textBox_订单地址;
        private Button button_组合;
        private TextBox textBox_状态;
        private TextBox textBox_附件地址;
        private Button button_附件导入1;
        private Button button1;
        private Sunny.UI.UIStyleManager uiStyleManager1;
        private Sunny.UI.UITextBox uiTextBox1;
        private Sunny.UI.UIButton button_订单导入;
        private Sunny.UI.UITextBox uiTextBox2;
        private Sunny.UI.UIButton button_附件导入;
        private Sunny.UI.UITextBox uiTextBox3;
        private Sunny.UI.UIButton button_开始处理;
        private Sunny.UI.UITextBox uiTextBox_订单地址;
        private Sunny.UI.UITextBox uiTextBox_附件地址;
        private Sunny.UI.UITextBox uiTextBox_状态;
        private Sunny.UI.UIButton uiButton1;
        private Sunny.UI.UIButton uiButton2;
        private Sunny.UI.UICheckBox uiCheckBox_RU客户;
        private Sunny.UI.UIButton uiButton3;
    }
}
