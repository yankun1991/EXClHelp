namespace ExclHelp
{
    partial class Form1
    {
        /// <summary>
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 清理所有正在使用的资源。
        /// </summary>
        /// <param name="disposing">如果应释放托管资源，为 true；否则为 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows 窗体设计器生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.btn_FileChoseOne = new System.Windows.Forms.Button();
            this.txt_FileChoseone = new System.Windows.Forms.TextBox();
            this.txt_FileChoseTwo = new System.Windows.Forms.TextBox();
            this.btn_FileChoseTwo = new System.Windows.Forms.Button();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.button3 = new System.Windows.Forms.Button();
            this.cb_sheetNameTwo = new System.Windows.Forms.ComboBox();
            this.cb_sheetNameOne = new System.Windows.Forms.ComboBox();
            this.txt_indexTwo = new System.Windows.Forms.TextBox();
            this.txt_indexOne = new System.Windows.Forms.TextBox();
            this.button2 = new System.Windows.Forms.Button();
            this.button1 = new System.Windows.Forms.Button();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.pb = new System.Windows.Forms.ProgressBar();
            this.dt_result = new System.Windows.Forms.DataGridView();
            this.imageList1 = new System.Windows.Forms.ImageList(this.components);
            this.button4 = new System.Windows.Forms.Button();
            this.groupBox1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dt_result)).BeginInit();
            this.SuspendLayout();
            // 
            // btn_FileChoseOne
            // 
            this.btn_FileChoseOne.ForeColor = System.Drawing.SystemColors.ActiveCaptionText;
            this.btn_FileChoseOne.Location = new System.Drawing.Point(251, 25);
            this.btn_FileChoseOne.Name = "btn_FileChoseOne";
            this.btn_FileChoseOne.Size = new System.Drawing.Size(70, 20);
            this.btn_FileChoseOne.TabIndex = 0;
            this.btn_FileChoseOne.Text = "文件选择";
            this.btn_FileChoseOne.UseVisualStyleBackColor = true;
            this.btn_FileChoseOne.Click += new System.EventHandler(this.btn_FileChoseOne_Click);
            // 
            // txt_FileChoseone
            // 
            this.txt_FileChoseone.Location = new System.Drawing.Point(91, 23);
            this.txt_FileChoseone.Name = "txt_FileChoseone";
            this.txt_FileChoseone.Size = new System.Drawing.Size(154, 21);
            this.txt_FileChoseone.TabIndex = 1;
            // 
            // txt_FileChoseTwo
            // 
            this.txt_FileChoseTwo.Location = new System.Drawing.Point(91, 76);
            this.txt_FileChoseTwo.Name = "txt_FileChoseTwo";
            this.txt_FileChoseTwo.Size = new System.Drawing.Size(154, 21);
            this.txt_FileChoseTwo.TabIndex = 3;
            // 
            // btn_FileChoseTwo
            // 
            this.btn_FileChoseTwo.ForeColor = System.Drawing.SystemColors.ActiveCaptionText;
            this.btn_FileChoseTwo.Location = new System.Drawing.Point(251, 77);
            this.btn_FileChoseTwo.Name = "btn_FileChoseTwo";
            this.btn_FileChoseTwo.Size = new System.Drawing.Size(70, 20);
            this.btn_FileChoseTwo.TabIndex = 2;
            this.btn_FileChoseTwo.Text = "文件选择";
            this.btn_FileChoseTwo.UseVisualStyleBackColor = true;
            this.btn_FileChoseTwo.Click += new System.EventHandler(this.btn_FileChoseTwo_Click);
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.button4);
            this.groupBox1.Controls.Add(this.button3);
            this.groupBox1.Controls.Add(this.cb_sheetNameTwo);
            this.groupBox1.Controls.Add(this.cb_sheetNameOne);
            this.groupBox1.Controls.Add(this.txt_indexTwo);
            this.groupBox1.Controls.Add(this.txt_indexOne);
            this.groupBox1.Controls.Add(this.button2);
            this.groupBox1.Controls.Add(this.button1);
            this.groupBox1.Controls.Add(this.btn_FileChoseOne);
            this.groupBox1.Controls.Add(this.txt_FileChoseTwo);
            this.groupBox1.Controls.Add(this.txt_FileChoseone);
            this.groupBox1.Controls.Add(this.btn_FileChoseTwo);
            this.groupBox1.Location = new System.Drawing.Point(48, 22);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(1209, 131);
            this.groupBox1.TabIndex = 4;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "文件选择";
            // 
            // button3
            // 
            this.button3.ForeColor = System.Drawing.SystemColors.ActiveCaptionText;
            this.button3.Location = new System.Drawing.Point(824, 50);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(75, 23);
            this.button3.TabIndex = 11;
            this.button3.Text = "筛选";
            this.button3.UseVisualStyleBackColor = true;
            this.button3.Click += new System.EventHandler(this.button3_Click);
            // 
            // cb_sheetNameTwo
            // 
            this.cb_sheetNameTwo.FormattingEnabled = true;
            this.cb_sheetNameTwo.Location = new System.Drawing.Point(389, 78);
            this.cb_sheetNameTwo.Name = "cb_sheetNameTwo";
            this.cb_sheetNameTwo.Size = new System.Drawing.Size(176, 20);
            this.cb_sheetNameTwo.TabIndex = 10;
            this.cb_sheetNameTwo.SelectedIndexChanged += new System.EventHandler(this.cb_sheetNameTwo_SelectedIndexChanged);
            // 
            // cb_sheetNameOne
            // 
            this.cb_sheetNameOne.FormattingEnabled = true;
            this.cb_sheetNameOne.Location = new System.Drawing.Point(389, 26);
            this.cb_sheetNameOne.Name = "cb_sheetNameOne";
            this.cb_sheetNameOne.Size = new System.Drawing.Size(176, 20);
            this.cb_sheetNameOne.TabIndex = 9;
            this.cb_sheetNameOne.SelectedIndexChanged += new System.EventHandler(this.cb_sheetNameOne_SelectedIndexChanged);
            // 
            // txt_indexTwo
            // 
            this.txt_indexTwo.Location = new System.Drawing.Point(327, 77);
            this.txt_indexTwo.Name = "txt_indexTwo";
            this.txt_indexTwo.Size = new System.Drawing.Size(56, 21);
            this.txt_indexTwo.TabIndex = 7;
            // 
            // txt_indexOne
            // 
            this.txt_indexOne.Location = new System.Drawing.Point(327, 25);
            this.txt_indexOne.Name = "txt_indexOne";
            this.txt_indexOne.Size = new System.Drawing.Size(56, 21);
            this.txt_indexOne.TabIndex = 6;
            // 
            // button2
            // 
            this.button2.ForeColor = System.Drawing.SystemColors.ActiveCaptionText;
            this.button2.Location = new System.Drawing.Point(703, 50);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(75, 23);
            this.button2.TabIndex = 5;
            this.button2.Text = "导出";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // button1
            // 
            this.button1.ForeColor = System.Drawing.SystemColors.ActiveCaptionText;
            this.button1.Location = new System.Drawing.Point(594, 50);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(75, 23);
            this.button1.TabIndex = 4;
            this.button1.Text = "合并";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.pb);
            this.groupBox2.Controls.Add(this.dt_result);
            this.groupBox2.Location = new System.Drawing.Point(48, 159);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(1199, 483);
            this.groupBox2.TabIndex = 5;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "结果";
            // 
            // pb
            // 
            this.pb.BackColor = System.Drawing.SystemColors.ActiveCaptionText;
            this.pb.ForeColor = System.Drawing.SystemColors.GradientActiveCaption;
            this.pb.Location = new System.Drawing.Point(306, 194);
            this.pb.Name = "pb";
            this.pb.Size = new System.Drawing.Size(646, 28);
            this.pb.TabIndex = 12;
            this.pb.Visible = false;
            // 
            // dt_result
            // 
            this.dt_result.BackgroundColor = System.Drawing.SystemColors.ControlText;
            this.dt_result.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dt_result.GridColor = System.Drawing.SystemColors.ControlText;
            this.dt_result.Location = new System.Drawing.Point(22, 20);
            this.dt_result.Name = "dt_result";
            this.dt_result.RowTemplate.Height = 23;
            this.dt_result.Size = new System.Drawing.Size(1156, 433);
            this.dt_result.TabIndex = 0;
            // 
            // imageList1
            // 
            this.imageList1.ColorDepth = System.Windows.Forms.ColorDepth.Depth8Bit;
            this.imageList1.ImageSize = new System.Drawing.Size(16, 16);
            this.imageList1.TransparentColor = System.Drawing.Color.Transparent;
            // 
            // button4
            // 
            this.button4.Location = new System.Drawing.Point(954, 49);
            this.button4.Name = "button4";
            this.button4.Size = new System.Drawing.Size(75, 23);
            this.button4.TabIndex = 12;
            this.button4.Text = "平均时间";
            this.button4.UseVisualStyleBackColor = true;
            this.button4.Click += new System.EventHandler(this.button4_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.PowderBlue;
            this.ClientSize = new System.Drawing.Size(1277, 699);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.groupBox1);
            this.ForeColor = System.Drawing.SystemColors.ControlText;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "Form1";
            this.Text = "Excl合并";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dt_result)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button btn_FileChoseOne;
        private System.Windows.Forms.TextBox txt_FileChoseone;
        private System.Windows.Forms.TextBox txt_FileChoseTwo;
        private System.Windows.Forms.Button btn_FileChoseTwo;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.DataGridView dt_result;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.TextBox txt_indexTwo;
        private System.Windows.Forms.TextBox txt_indexOne;
        private System.Windows.Forms.ComboBox cb_sheetNameOne;
        private System.Windows.Forms.ComboBox cb_sheetNameTwo;
        private System.Windows.Forms.ImageList imageList1;
        private System.Windows.Forms.ProgressBar pb;
        private System.Windows.Forms.Button button3;
        private System.Windows.Forms.Button button4;

    }
}

