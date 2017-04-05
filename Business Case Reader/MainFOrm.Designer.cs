namespace Business_Case_Reader
{
    partial class MainForm
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
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
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.openFileDialog = new System.Windows.Forms.OpenFileDialog();
            this.tabovi = new System.Windows.Forms.TabControl();
            this.tabPage1 = new System.Windows.Forms.TabPage();
            this.cb_ProgramCategory = new System.Windows.Forms.ComboBox();
            this.cb_Market = new System.Windows.Forms.ComboBox();
            this.cb_ProductGroup = new System.Windows.Forms.ComboBox();
            this.cb_SDTPrimaryLocation = new System.Windows.Forms.ComboBox();
            this.cb_OEMAccount = new System.Windows.Forms.ComboBox();
            this.label6 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.btn_browseBC = new System.Windows.Forms.Button();
            this.pic_Loading = new System.Windows.Forms.PictureBox();
            this.tb_FileName = new System.Windows.Forms.TextBox();
            this.tabPage2 = new System.Windows.Forms.TabPage();
            this.lbl_Info = new System.Windows.Forms.Label();
            this.bth_ReadSheet = new System.Windows.Forms.Button();
            this.lb_SheetNames = new System.Windows.Forms.ListBox();
            this.tabPage3 = new System.Windows.Forms.TabPage();
            this.label1 = new System.Windows.Forms.Label();
            this.btn_ExportCSV = new System.Windows.Forms.Button();
            this.dataGridView = new System.Windows.Forms.DataGridView();
            this.pictureBox3 = new System.Windows.Forms.PictureBox();
            this.pictureBox2 = new System.Windows.Forms.PictureBox();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.tabovi.SuspendLayout();
            this.tabPage1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pic_Loading)).BeginInit();
            this.tabPage2.SuspendLayout();
            this.tabPage3.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox3)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox2)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.SuspendLayout();
            // 
            // openFileDialog
            // 
            this.openFileDialog.DefaultExt = "xlsm";
            this.openFileDialog.Filter = "Excel Files(.xlsm)|*.xlsm";
            this.openFileDialog.Title = "Business Case File Browser";
            // 
            // tabovi
            // 
            this.tabovi.Controls.Add(this.tabPage1);
            this.tabovi.Controls.Add(this.tabPage2);
            this.tabovi.Controls.Add(this.tabPage3);
            this.tabovi.Location = new System.Drawing.Point(12, 138);
            this.tabovi.Name = "tabovi";
            this.tabovi.SelectedIndex = 0;
            this.tabovi.Size = new System.Drawing.Size(952, 508);
            this.tabovi.TabIndex = 10;
            // 
            // tabPage1
            // 
            this.tabPage1.BackgroundImage = global::Business_Case_Reader.Properties.Resources.The_Right_Products_Hero;
            this.tabPage1.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.tabPage1.Controls.Add(this.cb_ProgramCategory);
            this.tabPage1.Controls.Add(this.cb_Market);
            this.tabPage1.Controls.Add(this.cb_ProductGroup);
            this.tabPage1.Controls.Add(this.cb_SDTPrimaryLocation);
            this.tabPage1.Controls.Add(this.cb_OEMAccount);
            this.tabPage1.Controls.Add(this.label6);
            this.tabPage1.Controls.Add(this.label5);
            this.tabPage1.Controls.Add(this.label4);
            this.tabPage1.Controls.Add(this.label3);
            this.tabPage1.Controls.Add(this.label2);
            this.tabPage1.Controls.Add(this.btn_browseBC);
            this.tabPage1.Controls.Add(this.pic_Loading);
            this.tabPage1.Controls.Add(this.tb_FileName);
            this.tabPage1.Location = new System.Drawing.Point(4, 22);
            this.tabPage1.Name = "tabPage1";
            this.tabPage1.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage1.Size = new System.Drawing.Size(944, 482);
            this.tabPage1.TabIndex = 0;
            this.tabPage1.Text = "Browse File";
            this.tabPage1.UseVisualStyleBackColor = true;
            // 
            // cb_ProgramCategory
            // 
            this.cb_ProgramCategory.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cb_ProgramCategory.Font = new System.Drawing.Font("Microsoft Sans Serif", 15.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.cb_ProgramCategory.FormattingEnabled = true;
            this.cb_ProgramCategory.Items.AddRange(new object[] {
            "Select Item",
            "Iron",
            "Bronze",
            "Silver",
            "Gold",
            "Platinum"});
            this.cb_ProgramCategory.Location = new System.Drawing.Point(763, 349);
            this.cb_ProgramCategory.Name = "cb_ProgramCategory";
            this.cb_ProgramCategory.Size = new System.Drawing.Size(175, 33);
            this.cb_ProgramCategory.TabIndex = 18;
            // 
            // cb_Market
            // 
            this.cb_Market.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cb_Market.Font = new System.Drawing.Font("Microsoft Sans Serif", 15.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.cb_Market.FormattingEnabled = true;
            this.cb_Market.Items.AddRange(new object[] {
            "Select Item",
            "Automotive",
            "Industrial",
            "Recreational",
            "Military",
            "Other"});
            this.cb_Market.Location = new System.Drawing.Point(763, 274);
            this.cb_Market.Name = "cb_Market";
            this.cb_Market.Size = new System.Drawing.Size(175, 33);
            this.cb_Market.TabIndex = 17;
            // 
            // cb_ProductGroup
            // 
            this.cb_ProductGroup.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cb_ProductGroup.Font = new System.Drawing.Font("Microsoft Sans Serif", 15.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.cb_ProductGroup.FormattingEnabled = true;
            this.cb_ProductGroup.Items.AddRange(new object[] {
            "Select Item",
            "China",
            "Fabrics",
            "Foam",
            "Full Value Chain",
            "JIT",
            "Metals and Mechanism",
            "Recaro",
            "Trim",
            "OHS"});
            this.cb_ProductGroup.Location = new System.Drawing.Point(763, 199);
            this.cb_ProductGroup.Name = "cb_ProductGroup";
            this.cb_ProductGroup.Size = new System.Drawing.Size(175, 33);
            this.cb_ProductGroup.TabIndex = 16;
            // 
            // cb_SDTPrimaryLocation
            // 
            this.cb_SDTPrimaryLocation.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cb_SDTPrimaryLocation.Font = new System.Drawing.Font("Microsoft Sans Serif", 15.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.cb_SDTPrimaryLocation.FormattingEnabled = true;
            this.cb_SDTPrimaryLocation.Items.AddRange(new object[] {
            "Select Item",
            "Africa",
            "Asia",
            "Australia",
            "EUrope",
            "North America",
            "South America"});
            this.cb_SDTPrimaryLocation.Location = new System.Drawing.Point(763, 124);
            this.cb_SDTPrimaryLocation.Name = "cb_SDTPrimaryLocation";
            this.cb_SDTPrimaryLocation.Size = new System.Drawing.Size(175, 33);
            this.cb_SDTPrimaryLocation.TabIndex = 15;
            // 
            // cb_OEMAccount
            // 
            this.cb_OEMAccount.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cb_OEMAccount.Font = new System.Drawing.Font("Microsoft Sans Serif", 15.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.cb_OEMAccount.FormattingEnabled = true;
            this.cb_OEMAccount.Items.AddRange(new object[] {
            "Select Item",
            "Volvo",
            "Ford",
            "JLR",
            "TBD"});
            this.cb_OEMAccount.Location = new System.Drawing.Point(763, 43);
            this.cb_OEMAccount.Name = "cb_OEMAccount";
            this.cb_OEMAccount.Size = new System.Drawing.Size(175, 33);
            this.cb_OEMAccount.TabIndex = 14;
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Font = new System.Drawing.Font("Microsoft Sans Serif", 15.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label6.ForeColor = System.Drawing.Color.White;
            this.label6.Location = new System.Drawing.Point(524, 351);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(202, 25);
            this.label6.TabIndex = 13;
            this.label6.Text = "Program Category";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Font = new System.Drawing.Font("Microsoft Sans Serif", 15.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label5.ForeColor = System.Drawing.Color.White;
            this.label5.Location = new System.Drawing.Point(524, 276);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(84, 25);
            this.label5.TabIndex = 12;
            this.label5.Text = "Market";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("Microsoft Sans Serif", 15.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label4.ForeColor = System.Drawing.Color.White;
            this.label4.Location = new System.Drawing.Point(524, 201);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(160, 25);
            this.label4.TabIndex = 11;
            this.label4.Text = "Product group";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 15.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.ForeColor = System.Drawing.Color.White;
            this.label3.Location = new System.Drawing.Point(524, 126);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(234, 25);
            this.label3.TabIndex = 10;
            this.label3.Text = "SDT Primary location";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 15.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.ForeColor = System.Drawing.Color.White;
            this.label2.Location = new System.Drawing.Point(524, 51);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(155, 25);
            this.label2.TabIndex = 9;
            this.label2.Text = "OEM Account";
            // 
            // btn_browseBC
            // 
            this.btn_browseBC.Location = new System.Drawing.Point(808, 426);
            this.btn_browseBC.Name = "btn_browseBC";
            this.btn_browseBC.Size = new System.Drawing.Size(131, 32);
            this.btn_browseBC.TabIndex = 3;
            this.btn_browseBC.Text = "Browse Business Case";
            this.btn_browseBC.UseVisualStyleBackColor = true;
            this.btn_browseBC.Click += new System.EventHandler(this.btn_browseBC_Click);
            // 
            // pic_Loading
            // 
            this.pic_Loading.Image = global::Business_Case_Reader.Properties.Resources.loading;
            this.pic_Loading.Location = new System.Drawing.Point(28, 164);
            this.pic_Loading.Name = "pic_Loading";
            this.pic_Loading.Size = new System.Drawing.Size(334, 104);
            this.pic_Loading.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.pic_Loading.TabIndex = 7;
            this.pic_Loading.TabStop = false;
            this.pic_Loading.Visible = false;
            // 
            // tb_FileName
            // 
            this.tb_FileName.Enabled = false;
            this.tb_FileName.Font = new System.Drawing.Font("Microsoft Sans Serif", 15F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.tb_FileName.Location = new System.Drawing.Point(410, 428);
            this.tb_FileName.Name = "tb_FileName";
            this.tb_FileName.Size = new System.Drawing.Size(392, 30);
            this.tb_FileName.TabIndex = 8;
            this.tb_FileName.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // tabPage2
            // 
            this.tabPage2.BackgroundImage = global::Business_Case_Reader.Properties.Resources.The_Right_Advancements_Hero;
            this.tabPage2.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.tabPage2.Controls.Add(this.lbl_Info);
            this.tabPage2.Controls.Add(this.bth_ReadSheet);
            this.tabPage2.Controls.Add(this.lb_SheetNames);
            this.tabPage2.Location = new System.Drawing.Point(4, 22);
            this.tabPage2.Name = "tabPage2";
            this.tabPage2.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage2.Size = new System.Drawing.Size(944, 482);
            this.tabPage2.TabIndex = 1;
            this.tabPage2.Text = "Available Sheets";
            this.tabPage2.UseVisualStyleBackColor = true;
            // 
            // lbl_Info
            // 
            this.lbl_Info.AutoSize = true;
            this.lbl_Info.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lbl_Info.ForeColor = System.Drawing.Color.White;
            this.lbl_Info.Location = new System.Drawing.Point(210, 131);
            this.lbl_Info.Name = "lbl_Info";
            this.lbl_Info.Size = new System.Drawing.Size(0, 16);
            this.lbl_Info.TabIndex = 11;
            // 
            // bth_ReadSheet
            // 
            this.bth_ReadSheet.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.bth_ReadSheet.Location = new System.Drawing.Point(794, 445);
            this.bth_ReadSheet.Name = "bth_ReadSheet";
            this.bth_ReadSheet.Size = new System.Drawing.Size(144, 31);
            this.bth_ReadSheet.TabIndex = 10;
            this.bth_ReadSheet.Text = "Read Sheet";
            this.bth_ReadSheet.UseVisualStyleBackColor = true;
            this.bth_ReadSheet.Click += new System.EventHandler(this.bth_ReadSheet_Click);
            // 
            // lb_SheetNames
            // 
            this.lb_SheetNames.FormattingEnabled = true;
            this.lb_SheetNames.Location = new System.Drawing.Point(27, 60);
            this.lb_SheetNames.Name = "lb_SheetNames";
            this.lb_SheetNames.Size = new System.Drawing.Size(165, 407);
            this.lb_SheetNames.TabIndex = 4;
            this.lb_SheetNames.SelectedIndexChanged += new System.EventHandler(this.lb_SheetNames_SelectedIndexChanged);
            // 
            // tabPage3
            // 
            this.tabPage3.BackgroundImage = global::Business_Case_Reader.Properties.Resources.bck_Vozilo;
            this.tabPage3.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.tabPage3.Controls.Add(this.label1);
            this.tabPage3.Controls.Add(this.btn_ExportCSV);
            this.tabPage3.Controls.Add(this.dataGridView);
            this.tabPage3.Location = new System.Drawing.Point(4, 22);
            this.tabPage3.Name = "tabPage3";
            this.tabPage3.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage3.Size = new System.Drawing.Size(944, 482);
            this.tabPage3.TabIndex = 2;
            this.tabPage3.Text = "Control Data";
            this.tabPage3.UseVisualStyleBackColor = true;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.ForeColor = System.Drawing.Color.White;
            this.label1.Location = new System.Drawing.Point(315, 23);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(287, 24);
            this.label1.TabIndex = 2;
            this.label1.Text = "Control data before Submiting";
            // 
            // btn_ExportCSV
            // 
            this.btn_ExportCSV.Location = new System.Drawing.Point(780, 443);
            this.btn_ExportCSV.Name = "btn_ExportCSV";
            this.btn_ExportCSV.Size = new System.Drawing.Size(158, 33);
            this.btn_ExportCSV.TabIndex = 1;
            this.btn_ExportCSV.Text = "Submit";
            this.btn_ExportCSV.UseVisualStyleBackColor = true;
            this.btn_ExportCSV.Click += new System.EventHandler(this.btn_ExportCSV_Click);
            // 
            // dataGridView
            // 
            this.dataGridView.BackgroundColor = System.Drawing.SystemColors.GrayText;
            this.dataGridView.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView.Location = new System.Drawing.Point(10, 64);
            this.dataGridView.Name = "dataGridView";
            this.dataGridView.Size = new System.Drawing.Size(928, 373);
            this.dataGridView.TabIndex = 0;
            // 
            // pictureBox3
            // 
            this.pictureBox3.Image = global::Business_Case_Reader.Properties.Resources.businessCaseReader;
            this.pictureBox3.Location = new System.Drawing.Point(606, 22);
            this.pictureBox3.Name = "pictureBox3";
            this.pictureBox3.Size = new System.Drawing.Size(358, 50);
            this.pictureBox3.TabIndex = 2;
            this.pictureBox3.TabStop = false;
            // 
            // pictureBox2
            // 
            this.pictureBox2.Image = global::Business_Case_Reader.Properties.Resources.baner_blue;
            this.pictureBox2.Location = new System.Drawing.Point(-2, 92);
            this.pictureBox2.Name = "pictureBox2";
            this.pictureBox2.Size = new System.Drawing.Size(980, 40);
            this.pictureBox2.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.pictureBox2.TabIndex = 1;
            this.pictureBox2.TabStop = false;
            // 
            // pictureBox1
            // 
            this.pictureBox1.Image = global::Business_Case_Reader.Properties.Resources.logo;
            this.pictureBox1.Location = new System.Drawing.Point(12, 12);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(182, 74);
            this.pictureBox1.TabIndex = 0;
            this.pictureBox1.TabStop = false;
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.Window;
            this.ClientSize = new System.Drawing.Size(976, 658);
            this.Controls.Add(this.tabovi);
            this.Controls.Add(this.pictureBox3);
            this.Controls.Add(this.pictureBox2);
            this.Controls.Add(this.pictureBox1);
            this.Name = "MainForm";
            this.Text = "Business Case Reader";
            this.tabovi.ResumeLayout(false);
            this.tabPage1.ResumeLayout(false);
            this.tabPage1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pic_Loading)).EndInit();
            this.tabPage2.ResumeLayout(false);
            this.tabPage2.PerformLayout();
            this.tabPage3.ResumeLayout(false);
            this.tabPage3.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox3)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox2)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.PictureBox pictureBox1;
        private System.Windows.Forms.PictureBox pictureBox2;
        private System.Windows.Forms.PictureBox pictureBox3;
        private System.Windows.Forms.OpenFileDialog openFileDialog;
        private System.Windows.Forms.Button btn_browseBC;
        private System.Windows.Forms.ListBox lb_SheetNames;
        private System.Windows.Forms.PictureBox pic_Loading;
        private System.Windows.Forms.TextBox tb_FileName;
        private System.Windows.Forms.TabControl tabovi;
        private System.Windows.Forms.TabPage tabPage1;
        private System.Windows.Forms.TabPage tabPage2;
        private System.Windows.Forms.Button bth_ReadSheet;
        private System.Windows.Forms.TabPage tabPage3;
        private System.Windows.Forms.Button btn_ExportCSV;
        private System.Windows.Forms.DataGridView dataGridView;
        private System.Windows.Forms.Label lbl_Info;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.ComboBox cb_ProgramCategory;
        private System.Windows.Forms.ComboBox cb_Market;
        private System.Windows.Forms.ComboBox cb_ProductGroup;
        private System.Windows.Forms.ComboBox cb_SDTPrimaryLocation;
        private System.Windows.Forms.ComboBox cb_OEMAccount;
    }
}

