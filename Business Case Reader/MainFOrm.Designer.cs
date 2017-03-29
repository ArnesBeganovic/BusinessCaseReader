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
            this.btn_browseBC = new System.Windows.Forms.Button();
            this.pic_Loading = new System.Windows.Forms.PictureBox();
            this.tb_FileName = new System.Windows.Forms.TextBox();
            this.tabPage2 = new System.Windows.Forms.TabPage();
            this.lbl_Info = new System.Windows.Forms.Label();
            this.bth_ReadSheet = new System.Windows.Forms.Button();
            this.lb_SheetNames = new System.Windows.Forms.ListBox();
            this.tabPage3 = new System.Windows.Forms.TabPage();
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
            this.tabovi.Location = new System.Drawing.Point(-2, 138);
            this.tabovi.Name = "tabovi";
            this.tabovi.SelectedIndex = 0;
            this.tabovi.Size = new System.Drawing.Size(637, 416);
            this.tabovi.TabIndex = 10;
            // 
            // tabPage1
            // 
            this.tabPage1.BackgroundImage = global::Business_Case_Reader.Properties.Resources.The_Right_Products_Hero;
            this.tabPage1.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.tabPage1.Controls.Add(this.btn_browseBC);
            this.tabPage1.Controls.Add(this.pic_Loading);
            this.tabPage1.Controls.Add(this.tb_FileName);
            this.tabPage1.Location = new System.Drawing.Point(4, 22);
            this.tabPage1.Name = "tabPage1";
            this.tabPage1.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage1.Size = new System.Drawing.Size(629, 390);
            this.tabPage1.TabIndex = 0;
            this.tabPage1.Text = "Browse File";
            this.tabPage1.UseVisualStyleBackColor = true;
            // 
            // btn_browseBC
            // 
            this.btn_browseBC.Location = new System.Drawing.Point(481, 160);
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
            this.pic_Loading.Location = new System.Drawing.Point(143, 29);
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
            this.tb_FileName.Location = new System.Drawing.Point(68, 160);
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
            this.tabPage2.Size = new System.Drawing.Size(629, 390);
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
            this.bth_ReadSheet.Location = new System.Drawing.Point(468, 60);
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
            this.lb_SheetNames.Size = new System.Drawing.Size(165, 303);
            this.lb_SheetNames.TabIndex = 4;
            this.lb_SheetNames.SelectedIndexChanged += new System.EventHandler(this.lb_SheetNames_SelectedIndexChanged);
            // 
            // tabPage3
            // 
            this.tabPage3.BackgroundImage = global::Business_Case_Reader.Properties.Resources.bck_Vozilo;
            this.tabPage3.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.tabPage3.Controls.Add(this.btn_ExportCSV);
            this.tabPage3.Controls.Add(this.dataGridView);
            this.tabPage3.Location = new System.Drawing.Point(4, 22);
            this.tabPage3.Name = "tabPage3";
            this.tabPage3.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage3.Size = new System.Drawing.Size(629, 390);
            this.tabPage3.TabIndex = 2;
            this.tabPage3.Text = "Control Data";
            this.tabPage3.UseVisualStyleBackColor = true;
            // 
            // btn_ExportCSV
            // 
            this.btn_ExportCSV.Location = new System.Drawing.Point(454, 322);
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
            this.dataGridView.Location = new System.Drawing.Point(10, 22);
            this.dataGridView.Name = "dataGridView";
            this.dataGridView.Size = new System.Drawing.Size(613, 273);
            this.dataGridView.TabIndex = 0;
            // 
            // pictureBox3
            // 
            this.pictureBox3.Image = global::Business_Case_Reader.Properties.Resources.businessCaseReader;
            this.pictureBox3.Location = new System.Drawing.Point(256, 25);
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
            this.pictureBox2.Size = new System.Drawing.Size(637, 40);
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
            this.ClientSize = new System.Drawing.Size(632, 548);
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
    }
}

