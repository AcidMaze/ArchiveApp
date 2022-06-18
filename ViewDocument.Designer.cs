namespace ArchiveApp
{
    partial class ViewDocument
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(ViewDocument));
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.panel1 = new System.Windows.Forms.Panel();
            this.pageNum = new System.Windows.Forms.TextBox();
            this.btn_back = new System.Windows.Forms.Button();
            this.btn_next = new System.Windows.Forms.Button();
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.файлToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.эскпортироватьДокументКакToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.docToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.pNGToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.jPEGToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.экспортироватьСтраницуКакToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.dOCToolStripMenuItem1 = new System.Windows.Forms.ToolStripMenuItem();
            this.pNGToolStripMenuItem1 = new System.Windows.Forms.ToolStripMenuItem();
            this.jPGToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.saveFileDialog1 = new System.Windows.Forms.SaveFileDialog();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.panel1.SuspendLayout();
            this.menuStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // pictureBox1
            // 
            this.pictureBox1.BackColor = System.Drawing.Color.White;
            this.pictureBox1.Location = new System.Drawing.Point(12, 26);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(521, 590);
            this.pictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.pictureBox1.TabIndex = 0;
            this.pictureBox1.TabStop = false;
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.pageNum);
            this.panel1.Controls.Add(this.btn_back);
            this.panel1.Controls.Add(this.btn_next);
            this.panel1.Location = new System.Drawing.Point(129, 625);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(287, 48);
            this.panel1.TabIndex = 1;
            // 
            // pageNum
            // 
            this.pageNum.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.pageNum.Location = new System.Drawing.Point(119, 13);
            this.pageNum.Name = "pageNum";
            this.pageNum.Size = new System.Drawing.Size(48, 26);
            this.pageNum.TabIndex = 3;
            this.pageNum.Text = "1";
            this.pageNum.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.pageNum.TextChanged += new System.EventHandler(this.pageNum_TextChanged);
            this.pageNum.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.pageNum_KeyPress);
            // 
            // btn_back
            // 
            this.btn_back.Location = new System.Drawing.Point(16, 11);
            this.btn_back.Name = "btn_back";
            this.btn_back.Size = new System.Drawing.Size(75, 28);
            this.btn_back.TabIndex = 2;
            this.btn_back.Text = "Назад";
            this.btn_back.UseVisualStyleBackColor = true;
            this.btn_back.Click += new System.EventHandler(this.btn_back_Click);
            // 
            // btn_next
            // 
            this.btn_next.Location = new System.Drawing.Point(191, 11);
            this.btn_next.Name = "btn_next";
            this.btn_next.Size = new System.Drawing.Size(75, 26);
            this.btn_next.TabIndex = 1;
            this.btn_next.Text = "Впередёт";
            this.btn_next.UseVisualStyleBackColor = true;
            this.btn_next.Click += new System.EventHandler(this.btn_next_Click);
            // 
            // menuStrip1
            // 
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.файлToolStripMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(545, 24);
            this.menuStrip1.TabIndex = 2;
            this.menuStrip1.Text = "menuStrip1";
            // 
            // файлToolStripMenuItem
            // 
            this.файлToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.эскпортироватьДокументКакToolStripMenuItem,
            this.экспортироватьСтраницуКакToolStripMenuItem});
            this.файлToolStripMenuItem.Name = "файлToolStripMenuItem";
            this.файлToolStripMenuItem.Size = new System.Drawing.Size(48, 20);
            this.файлToolStripMenuItem.Text = "Файл";
            // 
            // эскпортироватьДокументКакToolStripMenuItem
            // 
            this.эскпортироватьДокументКакToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.docToolStripMenuItem,
            this.pNGToolStripMenuItem,
            this.jPEGToolStripMenuItem});
            this.эскпортироватьДокументКакToolStripMenuItem.Name = "эскпортироватьДокументКакToolStripMenuItem";
            this.эскпортироватьДокументКакToolStripMenuItem.Size = new System.Drawing.Size(242, 22);
            this.эскпортироватьДокументКакToolStripMenuItem.Text = "Эскпортировать документ  как";
            // 
            // docToolStripMenuItem
            // 
            this.docToolStripMenuItem.Name = "docToolStripMenuItem";
            this.docToolStripMenuItem.Size = new System.Drawing.Size(99, 22);
            this.docToolStripMenuItem.Text = "DOC";
            // 
            // pNGToolStripMenuItem
            // 
            this.pNGToolStripMenuItem.Name = "pNGToolStripMenuItem";
            this.pNGToolStripMenuItem.Size = new System.Drawing.Size(99, 22);
            this.pNGToolStripMenuItem.Text = "PNG";
            // 
            // jPEGToolStripMenuItem
            // 
            this.jPEGToolStripMenuItem.Name = "jPEGToolStripMenuItem";
            this.jPEGToolStripMenuItem.Size = new System.Drawing.Size(99, 22);
            this.jPEGToolStripMenuItem.Text = "JPEG";
            // 
            // экспортироватьСтраницуКакToolStripMenuItem
            // 
            this.экспортироватьСтраницуКакToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.dOCToolStripMenuItem1,
            this.pNGToolStripMenuItem1,
            this.jPGToolStripMenuItem});
            this.экспортироватьСтраницуКакToolStripMenuItem.Name = "экспортироватьСтраницуКакToolStripMenuItem";
            this.экспортироватьСтраницуКакToolStripMenuItem.Size = new System.Drawing.Size(242, 22);
            this.экспортироватьСтраницуКакToolStripMenuItem.Text = "Экспортировать страницу как";
            // 
            // dOCToolStripMenuItem1
            // 
            this.dOCToolStripMenuItem1.Name = "dOCToolStripMenuItem1";
            this.dOCToolStripMenuItem1.Size = new System.Drawing.Size(99, 22);
            this.dOCToolStripMenuItem1.Text = "DOC";
            // 
            // pNGToolStripMenuItem1
            // 
            this.pNGToolStripMenuItem1.Name = "pNGToolStripMenuItem1";
            this.pNGToolStripMenuItem1.Size = new System.Drawing.Size(99, 22);
            this.pNGToolStripMenuItem1.Text = "PNG";
            this.pNGToolStripMenuItem1.Click += new System.EventHandler(this.pNGToolStripMenuItem1_Click);
            // 
            // jPGToolStripMenuItem
            // 
            this.jPGToolStripMenuItem.Name = "jPGToolStripMenuItem";
            this.jPGToolStripMenuItem.Size = new System.Drawing.Size(99, 22);
            this.jPGToolStripMenuItem.Text = "JPG";
            // 
            // saveFileDialog1
            // 
            this.saveFileDialog1.FileOk += new System.ComponentModel.CancelEventHandler(this.saveFileDialog1_FileOk);
            // 
            // ViewDocument
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(545, 685);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.pictureBox1);
            this.Controls.Add(this.menuStrip1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MainMenuStrip = this.menuStrip1;
            this.Name = "ViewDocument";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Просмотр документа -";
            this.Load += new System.EventHandler(this.ViewDocument_Load);
            this.Resize += new System.EventHandler(this.ViewDocument_Resize);
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.PictureBox pictureBox1;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.TextBox pageNum;
        private System.Windows.Forms.Button btn_back;
        private System.Windows.Forms.Button btn_next;
        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem файлToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem эскпортироватьДокументКакToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem docToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem pNGToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem jPEGToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem экспортироватьСтраницуКакToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem dOCToolStripMenuItem1;
        private System.Windows.Forms.ToolStripMenuItem pNGToolStripMenuItem1;
        private System.Windows.Forms.ToolStripMenuItem jPGToolStripMenuItem;
        private System.Windows.Forms.SaveFileDialog saveFileDialog1;
    }
}