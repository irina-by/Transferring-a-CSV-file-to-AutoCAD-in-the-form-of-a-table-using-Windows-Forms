namespace for_AutoCAD_CSV_files
{
    partial class AutoCADTextInserterForm
    {
        /// <summary>
        /// Обязательная переменная конструктора.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Освободить все используемые ресурсы.
        /// </summary>
        /// <param name="disposing">истинно, если управляемый ресурс должен быть удален; иначе ложно.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Код, автоматически созданный конструктором форм Windows

        /// <summary>
        /// Требуемый метод для поддержки конструктора — не изменяйте 
        /// содержимое этого метода с помощью редактора кода.
        /// </summary>
        private void InitializeComponent()
        {
            this.button_choose = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // button_choose
            // 
            this.button_choose.DialogResult = System.Windows.Forms.DialogResult.OK;
            this.button_choose.ForeColor = System.Drawing.Color.Black;
            this.button_choose.Location = new System.Drawing.Point(151, 156);
            this.button_choose.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.button_choose.Name = "button_choose";
            this.button_choose.Size = new System.Drawing.Size(105, 49);
            this.button_choose.TabIndex = 0;
            this.button_choose.Text = "Сhoose";
            this.button_choose.UseVisualStyleBackColor = true;
            this.button_choose.Click += new System.EventHandler(this.button_choose_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 16.2F, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label1.Location = new System.Drawing.Point(102, 75);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(204, 32);
            this.label1.TabIndex = 1;
            this.label1.Text = "Select CSV file";
            // 
            // AutoCADTextInserterForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 19F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(393, 304);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.button_choose);
            this.Font = new System.Drawing.Font("Times New Roman", 10.2F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.ForeColor = System.Drawing.SystemColors.ControlText;
            this.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.MaximizeBox = false;
            this.Name = "AutoCADTextInserterForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Select CSV file";
            this.TopMost = true;
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button button_choose;
        private System.Windows.Forms.Label label1;
    }
}

