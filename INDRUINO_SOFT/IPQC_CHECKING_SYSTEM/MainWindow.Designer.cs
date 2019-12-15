namespace IPQC_CHECKING_SYSTEM
{
    partial class MainWindow
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
            this.components = new System.ComponentModel.Container();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle16 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle17 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle18 = new System.Windows.Forms.DataGridViewCellStyle();
            this.lbl_TitleSyestem = new System.Windows.Forms.Label();
            this.pnl_Body = new System.Windows.Forms.Panel();
            this.lbl_CurrentDate = new System.Windows.Forms.Label();
            this.dgvData = new System.Windows.Forms.DataGridView();
            this.PartNumber = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.TimeIn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Type = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Result = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Status = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.pnl_Header = new System.Windows.Forms.Panel();
            this.label2 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.SendData = new System.Windows.Forms.Button();
            this.txtValue = new System.Windows.Forms.TextBox();
            this.txtPartNumber = new System.Windows.Forms.TextBox();
            this.t_Reload_Data = new System.Windows.Forms.Timer(this.components);
            this.t_Reload_View = new System.Windows.Forms.Timer(this.components);
            this.pnl_Body.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgvData)).BeginInit();
            this.pnl_Header.SuspendLayout();
            this.SuspendLayout();
            // 
            // lbl_TitleSyestem
            // 
            this.lbl_TitleSyestem.Anchor = System.Windows.Forms.AnchorStyles.Top;
            this.lbl_TitleSyestem.AutoSize = true;
            this.lbl_TitleSyestem.Font = new System.Drawing.Font("Arial Narrow", 20.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lbl_TitleSyestem.Location = new System.Drawing.Point(356, 7);
            this.lbl_TitleSyestem.Name = "lbl_TitleSyestem";
            this.lbl_TitleSyestem.Size = new System.Drawing.Size(385, 31);
            this.lbl_TitleSyestem.TabIndex = 4;
            this.lbl_TitleSyestem.Text = "DIGITAL CHECKING IPQC SYSTEM";
            // 
            // pnl_Body
            // 
            this.pnl_Body.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.pnl_Body.BackColor = System.Drawing.SystemColors.ButtonHighlight;
            this.pnl_Body.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.pnl_Body.Controls.Add(this.lbl_CurrentDate);
            this.pnl_Body.Controls.Add(this.dgvData);
            this.pnl_Body.Location = new System.Drawing.Point(0, 58);
            this.pnl_Body.Name = "pnl_Body";
            this.pnl_Body.Size = new System.Drawing.Size(1052, 470);
            this.pnl_Body.TabIndex = 9;
            // 
            // lbl_CurrentDate
            // 
            this.lbl_CurrentDate.AutoSize = true;
            this.lbl_CurrentDate.Font = new System.Drawing.Font("Arial Narrow", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lbl_CurrentDate.Location = new System.Drawing.Point(3, 0);
            this.lbl_CurrentDate.Name = "lbl_CurrentDate";
            this.lbl_CurrentDate.Size = new System.Drawing.Size(133, 23);
            this.lbl_CurrentDate.TabIndex = 5;
            this.lbl_CurrentDate.Text = "Date: 20/10/2019";
            // 
            // dgvData
            // 
            this.dgvData.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dgvData.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
            this.dgvData.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells;
            this.dgvData.BorderStyle = System.Windows.Forms.BorderStyle.None;
            dataGridViewCellStyle16.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter;
            dataGridViewCellStyle16.BackColor = System.Drawing.Color.DodgerBlue;
            dataGridViewCellStyle16.Font = new System.Drawing.Font("Arial", 30F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle16.ForeColor = System.Drawing.SystemColors.WindowText;
            dataGridViewCellStyle16.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle16.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle16.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.dgvData.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle16;
            this.dgvData.ColumnHeadersHeight = 50;
            this.dgvData.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing;
            this.dgvData.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.PartNumber,
            this.TimeIn,
            this.Type,
            this.Result,
            this.Status});
            dataGridViewCellStyle17.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle17.BackColor = System.Drawing.SystemColors.Window;
            dataGridViewCellStyle17.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle17.ForeColor = System.Drawing.SystemColors.ControlText;
            dataGridViewCellStyle17.SelectionBackColor = System.Drawing.SystemColors.Window;
            dataGridViewCellStyle17.SelectionForeColor = System.Drawing.SystemColors.ControlText;
            dataGridViewCellStyle17.WrapMode = System.Windows.Forms.DataGridViewTriState.False;
            this.dgvData.DefaultCellStyle = dataGridViewCellStyle17;
            this.dgvData.GridColor = System.Drawing.SystemColors.ControlLight;
            this.dgvData.Location = new System.Drawing.Point(1, 25);
            this.dgvData.MultiSelect = false;
            this.dgvData.Name = "dgvData";
            this.dgvData.ReadOnly = true;
            this.dgvData.RowHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.Sunken;
            dataGridViewCellStyle18.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle18.BackColor = System.Drawing.SystemColors.Control;
            dataGridViewCellStyle18.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle18.ForeColor = System.Drawing.SystemColors.WindowText;
            dataGridViewCellStyle18.SelectionBackColor = System.Drawing.Color.Transparent;
            dataGridViewCellStyle18.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle18.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.dgvData.RowHeadersDefaultCellStyle = dataGridViewCellStyle18;
            this.dgvData.RowHeadersVisible = false;
            this.dgvData.RowHeadersWidth = 42;
            this.dgvData.RowTemplate.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter;
            this.dgvData.RowTemplate.DefaultCellStyle.Font = new System.Drawing.Font("Arial Narrow", 24.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.dgvData.RowTemplate.DefaultCellStyle.ForeColor = System.Drawing.Color.Black;
            this.dgvData.RowTemplate.Height = 50;
            this.dgvData.RowTemplate.ReadOnly = true;
            this.dgvData.RowTemplate.Resizable = System.Windows.Forms.DataGridViewTriState.True;
            this.dgvData.Size = new System.Drawing.Size(1049, 446);
            this.dgvData.TabIndex = 3;
            this.dgvData.CellFormatting += new System.Windows.Forms.DataGridViewCellFormattingEventHandler(this.dgvData_CellFormatting);
            this.dgvData.CellPainting += new System.Windows.Forms.DataGridViewCellPaintingEventHandler(this.dgvData_CellPainting);
            // 
            // PartNumber
            // 
            this.PartNumber.DataPropertyName = "PartNumber";
            this.PartNumber.HeaderText = "PART NUMBER";
            this.PartNumber.Name = "PartNumber";
            this.PartNumber.ReadOnly = true;
            // 
            // TimeIn
            // 
            this.TimeIn.DataPropertyName = "TimeSubmit";
            this.TimeIn.HeaderText = "TIME IN";
            this.TimeIn.Name = "TimeIn";
            this.TimeIn.ReadOnly = true;
            // 
            // Type
            // 
            this.Type.DataPropertyName = "Type";
            this.Type.HeaderText = "TYPE";
            this.Type.Name = "Type";
            this.Type.ReadOnly = true;
            // 
            // Result
            // 
            this.Result.DataPropertyName = "Result";
            this.Result.HeaderText = "RESULT";
            this.Result.Name = "Result";
            this.Result.ReadOnly = true;
            // 
            // Status
            // 
            this.Status.DataPropertyName = "Status";
            this.Status.HeaderText = "STATUS";
            this.Status.Name = "Status";
            this.Status.ReadOnly = true;
            // 
            // pnl_Header
            // 
            this.pnl_Header.BackColor = System.Drawing.Color.Yellow;
            this.pnl_Header.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.pnl_Header.Controls.Add(this.label2);
            this.pnl_Header.Controls.Add(this.label1);
            this.pnl_Header.Controls.Add(this.SendData);
            this.pnl_Header.Controls.Add(this.txtValue);
            this.pnl_Header.Controls.Add(this.txtPartNumber);
            this.pnl_Header.Controls.Add(this.lbl_TitleSyestem);
            this.pnl_Header.Dock = System.Windows.Forms.DockStyle.Top;
            this.pnl_Header.Location = new System.Drawing.Point(0, 0);
            this.pnl_Header.Name = "pnl_Header";
            this.pnl_Header.Size = new System.Drawing.Size(1052, 50);
            this.pnl_Header.TabIndex = 8;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Arial Narrow", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.Location = new System.Drawing.Point(112, 2);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(44, 20);
            this.label2.TabIndex = 9;
            this.label2.Text = "Value";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Arial Narrow", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(6, 3);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(86, 20);
            this.label1.TabIndex = 6;
            this.label1.Text = "Part number";
            // 
            // SendData
            // 
            this.SendData.Location = new System.Drawing.Point(222, 21);
            this.SendData.Name = "SendData";
            this.SendData.Size = new System.Drawing.Size(75, 23);
            this.SendData.TabIndex = 8;
            this.SendData.Text = "Send";
            this.SendData.UseVisualStyleBackColor = true;
            this.SendData.Click += new System.EventHandler(this.SendData_Click);
            // 
            // txtValue
            // 
            this.txtValue.Location = new System.Drawing.Point(116, 23);
            this.txtValue.Name = "txtValue";
            this.txtValue.Size = new System.Drawing.Size(100, 20);
            this.txtValue.TabIndex = 7;
            // 
            // txtPartNumber
            // 
            this.txtPartNumber.Location = new System.Drawing.Point(10, 23);
            this.txtPartNumber.Name = "txtPartNumber";
            this.txtPartNumber.Size = new System.Drawing.Size(100, 20);
            this.txtPartNumber.TabIndex = 6;
            // 
            // t_Reload_Data
            // 
            this.t_Reload_Data.Interval = 5000;
            this.t_Reload_Data.Tick += new System.EventHandler(this.t_Reload_Data_Tick);
            // 
            // t_Reload_View
            // 
            this.t_Reload_View.Interval = 2000;
            this.t_Reload_View.Tick += new System.EventHandler(this.t_Reload_View_Tick);
            // 
            // MainWindow
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1052, 529);
            this.Controls.Add(this.pnl_Body);
            this.Controls.Add(this.pnl_Header);
            this.Name = "MainWindow";
            this.Text = "DisplayChecking";
            this.WindowState = System.Windows.Forms.FormWindowState.Maximized;
            this.Load += new System.EventHandler(this.MainWindow_Load);
            this.Shown += new System.EventHandler(this.MainWindow_Shown);
            this.pnl_Body.ResumeLayout(false);
            this.pnl_Body.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgvData)).EndInit();
            this.pnl_Header.ResumeLayout(false);
            this.pnl_Header.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Label lbl_TitleSyestem;
        private System.Windows.Forms.Panel pnl_Body;
        private System.Windows.Forms.Label lbl_CurrentDate;
        public System.Windows.Forms.DataGridView dgvData;
        private System.Windows.Forms.Panel pnl_Header;
        private System.Windows.Forms.DataGridViewTextBoxColumn PartNumber;
        private System.Windows.Forms.DataGridViewTextBoxColumn TimeIn;
        private System.Windows.Forms.DataGridViewTextBoxColumn Type;
        private System.Windows.Forms.DataGridViewTextBoxColumn Result;
        private System.Windows.Forms.DataGridViewTextBoxColumn Status;
        private System.Windows.Forms.Timer t_Reload_Data;
        private System.Windows.Forms.Timer t_Reload_View;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button SendData;
        private System.Windows.Forms.TextBox txtValue;
        private System.Windows.Forms.TextBox txtPartNumber;
    }
}

