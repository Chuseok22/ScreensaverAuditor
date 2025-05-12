// UI/MainForm.Designer.cs
using System;
using System.Windows.Forms;

namespace ScreensaverAuditor.UI
{
    partial class MainForm
    {
        /// <summary>
        /// 필수 디자이너 변수입니다.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 사용 중인 모든 리소스를 정리합니다.
        /// </summary>
        /// <param name="disposing">관리되는 리소스를 삭제해야 하면 true이고, 그렇지 않으면 false입니다.</param>
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
        /// 디자이너 지원에 필요한 메서드입니다.
        /// 이 메서드의 내용을 코드 편집기로 수정하지 마세요.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.tabControl = new System.Windows.Forms.TabControl();
            this.tabAudit = new System.Windows.Forms.TabPage();
            this.btnEnablePolicy = new System.Windows.Forms.Button();
            this.btnRunAudit = new System.Windows.Forms.Button();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.statusLabel = new System.Windows.Forms.Label();
            this.progressBar = new System.Windows.Forms.ProgressBar();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.btnBrowse = new System.Windows.Forms.Button();
            this.txtOutputPath = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.txtUsername = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.datePickerEnd = new System.Windows.Forms.DateTimePicker();
            this.label2 = new System.Windows.Forms.Label();
            this.datePickerStart = new System.Windows.Forms.DateTimePicker();
            this.label1 = new System.Windows.Forms.Label();
            this.tabResults = new System.Windows.Forms.TabPage();
            this.btnOpenExcel = new System.Windows.Forms.Button();
            this.lblOutputPath = new System.Windows.Forms.Label();
            this.dataGridResults = new System.Windows.Forms.DataGridView();
            this.tabHelp = new System.Windows.Forms.TabPage();
            this.richTextBoxHelp = new System.Windows.Forms.RichTextBox();
            this.toolTip = new System.Windows.Forms.ToolTip(this.components);
            this.tabControl.SuspendLayout();
            this.tabAudit.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.groupBox1.SuspendLayout();
            this.tabResults.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridResults)).BeginInit();
            this.tabHelp.SuspendLayout();
            this.SuspendLayout();
            // 
            // tabControl
            // 
            this.tabControl.Controls.Add(this.tabAudit);
            this.tabControl.Controls.Add(this.tabResults);
            this.tabControl.Controls.Add(this.tabHelp);
            this.tabControl.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tabControl.Location = new System.Drawing.Point(0, 0);
            this.tabControl.Name = "tabControl";
            this.tabControl.SelectedIndex = 0;
            this.tabControl.Size = new System.Drawing.Size(884, 561);
            this.tabControl.TabIndex = 0;
            // 
            // tabAudit
            // 
            this.tabAudit.Controls.Add(this.btnEnablePolicy);
            this.tabAudit.Controls.Add(this.btnRunAudit);
            this.tabAudit.Controls.Add(this.groupBox2);
            this.tabAudit.Controls.Add(this.groupBox1);
            this.tabAudit.Location = new System.Drawing.Point(4, 24);
            this.tabAudit.Name = "tabAudit";
            this.tabAudit.Padding = new System.Windows.Forms.Padding(3);
            this.tabAudit.Size = new System.Drawing.Size(876, 533);
            this.tabAudit.TabIndex = 0;
            this.tabAudit.Text = "감사 설정";
            this.tabAudit.UseVisualStyleBackColor = true;
            // 
            // btnEnablePolicy
            // 
            this.btnEnablePolicy.BackColor = System.Drawing.Color.LightSteelBlue;
            this.btnEnablePolicy.Font = new System.Drawing.Font("맑은 고딕", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point);
            this.btnEnablePolicy.Location = new System.Drawing.Point(242, 336);
            this.btnEnablePolicy.Name = "btnEnablePolicy";
            this.btnEnablePolicy.Size = new System.Drawing.Size(168, 46);
            this.btnEnablePolicy.TabIndex = 3;
            this.btnEnablePolicy.Text = "감사 정책 활성화";
            this.toolTip.SetToolTip(this.btnEnablePolicy, "화면보호기 감사 정책을 활성화합니다. 관리자 권한이 필요합니다.");
            this.btnEnablePolicy.UseVisualStyleBackColor = false;
            this.btnEnablePolicy.Click += new System.EventHandler(this.btnEnablePolicy_Click);
            // 
            // btnRunAudit
            // 
            this.btnRunAudit.BackColor = System.Drawing.Color.MediumSeaGreen;
            this.btnRunAudit.Font = new System.Drawing.Font("맑은 고딕", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point);
            this.btnRunAudit.ForeColor = System.Drawing.Color.White;
            this.btnRunAudit.Location = new System.Drawing.Point(466, 336);
            this.btnRunAudit.Name = "btnRunAudit";
            this.btnRunAudit.Size = new System.Drawing.Size(168, 46);
            this.btnRunAudit.TabIndex = 4;
            this.btnRunAudit.Text = "감사 실행";
            this.toolTip.SetToolTip(this.btnRunAudit, "설정한 조건으로 화면보호기 감사를 실행합니다.");
            this.btnRunAudit.UseVisualStyleBackColor = false;
            this.btnRunAudit.Click += new System.EventHandler(this.btnRunAudit_Click);
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.statusLabel);
            this.groupBox2.Controls.Add(this.progressBar);
            this.groupBox2.Location = new System.Drawing.Point(25, 399);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(826, 115);
            this.groupBox2.TabIndex = 5;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "상태";
            // 
            // statusLabel
            // 
            this.statusLabel.AutoSize = true;
            this.statusLabel.Location = new System.Drawing.Point(19, 27);
            this.statusLabel.Name = "statusLabel";
            this.statusLabel.Size = new System.Drawing.Size(87, 15);
            this.statusLabel.TabIndex = 1;
            this.statusLabel.Text = "준비 완료";
            // 
            // progressBar
            // 
            this.progressBar.Location = new System.Drawing.Point(19, 58);
            this.progressBar.Name = "progressBar";
            this.progressBar.Size = new System.Drawing.Size(792, 27);
            this.progressBar.TabIndex = 0;
            this.progressBar.Visible = false;
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.btnBrowse);
            this.groupBox1.Controls.Add(this.txtOutputPath);
            this.groupBox1.Controls.Add(this.label4);
            this.groupBox1.Controls.Add(this.txtUsername);
            this.groupBox1.Controls.Add(this.label3);
            this.groupBox1.Controls.Add(this.datePickerEnd);
            this.groupBox1.Controls.Add(this.label2);
            this.groupBox1.Controls.Add(this.datePickerStart);
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Location = new System.Drawing.Point(25, 21);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(826, 297);
            this.groupBox1.TabIndex = 0;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "감사 설정";
            // 
            // btnBrowse
            // 
            this.btnBrowse.Location = new System.Drawing.Point(708, 210);
            this.btnBrowse.Name = "btnBrowse";
            this.btnBrowse.Size = new System.Drawing.Size(103, 28);
            this.btnBrowse.TabIndex = 8;
            this.btnBrowse.Text = "찾아보기...";
            this.btnBrowse.UseVisualStyleBackColor = true;
            this.btnBrowse.Click += new System.EventHandler(this.btnBrowse_Click);
            // 
            // txtOutputPath
            // 
            this.txtOutputPath.Location = new System.Drawing.Point(153, 214);
            this.txtOutputPath.Name = "txtOutputPath";
            this.txtOutputPath.Size = new System.Drawing.Size(534, 23);
            this.txtOutputPath.TabIndex = 7;
            this.txtOutputPath.Text = "화면보호기감사.xlsx";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(19, 217);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(89, 15);
            this.label4.TabIndex = 6;
            this.label4.Text = "결과 저장 경로:";
            // 
            // txtUsername
            // 
            this.txtUsername.Location = new System.Drawing.Point(153, 156);
            this.txtUsername.Name = "txtUsername";
            this.txtUsername.Size = new System.Drawing.Size(202, 23);
            this.txtUsername.TabIndex = 5;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(19, 159);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(117, 15);
            this.label3.TabIndex = 4;
            this.label3.Text = "사용자 이름 (선택적):";
            // 
            // datePickerEnd
            // 
            this.datePickerEnd.Location = new System.Drawing.Point(153, 98);
            this.datePickerEnd.Name = "datePickerEnd";
            this.datePickerEnd.Size = new System.Drawing.Size(202, 23);
            this.datePickerEnd.TabIndex = 3;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(19, 104);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(62, 15);
            this.label2.TabIndex = 2;
            this.label2.Text = "종료 날짜:";
            // 
            // datePickerStart
            // 
            this.datePickerStart.Location = new System.Drawing.Point(153, 43);
            this.datePickerStart.Name = "datePickerStart";
            this.datePickerStart.Size = new System.Drawing.Size(202, 23);
            this.datePickerStart.TabIndex = 1;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(19, 49);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(62, 15);
            this.label1.TabIndex = 0;
            this.label1.Text = "시작 날짜:";
            // 
            // tabResults
            // 
            this.tabResults.Controls.Add(this.btnOpenExcel);
            this.tabResults.Controls.Add(this.lblOutputPath);
            this.tabResults.Controls.Add(this.dataGridResults);
            this.tabResults.Location = new System.Drawing.Point(4, 24);
            this.tabResults.Name = "tabResults";
            this.tabResults.Padding = new System.Windows.Forms.Padding(3);
            this.tabResults.Size = new System.Drawing.Size(876, 533);
            this.tabResults.TabIndex = 1;
            this.tabResults.Text = "결과";
            this.tabResults.UseVisualStyleBackColor = true;
            // 
            // btnOpenExcel
            // 
            this.btnOpenExcel.Location = new System.Drawing.Point(741, 489);
            this.btnOpenExcel.Name = "btnOpenExcel";
            this.btnOpenExcel.Size = new System.Drawing.Size(127, 36);
            this.btnOpenExcel.TabIndex = 2;
            this.btnOpenExcel.Text = "Excel 파일 열기";
            this.btnOpenExcel.UseVisualStyleBackColor = true;
            this.btnOpenExcel.Visible = false;
            this.btnOpenExcel.Click += new System.EventHandler(this.btnOpenExcel_Click);
            // 
            // lblOutputPath
            // 
            this.lblOutputPath.AutoSize = true;
            this.lblOutputPath.Location = new System.Drawing.Point(8, 500);
            this.lblOutputPath.Name = "lblOutputPath";
            this.lblOutputPath.Size = new System.Drawing.Size(0, 15);
            this.lblOutputPath.TabIndex = 1;
            // 
            // dataGridResults
            // 
            this.dataGridResults.AllowUserToAddRows = false;
            this.dataGridResults.AllowUserToDeleteRows = false;
            this.dataGridResults.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
            this.dataGridResults.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridResults.Dock = System.Windows.Forms.DockStyle.Top;
            this.dataGridResults.Location = new System.Drawing.Point(3, 3);
            this.dataGridResults.Name = "dataGridResults";
            this.dataGridResults.ReadOnly = true;
            this.dataGridResults.RowTemplate.Height = 25;
            this.dataGridResults.Size = new System.Drawing.Size(870, 480);
            this.dataGridResults.TabIndex = 0;
            // 
            // tabHelp
            // 
            this.tabHelp.Controls.Add(this.richTextBoxHelp);
            this.tabHelp.Location = new System.Drawing.Point(4, 24);
            this.tabHelp.Name = "tabHelp";
            this.tabHelp.Padding = new System.Windows.Forms.Padding(3);
            this.tabHelp.Size = new System.Drawing.Size(876, 533);
            this.tabHelp.TabIndex = 2;
            this.tabHelp.Text = "도움말";
            this.tabHelp.UseVisualStyleBackColor = true;
            // 
            // richTextBoxHelp
            // 
            this.richTextBoxHelp.BackColor = System.Drawing.SystemColors.Window;
            this.richTextBoxHelp.Dock = System.Windows.Forms.DockStyle.Fill;
            this.richTextBoxHelp.Location = new System.Drawing.Point(3, 3);
            this.richTextBoxHelp.Name = "richTextBoxHelp";
            this.richTextBoxHelp.ReadOnly = true;
            this.richTextBoxHelp.Size = new System.Drawing.Size(870, 527);
            this.richTextBoxHelp.TabIndex = 0;
            this.richTextBoxHelp.Text = "# 화면보호기 감사 도구 사용 설명서\n\n1. 감사 정책 활성화 버튼을 클릭하여 감사 정책을 활성화합니다. (관리자 권한 필요)\n\n2. 시작" +
                                           " 날짜와 종료 날짜를 설정합니다.\n\n3. 필요에 따라 사용자 이름을 입력하여 특정 사용자만 조회할 수 있습니다.\n\n4. 결과 저" +
                                           "장 경로를 설정합니다.\n\n5. 감사 실행 버튼을 클릭하여 화면보호기 감사를 시작합니다.\n\n6. 감사 완료 후에는 결과 탭에서 " +
                                           "확인할 수 있으며, Excel 파일로 저장됩니다.";
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(884, 561);
            this.Controls.Add(this.tabControl);
            this.Name = "MainForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "화면보호기 감사 도구 (ScreensaverAuditor)";
            this.tabControl.ResumeLayout(false);
            this.tabAudit.ResumeLayout(false);
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.tabResults.ResumeLayout(false);
            this.tabResults.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridResults)).EndInit();
            this.tabHelp.ResumeLayout(false);
            this.ResumeLayout(false);
        }

        #endregion

        private System.Windows.Forms.TabControl tabControl;
        private System.Windows.Forms.TabPage tabAudit;
        private System.Windows.Forms.TabPage tabResults;
        private System.Windows.Forms.TabPage tabHelp;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.DateTimePicker datePickerStart;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.DateTimePicker datePickerEnd;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox txtUsername;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Button btnBrowse;
        private System.Windows.Forms.TextBox txtOutputPath;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Button btnRunAudit;
        private System.Windows.Forms.Button btnEnablePolicy;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.Label statusLabel;
        private System.Windows.Forms.ProgressBar progressBar;
        private System.Windows.Forms.DataGridView dataGridResults;
        private System.Windows.Forms.Label lblOutputPath;
        private System.Windows.Forms.Button btnOpenExcel;
        private System.Windows.Forms.RichTextBox richTextBoxHelp;
        private System.Windows.Forms.ToolTip toolTip;
    }
}
