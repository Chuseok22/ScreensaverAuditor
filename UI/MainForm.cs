// UI/MainForm.cs
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using ScreensaverAuditor.Models;
using ScreensaverAuditor.Services;
using ScreensaverAuditor.Utils;
using ScreensaverAuditor.ViewModels;

namespace ScreensaverAuditor.UI
{
    public partial class MainForm : Form
    {
        private readonly ScreensaverAuditorViewModel _viewModel;
        private readonly FormResizer _formResizer;

        public MainForm()
        {
            InitializeComponent();
            _viewModel = new ScreensaverAuditorViewModel();
            _formResizer = new FormResizer(this);
            
            // UI 컴포넌트와 뷰모델 바인딩
            datePickerStart.Value = DateTime.Now.AddDays(-7);
            datePickerEnd.Value = DateTime.Now;
            
            // 뷰모델 이벤트 핸들러 등록
            _viewModel.AuditCompleted += ViewModel_AuditCompleted;
            _viewModel.StatusChanged += ViewModel_StatusChanged;
            
            // 폼 크기 조정 이벤트 등록
            this.Resize += MainForm_Resize;
            
            // 버전 정보 표시
            this.Text = $"화면보호기 감사 도구 (ScreensaverAuditor) v2.0";
        }
        
        // 폼 크기 조정 이벤트 처리기
        private void MainForm_Resize(object? sender, EventArgs e)
        {
            if (WindowState != FormWindowState.Minimized)
            {
                _formResizer.ResizeForm();
            }
        }

        private void ViewModel_StatusChanged(object? sender, StatusChangedEventArgs e)
        {
            // UI 스레드에서 실행
            if (InvokeRequired)
            {
                Invoke(new Action(() => UpdateStatus(e.Status)));
                return;
            }
            
            UpdateStatus(e.Status);
        }

        private void UpdateStatus(string status)
        {
            statusLabel.Text = status;
            Application.DoEvents();
        }

        private void ViewModel_AuditCompleted(object? sender, AuditCompletedEventArgs e)
        {
            // UI 스레드에서 실행
            if (InvokeRequired)
            {
                Invoke(new Action(() => DisplayResults(e.Events, e.OutputPath)));
                return;
            }
            
            DisplayResults(e.Events, e.OutputPath);
        }

        private void DisplayResults(List<ScreensaverEvent> events, string outputPath)
        {
            // 결과 데이터그리드뷰에 표시
            dataGridResults.DataSource = null;
            
            if (events == null || events.Count == 0)
            {
                MessageBox.Show("지정한 기간에 화면보호기 이벤트가 없습니다.", "결과 없음", 
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            
            dataGridResults.DataSource = events;
            
            // 결과 파일 경로 표시
            if (!string.IsNullOrEmpty(outputPath))
            {
                lblOutputPath.Text = $"결과가 저장된 경로: {outputPath}";
                btnOpenExcel.Visible = true;
                _lastOutputPath = outputPath;
            }
            
            tabControl.SelectedIndex = 1; // 결과 탭으로 전환
        }
        
        private string? _lastOutputPath;

        private void btnRunAudit_Click(object sender, EventArgs e)
        {
            // 감사 실행 버튼 클릭 시
            var options = new CommandLineOptions
            {
                StartDate = datePickerStart.Value,
                EndDate = datePickerEnd.Value,
                Username = string.IsNullOrWhiteSpace(txtUsername.Text) ? null : txtUsername.Text,
                OutputPath = txtOutputPath.Text
            };
            
            tabControl.SelectedIndex = 0; // 상태 탭으로 전환
            statusLabel.Text = "감사 시작 중...";
            progressBar.Style = ProgressBarStyle.Marquee;
            progressBar.Visible = true;
            
            // 백그라운드 작업으로 감사 실행
            Task.Run(() =>
            {
                try
                {
                    _viewModel.RunAudit(options);
                }
                catch (UnauthorizedAccessException)
                {
                    Invoke(new Action(() =>
                    {
                        MessageBox.Show(
                            "화면보호기 이벤트를 조회하려면 관리자 권한이 필요합니다.\n" +
                            "프로그램을 관리자 권한으로 다시 실행해주세요.", 
                            "권한 부족", 
                            MessageBoxButtons.OK, 
                            MessageBoxIcon.Warning);
                    }));
                }
                catch (Exception ex)
                {
                    Invoke(new Action(() =>
                    {
                        MessageBox.Show($"감사 중 오류가 발생했습니다: {ex.Message}", "오류", 
                            MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }));
                }
                finally
                {
                    // UI 스레드에서 실행
                    Invoke(new Action(() =>
                    {
                        progressBar.Visible = false;
                        progressBar.Style = ProgressBarStyle.Blocks;
                    }));
                }
            });
        }

        private void btnEnablePolicy_Click(object sender, EventArgs e)
        {
            try
            {
                _viewModel.EnableAuditPolicy();
                MessageBox.Show("감사 정책이 성공적으로 활성화되었습니다.", "정책 활성화", 
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (UnauthorizedAccessException)
            {
                MessageBox.Show("감사 정책을 설정하려면 관리자 권한이 필요합니다. " +
                                "프로그램을 관리자 권한으로 다시 실행하세요.", 
                    "권한 부족", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"정책 설정 중 오류가 발생했습니다: {ex.Message}", 
                    "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnBrowse_Click(object sender, EventArgs e)
        {
            using var dialog = new SaveFileDialog
            {
                Filter = "Excel 파일 (*.xlsx)|*.xlsx",
                DefaultExt = "xlsx",
                FileName = "화면보호기감사_" + DateTime.Now.ToString("yyyyMMdd")
            };
            
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                txtOutputPath.Text = dialog.FileName;
            }
        }

        private void btnOpenExcel_Click(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(_lastOutputPath))
            {
                try
                {
                    System.Diagnostics.Process.Start("explorer.exe", _lastOutputPath);
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"파일을 열 수 없습니다: {ex.Message}", "오류", 
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }
    }
}
