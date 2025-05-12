// Utils/FormResizer.cs
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;

namespace ScreensaverAuditor.Utils
{
    /// <summary>
    /// 폼과 컨트롤의 크기를 화면 해상도에 맞게 조정하는 유틸리티 클래스입니다.
    /// </summary>
    public class FormResizer
    {
        private readonly Dictionary<string, Rectangle> _originalControlRectangles = new Dictionary<string, Rectangle>();
        private readonly Dictionary<string, Font> _originalFonts = new Dictionary<string, Font>();
        private Size _originalFormSize;
        private Form _form;

        /// <summary>
        /// FormResizer 인스턴스를 초기화합니다.
        /// </summary>
        /// <param name="form">크기를 조정할 폼 객체</param>
        public FormResizer(Form form)
        {
            _form = form;
            _originalFormSize = form.Size;
            SaveInitialSizes(form);
        }

        /// <summary>
        /// 폼과 모든 컨트롤의 초기 크기를 저장합니다.
        /// </summary>
        private void SaveInitialSizes(Control control)
        {
            // 컨트롤의 원래 위치와 크기 저장
            _originalControlRectangles[control.Name] = new Rectangle(control.Location, control.Size);
            
            // 폰트가 있는 컨트롤의 원래 폰트 저장
            if (control is Control controlWithFont && controlWithFont.Font != null)
            {
                _originalFonts[control.Name] = controlWithFont.Font;
            }

            // 모든 자식 컨트롤에 대해 재귀적으로 적용
            foreach (Control childControl in control.Controls)
            {
                SaveInitialSizes(childControl);
            }
        }

        /// <summary>
        /// 화면 해상도 변화에 맞게 폼과 컨트롤의 크기를 조정합니다.
        /// </summary>
        public void ResizeForm()
        {
            ResizeControl(_form, _form);
        }

        /// <summary>
        /// 지정된 컨트롤과 그 자식 컨트롤의 크기를 조정합니다.
        /// </summary>
        /// <param name="control">크기를 조정할 컨트롤</param>
        /// <param name="containerControl">컨테이너 컨트롤</param>
        private void ResizeControl(Control control, Control containerControl)
        {
            if (!_originalControlRectangles.ContainsKey(control.Name)) return;

            // 크기 조정 비율 계산
            float xRatio = (float)containerControl.Width / _originalFormSize.Width;
            float yRatio = (float)containerControl.Height / _originalFormSize.Height;

            // 원래 위치와 크기를 가져옴
            Rectangle originalRectangle = _originalControlRectangles[control.Name];

            // 위치와 크기 조정
            control.Left = (int)(originalRectangle.Left * xRatio);
            control.Top = (int)(originalRectangle.Top * yRatio);
            control.Width = (int)(originalRectangle.Width * xRatio);
            control.Height = (int)(originalRectangle.Height * yRatio);

            // 폰트 크기 조정 (필요한 경우)
            if (_originalFonts.ContainsKey(control.Name))
            {
                Font originalFont = _originalFonts[control.Name];
                float ratio = Math.Min(xRatio, yRatio);
                float newSize = originalFont.Size * ratio;
                
                // 최소 및 최대 크기 제한
                newSize = Math.Max(7, Math.Min(newSize, 24));
                
                control.Font = new Font(originalFont.FontFamily, newSize, originalFont.Style);
            }

            // 자식 컨트롤에 대해 크기 조정 적용
            foreach (Control childControl in control.Controls)
            {
                ResizeControl(childControl, control);
            }
        }
    }
}
