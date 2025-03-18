using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Media;
using System.Windows;

namespace WpfCSVExcelReader
{
	public class TextBoxWatermarkAdorner : Adorner
	{
		private readonly TextBlock _watermarkText;
		private readonly TextBox _associatedTextBox;

		public TextBoxWatermarkAdorner(TextBox textBox, string watermark) : base(textBox)
		{
			_associatedTextBox = textBox;

			_watermarkText = new TextBlock
			{
				Text = watermark,
				Foreground = Brushes.Gray,
				Margin = new Thickness(5, 2, 0, 0),
				IsHitTestVisible = false
			};

			// 當 TextBox 內容變更、獲得焦點、或失去焦點時重新繪製 Adorner
			_associatedTextBox.TextChanged += (s, e) => InvalidateVisual();
			_associatedTextBox.GotFocus += (s, e) => InvalidateVisual();
			_associatedTextBox.LostFocus += (s, e) => InvalidateVisual();
		}

		protected override void OnRender(DrawingContext drawingContext)
		{
			base.OnRender(drawingContext);

			// 如果 TextBox 內容為空且沒有焦點，顯示水印
			if (string.IsNullOrEmpty(_associatedTextBox.Text) && !_associatedTextBox.IsFocused)
			{
				drawingContext.DrawText(
					new FormattedText(
						_watermarkText.Text,
						System.Globalization.CultureInfo.CurrentCulture,
						FlowDirection.LeftToRight,
						new Typeface("Segoe UI"),
						_associatedTextBox.FontSize,
						_watermarkText.Foreground,
						VisualTreeHelper.GetDpi(this).PixelsPerDip
					),
					new Point(5, 2) // 可以調整水印顯示的位置
				);
			}
		}
	}
}
