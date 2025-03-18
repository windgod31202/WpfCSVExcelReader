using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows;

namespace WpfCSVExcelReader
{
	public class WatermarkTextBox : TextBox
	{
		public static readonly DependencyProperty WatermarkProperty =
			DependencyProperty.Register("Watermark", typeof(string), typeof(WatermarkTextBox));

		public string Watermark
		{
			get { return (string)GetValue(WatermarkProperty); }
			set { SetValue(WatermarkProperty, value); }
		}

		protected override void OnGotFocus(RoutedEventArgs e)
		{
			base.OnGotFocus(e);
			RemoveWatermark();
		}

		protected override void OnLostFocus(RoutedEventArgs e)
		{
			base.OnLostFocus(e);
			AddWatermark();
		}

		private void AddWatermark()
		{
			if (string.IsNullOrEmpty(this.Text) && !this.IsFocused)
			{
				var adornerLayer = AdornerLayer.GetAdornerLayer(this);
				var adorner = new TextBoxWatermarkAdorner(this, Watermark);
				adornerLayer.Add(adorner);
			}
		}

		private void RemoveWatermark()
		{
			if (string.IsNullOrEmpty(this.Text))
			{
				var adornerLayer = AdornerLayer.GetAdornerLayer(this);
				var adorners = adornerLayer.GetAdorners(this);
				foreach (var adorner in adorners)
				{
					if (adorner is TextBoxWatermarkAdorner)
					{
						adornerLayer.Remove(adorner);
					}
				}
			}
		}
	}
}
