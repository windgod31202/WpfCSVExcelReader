using Microsoft.Win32;
using OfficeOpenXml;
using System.IO;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace WpfCSVExcelReader;

/// <summary>
/// Interaction logic for MainWindow.xaml
/// </summary>
public partial class MainWindow : Window
{
	private List<DataModel> originalData = new List<DataModel>(); // 儲存原始資料
	private List<DataModel> data; // 儲存讀取的資料
	private ExcelPackage currentPackage; // 當前的 Excel 包
	private Dictionary<string, List<DataModel>> worksheetData = new Dictionary<string, List<DataModel>>(); // 儲存所有工作表的資料

	public MainWindow()
    {
        InitializeComponent();

		// 初始化 data 為空列表，避免未初始化的情況
		data = new List<DataModel>();

		// 為 TextBox 添加水印
		filterTimeStampTextBox.Loaded += (sender, e) => AddWatermark(filterTimeStampTextBox, "篩選 TimeStamp");
		filterNameTextBox.Loaded += (sender, e) => AddWatermark(filterNameTextBox, "篩選 Name");
		filterTypeTextBox.Loaded += (sender, e) => AddWatermark(filterTypeTextBox, "篩選 Type");
		filterRareTextBox.Loaded += (sender, e) => AddWatermark(filterRareTextBox, "篩選 Rare");
	}

	private void LoadCsv_Click(object sender, RoutedEventArgs e)
	{
		OpenFileDialog openFileDialog = new OpenFileDialog
		{
			Filter = "CSV Files (*.csv)|*.csv",
			Title = "選擇 CSV 檔案"
		};

		if (openFileDialog.ShowDialog() == true)
		{
			LoadCsv(openFileDialog.FileName);
		}
	}

	// 按鈕事件 - 讀取 Excel
	private void LoadExcel_Click(object sender, RoutedEventArgs e)
	{
		OpenFileDialog openFileDialog = new OpenFileDialog
		{
			Filter = "Excel Files (*.xlsx)|*.xlsx",
			Title = "選擇 Excel 檔案"
		};

		if (openFileDialog.ShowDialog() == true)
		{
			LoadExcel(openFileDialog.FileName);
		}
	}

	// 讀取 CSV 檔案
	private void LoadCsv(string filePath)
	{
		var data = File.ReadAllLines(filePath)
					   .Skip(1) // 跳過標題
					   .Select(line => line.Split(',')) // 解析 CSV
					   .Select(values => new DataModel
					   {
						   TimeStamp = values[0],
						   Name = values[1],
						   Type = values.Length > 2 ? values[2] : string.Empty,
						   Rare = values.Length > 3 ? values[3] : string.Empty
					   })
					   .ToList();

		// 儲存資料
		originalData = data;
		dataGrid.ItemsSource = originalData; // 顯示資料

		// 更新統計資料
		UpdateStatistics(originalData);
	}

	// 讀取 Excel 檔案
	private void LoadExcel(string filePath)
	{
		var data = new List<DataModel>();

		ExcelPackage.LicenseContext = LicenseContext.NonCommercial; // EPPlus 6 需要授權
		using (var package = new ExcelPackage(new FileInfo(filePath)))
		{
			// 獲取所有工作表
			var worksheets = package.Workbook.Worksheets;

			// 清空 ListBox 並添加工作表名稱
			worksheetListBox.Items.Clear();

			// 遍歷工作表
			foreach (var worksheet in worksheets)
			{
				// 將工作表名稱添加到 ListBox
				worksheetListBox.Items.Add(worksheet.Name);

				// 儲存每個工作表的資料
				var worksheetDataList = new List<DataModel>();
				int rowCount = worksheet.Dimension.Rows;
				for (int row = 2; row <= rowCount; row++) // 從第二行開始讀取 (第一行是標題)
				{
					worksheetDataList.Add(new DataModel
					{
						TimeStamp = worksheet.Cells[row, 1].Text,
						Name = worksheet.Cells[row, 2].Text,
						Type = worksheet.Cells[row, 3].Text,
						Rare = worksheet.Cells[row, 4].Text,
						TotleNumber = worksheet.Cells[row, 5].Text,
						OneRoundOfNumber = worksheet.Cells[row, 6].Text
					});
				}

				// 儲存該工作表的資料
				worksheetData[worksheet.Name] = worksheetDataList;
			}
		}

		// 設置初始的工作表資料
		if (worksheetListBox.Items.Count > 0)
		{
			worksheetListBox.SelectedIndex = 0; // 默認選擇第一個工作表
			UpdateWorksheetDataGrid();
		}
	}

	private void UpdateWorksheetDataGrid()
	{
		// 確保有選擇工作表
		if (worksheetListBox.SelectedItem != null)
		{
			string selectedWorksheetName = worksheetListBox.SelectedItem.ToString();

			// 確保該工作表存在
			if (worksheetData.ContainsKey(selectedWorksheetName))
			{
				// 取得當前選擇工作表的資料
				var currentWorksheetData = worksheetData[selectedWorksheetName];

				// 更新 DataGrid 顯示該工作表的資料
				dataGrid.ItemsSource = currentWorksheetData;

				// 更新統計資料
				UpdateStatistics(currentWorksheetData);
			}
		}
	}


	// 當選擇不同的工作表時，更新顯示的資料
	private void WorksheetListBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
	{
		if (worksheetListBox.SelectedItem != null)
		{
			string? selectedWorksheetName = worksheetListBox.SelectedItem.ToString();

			// 如果工作表存在
			if (worksheetData.ContainsKey(selectedWorksheetName))
			{
				var currentWorksheetData = worksheetData[selectedWorksheetName];
				dataGrid.ItemsSource = currentWorksheetData; // 顯示該工作表的資料
				UpdateStatistics(currentWorksheetData); // 更新統計資料
			}
		}
	}
	// 根據選擇的工作表載入資料
	private void LoadDataFromSelectedWorksheet(string? worksheetName)
	{
		var data = new List<DataModel>();

		// 根據選擇的工作表名稱獲取該工作表
		var worksheet = currentPackage.Workbook.Worksheets[worksheetName];
		int rowCount = worksheet.Dimension.Rows;

		for (int row = 2; row <= rowCount; row++) // 跳過標題
		{
			data.Add(new DataModel
			{
				TimeStamp = worksheet.Cells[row, 1].Text,
				Name = worksheet.Cells[row, 2].Text,
				Type = worksheet.Cells[row, 3].Text,
				Rare = worksheet.Cells[row, 4].Text,
				TotleNumber = worksheet.Cells[row, 5].Text,
				OneRoundOfNumber = worksheet.Cells[row, 6].Text
			});
		}

		// 顯示該工作表的資料
		dataGrid.ItemsSource = data;

		// 更新統計資料
		UpdateStatistics(data);
	}


	// 篩選資料
	private void FilterButton_Click(object sender, RoutedEventArgs e)
	{
		// 取得篩選條件
		string timeStampFilter = filterTimeStampTextBox.Text.Trim();  // 去除前後空格
		string nameFilter = filterNameTextBox.Text.Trim();
		string typeFilter = filterTypeTextBox.Text.Trim();
		string rareFilter = filterRareTextBox.Text.Trim();

		// 確保選擇了某一個工作表
		if (worksheetListBox.SelectedItem != null)
		{
			string? selectedWorksheetName = worksheetListBox.SelectedItem.ToString();

			// 如果工作表存在
			if (worksheetData.ContainsKey(selectedWorksheetName))
			{
				var currentWorksheetData = worksheetData[selectedWorksheetName];

				// 在這裡進行篩選
				var filteredData = currentWorksheetData.Where(d => 
					(string.IsNullOrEmpty(timeStampFilter) || d.TimeStamp.Contains(timeStampFilter, StringComparison.OrdinalIgnoreCase)) &&
					(string.IsNullOrEmpty(nameFilter) || d.Name.Contains(nameFilter, StringComparison.OrdinalIgnoreCase)) &&
					(string.IsNullOrEmpty(typeFilter) || d.Type.Contains(typeFilter, StringComparison.OrdinalIgnoreCase)) &&
					(string.IsNullOrEmpty(rareFilter) || d.Rare.Contains(rareFilter, StringComparison.OrdinalIgnoreCase))
				).ToList();

				// 更新顯示篩選後的資料
				dataGrid.ItemsSource = filteredData;

				// 更新統計資料
				UpdateStatistics(filteredData);
			}
		}
	}

	// 更新統計資料
	private void UpdateStatistics(List<DataModel> filteredData)
	{
		// 更新篩選前的資料總數
		totalCountText.Text = $"資料總數: {filteredData.Count}";
		// 更新篩選後的資料數
		filteredCountText.Text = $"篩選後資料數: {filteredData.Count}";
	}

	private void AddWatermark(TextBox textBox, string watermark)
	{
		// 確保 TextBox 渲染後再添加水印
		if (textBox.IsLoaded)
		{
			var adornerLayer = AdornerLayer.GetAdornerLayer(textBox);
			if (adornerLayer != null)
			{
				var adorner = new TextBoxWatermarkAdorner(textBox, watermark);
				adornerLayer.Add(adorner);
			}
		}
	}
	// 如果選擇工作表時，更新資料顯示
	private void WorksheetListBox_SelectionChanged(object sender, RoutedEventArgs e)
	{
		// 確保選擇了某一個工作表
		if (worksheetListBox.SelectedItem != null)
		{
			string? selectedWorksheetName = worksheetListBox.SelectedItem.ToString();
			// 根據選擇的工作表名稱，重新載入資料
			LoadDataFromSelectedWorksheet(selectedWorksheetName);
		}
	}
}