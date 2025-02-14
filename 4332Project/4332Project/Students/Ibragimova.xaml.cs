using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using Excel = Microsoft.Office.Interop.Excel;

namespace _4332Project.Students
{
	/// <summary>
	/// Логика взаимодействия для Ibragimova.xaml
	/// </summary>
	public partial class Ibragimova : Window
	{
		public Ibragimova()
		{
			InitializeComponent();
		}

		private void Button_Click(object sender, RoutedEventArgs e)
		{
			OpenFileDialog ofd = new OpenFileDialog()
			{
				DefaultExt = "*.xls;*.xlsx",
				Filter = "файл Excel (Spisok.xlsx)|*.xlsx",
				Title = "Выберите файл базы данных"
			};
			if (!(ofd.ShowDialog() == true))
				return;
			string[,] list;
			Excel.Application ObjWorkExcel = new Excel.Application();
			Excel.Workbook ObjWorkBook = ObjWorkExcel.Workbooks.Open(ofd.FileName);
			Excel.Worksheet ObjWorkSheet = (Excel.Worksheet)ObjWorkBook.Sheets[1];
			var lastCell = ObjWorkSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell);
			int _columns = (int)lastCell.Column;
			int _rows = (int)lastCell.Row;
			list = new string[_rows, _columns];
			for (int j = 0; j < _columns; j++)
			{
				for (int i = 0; i < _rows; i++)
				{
					list[i, j] = ObjWorkSheet.Cells[i + 1, j + 1].Text;
				}
			}
			int lastRow = 0;
			for (int i = 0; i < _rows; i++)
			{
				if (list[i, 1] != string.Empty)
				{
					lastRow = i;
				}
			}
			ObjWorkBook.Close(false, Type.Missing, Type.Missing);
			ObjWorkExcel.Quit();
			GC.Collect();

			using (ISRPO_2_ParashaEntities3 usersEntities = new ISRPO_2_ParashaEntities3())
			{
				for (int i = 1; i <= lastRow; i++)
				{
					var ID = Convert.ToInt32(list[i, 0]);
					var Kod_Zakaza = list[i, 1];
					var data_cozdaniya = list[i, 2];
					var vremya_zakaza = list[i, 3];
					var kod_clienta = Convert.ToInt32(list[i, 4]);
					var uslugi = list[i, 5];
					var status = list[i, 6];
					var data_zakritiya = list[i, 7];
					var vremya_prokata = list[i, 8];
					string data_zakritiya1;
					//if (list[i,7] == "")
					//{
					//	data_zakritiya1 = " ";
					//}
					//else data_zakritiya1 = list[i, 7];
					var zakaz = new Zakazi()
					{
						ID = Convert.ToInt32(list[i, 0]),
						Kod_Zakaza = list[i, 1],
						data_cozdaniya = list[i, 2],
						vremya_zakaza = list[i, 3],
						kod_clienta = Convert.ToInt32(list[i, 4]),
						uslugi = list[i, 5],
						status = list[i, 6],
						data_zakritiya = list[i, 7],
						vremya_prokata = list[i, 8]
					};
					usersEntities.Zakazi.Add(zakaz);
				}
				usersEntities.SaveChanges();
			}
			MessageBox.Show("Успешное импортирование данных", "Успех", MessageBoxButton.OK, MessageBoxImage.Information);



		}

		private void Button_Click_1(object sender, RoutedEventArgs e)
		{
				try
				{
					var allZakazi = new List<Zakazi>();
					var groups = new[]
					{
							new { Id = 1, Name = "120 минут"},
							new { Id = 2, Name = "600 минут"},
							new { Id = 3, Name = "2 часа"},
							new { Id = 4, Name = "10 часов"},
							new { Id = 5, Name = "320 минут"},
							new { Id = 6, Name = "480 минут"},
							new { Id = 7, Name = "4 часа"},
							new { Id = 8, Name = "6 часов"},
							new { Id = 9, Name = "12 часов"}
					};

					using (ISRPO_2_ParashaEntities3 usersEntities = new ISRPO_2_ParashaEntities3())
					{
						foreach (var user in usersEntities.Zakazi)
						{
							allZakazi.Add(user);
						}
					}

					var app = new Excel.Application();
					app.SheetsInNewWorkbook = groups.Count();
					Excel.Workbook workbook = app.Workbooks.Add(Type.Missing);

					var groupedUsers = allZakazi.GroupBy(user =>
					{
						string vremya_zakaza = user.vremya_prokata;
						if (vremya_zakaza == "120 минут") return 1;
						else if (vremya_zakaza == "600 минут") return 2;
						else if (vremya_zakaza == "2 часа") return 3;
						else if (vremya_zakaza == "10 часов") return 4;
						else if (vremya_zakaza == "320 минут") return 5;
						else if (vremya_zakaza == "480 минут") return 6;
						else if (vremya_zakaza == "4 часа") return 7;
						else if (vremya_zakaza == "6 часов") return 8;
						else if (vremya_zakaza == "12 часов") return 9;
						else return 0;

					});

					for (int i = 0; i < groups.Count(); i++)
					{
						int startRowIndex = 1;
						Excel.Worksheet worksheet = app.Worksheets.Item[i + 1];
						worksheet.Name = Convert.ToString(groups[i].Name);
						worksheet.Cells[1][startRowIndex] = "Id";
						worksheet.Cells[2][startRowIndex] = "Код заказа";
						worksheet.Cells[3][startRowIndex] = "Дата создания";
						worksheet.Cells[4][startRowIndex] = "Код клиента";
						worksheet.Cells[5][startRowIndex] = "Услуги";
					startRowIndex++;
						foreach (var user in groupedUsers.FirstOrDefault(g => g.Key == i + 1))
						{
							worksheet.Cells[1][startRowIndex] = user.ID;
							worksheet.Cells[2][startRowIndex] = user.Kod_Zakaza;
							worksheet.Cells[3][startRowIndex] = user.data_cozdaniya;
							worksheet.Cells[4][startRowIndex] = user.kod_clienta;
							worksheet.Cells[5][startRowIndex] = user.uslugi;
							startRowIndex++;
						}
						worksheet.Columns.AutoFit();
					}
					app.Visible = true;
					MessageBox.Show("Успешное экспортирование данных", "Успех", MessageBoxButton.OK, MessageBoxImage.Information);
				}
				catch (Exception ex)
				{
					MessageBox.Show(ex.Message, "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
				}
			}

		}
	}
