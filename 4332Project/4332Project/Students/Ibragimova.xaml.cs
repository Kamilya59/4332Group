using Microsoft.Office.Interop.Excel;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;
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
using Word = Microsoft.Office.Interop.Word;

namespace _4332Project.Students
{
	/// <summary>
	/// Логика взаимодействия для Ibragimova.xaml
	/// </summary>
	public partial class Ibragimova : System.Windows.Window
	{
		public Ibragimova()
		{
			InitializeComponent();
		}

		public class Zakaz
		{

			public Zakaz(int iD, string kod_Zakaza, string data_cozdaniya, string vremya_zakaza, string kod_clienta, string uslugi, string status, string data_zakritiya, string vremya_prokata)
			{
				ID = iD;
				Kod_Zakaza = kod_Zakaza;
				this.data_cozdaniya = data_cozdaniya;
				this.vremya_zakaza = vremya_zakaza;
				this.kod_clienta = kod_clienta;
				this.uslugi = uslugi;
				this.status = status;
				this.data_zakritiya = data_zakritiya;
				this.vremya_prokata = vremya_prokata;
			}

			[JsonPropertyName("Id")]
			public int ID { get; set; }

			[JsonPropertyName("CodeOrder")]
			public string Kod_Zakaza { get; set; }

			[JsonPropertyName("CreateDate")]
			public string data_cozdaniya { get; set; }

			[JsonPropertyName("CreateTime")]
			public string vremya_zakaza { get; set; }

			[JsonPropertyName("CodeClient")]
			public string kod_clienta { get; set; }

			[JsonPropertyName("Services")]
			public string uslugi { get; set; }

			[JsonPropertyName("Status")]
			public string status { get; set; }

			[JsonPropertyName("ClosedDate")]
			public string data_zakritiya { get; set; }

			[JsonPropertyName("ProkatTime")]
			public string vremya_prokata { get; set; }


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
					//if (list[i,7] == "")
					//{
					//	data_zakritiya1 = " ";
					//}
					//else data_zakritiya1 = list[i, 7];
					var zakaz = new Zakazi()
					{
						ID = Convert.ToInt32(list[i, 0]) + 100,
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

		private void JsonImportButton_Click(object sender, RoutedEventArgs e)
		{
			try
			{
				List<Zakaz> zakazi = new List<Zakaz>();
				using (FileStream fs = new FileStream("C:/Users/Камиля/OneDrive/Рабочий стол/2.json", FileMode.OpenOrCreate))
				{
					zakazi = JsonSerializer.Deserialize<List<Zakaz>>(fs);
				}
				using (ISRPO_2_ParashaEntities3 usersEntities = new ISRPO_2_ParashaEntities3())
				{
					foreach (var zakaz in zakazi)
					{
						usersEntities.Zakazi.Add(new Zakazi()
						{
							ID = zakaz.ID,
							uslugi = zakaz.uslugi,
							kod_clienta =  Convert.ToInt32(zakaz.kod_clienta),
							data_cozdaniya = zakaz.data_cozdaniya,
							data_zakritiya = zakaz.data_zakritiya,
							Kod_Zakaza = zakaz.Kod_Zakaza,
							vremya_prokata = zakaz.vremya_prokata,
							status = zakaz.status,
							vremya_zakaza = zakaz.vremya_zakaza


						});
					}
					usersEntities.SaveChanges();
				}
				MessageBox.Show("Успешное импортирование данных", "Успех", MessageBoxButton.OK, MessageBoxImage.Information);
			}
			catch (Exception ex)
			{
				MessageBox.Show(ex.Message, "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
			}
		}

		private void JsonExportButton_Click(object sender, RoutedEventArgs e)
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
				var app = new Word.Application();
				Word.Document document = app.Documents.Add();
				for (int i = 0; i < groups.Count(); i++)
				{
					Word.Paragraph paragraph = document.Paragraphs.Add();
					Word.Range range = paragraph.Range;
					range.Text = Convert.ToString(groups[i].Name);
					paragraph.set_Style("Заголовок 1");
					range.InsertParagraphAfter();

					int count = groupedUsers.FirstOrDefault(g => g.Key == i + 1).Count();

					Word.Paragraph tableParagraph = document.Paragraphs.Add();
					Word.Range tableRange = tableParagraph.Range;
					Word.Table usersTable = document.Tables.Add(tableRange, count + 1, 5);
					usersTable.Borders.InsideLineStyle = usersTable.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
					usersTable.Range.Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
					Word.Range cellRange;


					cellRange = usersTable.Cell(1, 1).Range;
					cellRange.Text = "Id";
					cellRange = usersTable.Cell(1, 2).Range;
					cellRange.Text = "Код заказа";
					cellRange = usersTable.Cell(1, 3).Range;
					cellRange.Text = "Дата создания";
					cellRange = usersTable.Cell(1, 4).Range;
					cellRange.Text = "Код клиента";
					cellRange = usersTable.Cell(1, 5).Range;
					cellRange.Text = "Услуги";

					usersTable.Rows[1].Range.Bold = 1;
					usersTable.Rows[1].Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;

					int j = 1;
					foreach (var user in groupedUsers.FirstOrDefault(g => g.Key == i + 1))
					{
						cellRange = usersTable.Cell(j + 1, 1).Range;
						cellRange.Text = user.ID.ToString();
						cellRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
						cellRange = usersTable.Cell(j + 1, 2).Range;
						cellRange.Text = user.Kod_Zakaza;
						cellRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
						cellRange = usersTable.Cell(j + 1, 3).Range;
						cellRange.Text = user.data_cozdaniya;
						cellRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
						cellRange = usersTable.Cell(j + 1, 4).Range;
						cellRange.Text = user.kod_clienta.ToString();
						cellRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
						cellRange = usersTable.Cell(j + 1, 5).Range;
						cellRange.Text = user.uslugi;
						cellRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
						j++;
					}
					document.Words.Last.InsertBreak(Word.WdBreakType.wdPageBreak);
				}
				app.Visible = true;
				document.SaveAs2("C:/Users/Камиля/OneDrive/Рабочий стол/ParashaWord.docx");
				document.SaveAs2("C:/Users/Камиля/OneDrive/Рабочий стол/ParashaFilePdf.pdf", Word.WdExportFormat.wdExportFormatPDF);
				MessageBox.Show("Успешное экспортирование данных", "Успех", MessageBoxButton.OK, MessageBoxImage.Information);
			}
			catch (Exception ex)
			{
				MessageBox.Show(ex.Message, "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
			}

		}
	}
	}