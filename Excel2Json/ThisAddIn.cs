using System;
using System.Runtime.InteropServices;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using ClosedXML.Excel;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using NPOI.SS.Util;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

public static class Crc32
{
	static uint[] crcTable = MakeCrcTable();

	static uint[] MakeCrcTable()
	{
		uint[] a = new uint[256];
		for (uint i = 0; i < a.Length; i++)
		{
			uint c = i;
			for (int j = 0; j < 8; j++)
			{
				c = ((c & 1) != 0) ? (0xEDB88320 ^ (c >> 1)) : (c >> 1);
			}
			a[i] = c;
		}
		return a;
	}

	public static uint Calculate(byte[] buf)
	{
		return Calculate(buf, 0, buf.Length);
	}

	public static uint Calculate(byte[] buf, int start, int len)
	{
		uint c = 0xFFFFFFFF;
		checked
		{
			if (len < 0)
			{
				throw new ArgumentException();
			}
			if (start < 0 || start + len > buf.Length)
			{
				throw new IndexOutOfRangeException();
			}
		}
		for (int i = 0; i < len; i++)
		{
			c = crcTable[(c ^ buf[start + i]) & 0xFF] ^ (c >> 8);
		}
		return c ^ 0xFFFFFFFF;
	}
}

namespace Excel2Json
{
	public partial class ThisAddIn
	{
		// エクスポート情報
		class ExportInfo
		{
			public string exportFileName; // エクスポートするファイル名
			public string sheetName;      // 対象シート名
			public string formatClassName;// フォーマットのクラス名
		}

		// 参照情報
		class RefInfo
		{
			public string columnName;       // カラム名
			public string refFilePath;      // 参照ファイルパス
			public string refSheet;         // 参照シート
			public string refColumnName;    // 参照カラム表示名
			public string refColumnOutput;  // 参照カラム出力名
		}

		// Jsonフォーマット
		class JsonFormat
		{
			public string formatName;   // フォーマット名
			public string type1;        // 型1
			public string type2;        // 型2
			public string cellName;     // セル名
			public string outputName;   // 出力名
		}

		private bool isExportEnabled = false; // チェックボックスの状態を保存する変数
		private bool isSetupDropDown = false; // ドロップダウンの設定可否
		private string exportFolderPath = @"C:\Users\admin\Documents\Excel2Json"; // エクスポート先のフォルダパス
		private string skipRowName = "//";      // 行のスキップ文字列
		private List<ExportInfo> ExportInfos;   // エクスポート情報
		private Dictionary<string, RefInfo> RefInfos;       // 参照リスト情報
		private Dictionary<string, Dictionary<string, object>> RefDatas; // 参照リストの実データ
		private Dictionary<string, List<JsonFormat>> JsonFormats;
		private List<Dictionary<string, object>> tableData;
		private FileStream fs;
		private FileStream fs_ref;
		private IWorkbook iwb;
		private IXLAddress endAddressClassFormat;   // クラス定義の末尾アドレス
		private IXLAddress endAddressReferenceList; // 参照リストの末尾アドレス
		private IXLAddress endAddressTargetFile;    // データ設定の出力先リストの末尾アドレス

		// チェックボックスの状態を設定する
		public void SetExportEnabled(bool isEnabled)
		{
			isExportEnabled = isEnabled; // チェックボックスの状態をセット
		}

		// エクスポート先のフォルダパスを設定する
		public void SetExportFolderPath(string folderPath)
		{
			exportFolderPath = folderPath;
		}
		private void ThisAddIn_Startup(object sender, System.EventArgs e)
		{
			System.Diagnostics.Debug.WriteLine(RuntimeInformation.FrameworkDescription);

			// WorkbookBeforeCloseとWorkbookBeforeSaveイベントを追加
			this.Application.WorkbookBeforeClose += new Excel.AppEvents_WorkbookBeforeCloseEventHandler(ExportJsonOnClose);
			this.Application.WorkbookAfterSave += new Excel.AppEvents_WorkbookAfterSaveEventHandler(ExportJsonOnSave);
			// ブックが開かれた時のイベントを追加
			this.Application.WorkbookOpen += new Excel.AppEvents_WorkbookOpenEventHandler(ExportJsonOnBookOpen);
			// 保護ビューを閉じた際のイベントを追加
			this.Application.ProtectedViewWindowBeforeClose += new Excel.AppEvents_ProtectedViewWindowBeforeCloseEventHandler(OnProtectedViewWindowBeforeClose);
		}

		// WorkbookBeforeCloseの正しいシグネチャ
		private void ExportJsonOnClose(Excel.Workbook Wb, ref bool Cancel)
		{
			System.Diagnostics.Debug.WriteLine("Workbook is closing."); // イベント発生を確認するためのログ
			if (isExportEnabled && !string.IsNullOrEmpty(exportFolderPath))
			{
				ExportJson(Wb);
			}
		}

		private void ExportJsonOnSave(Excel.Workbook Wb, bool Success)
		{
			System.Diagnostics.Debug.WriteLine("Workbook is saving."); // イベント発生を確認するためのログ
			if (isExportEnabled && !string.IsNullOrEmpty(exportFolderPath))
			{
				ExportJson(Wb);
			}
		}

		// ブックが開かれた時のイベント
		private void ExportJsonOnBookOpen(Excel.Workbook Wb)
		{
			// 出力先を作業中のExcelファイルと同じ場所にする
			SetExportFolderPath(System.IO.Path.GetDirectoryName(Globals.ThisAddIn.Application.ActiveWorkbook.FullName));

			try
			{
				// Excelファイルのフルパス
				string filePath = Wb.FullName;

				if (!isSetupDropDown)
				{
					// NPOIで編集したいので一旦閉じる
					Wb.Close();

					// 開いているExcelを操作できるようにファイルストリームでClosedXMLを操作する
					fs = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);

					using (var xlwb = new XLWorkbook(fs))
					{
						// エクスポート情報の読み込み
						LoadExportInfo(xlwb);
						// 参照リスト情報の読み込み
						LoadRefInfo(xlwb);

						RefDatas = new Dictionary<string, Dictionary<string, object>>();
						foreach (var refInfo in RefInfos)
						{
							// 参照データ情報の読み込み
							LoadRefDataInfo(refInfo);
						}

						// 各シートにドロップダウンを設定する
						foreach (var xlws in xlwb.Worksheets)
						{
							// 特定のシートは設定不要の為、除外する
							if (xlws.Name == "参照リスト")
								continue;

							if (xlws.Name == "データ設定")
								continue;

							SetupDropDown(xlws, filePath);
						}
						using (var local_fs = new FileStream(filePath, FileMode.Create))
						{
							iwb.Write(local_fs);
						}
					}

					if (fs != null)
						fs.Close();

					if (fs_ref != null)
						fs_ref.Close();

					if (!isSetupDropDown)
					{
						isSetupDropDown = true;
						this.Application.Workbooks.Open(filePath);
					}
				}
				else
				{
					// 名前を定義する
					//SetupRangeName(Wb);

					// 前回のチェックボックスの状態を反映
					Properties.Settings setting = Properties.Settings.Default;
					isExportEnabled = setting.checkbox_EtoJson;
					Globals.Ribbons.Ribbon1.checkBox1.Checked = isExportEnabled;
				}
			}
			catch (Exception ex)
			{
				if (fs != null)
					fs.Close();

				if (fs_ref != null)
					fs_ref.Close();

				System.Diagnostics.Debug.WriteLine("エラーが発生しました: " + ex.Message);
			}
		}

		private void OnProtectedViewWindowBeforeClose(ProtectedViewWindow Pvw, XlProtectedViewCloseReason Reason, ref bool Cancel)
		{
			SetExportFolderPath(System.IO.Path.GetDirectoryName(Pvw.Workbook.FullName));
			// 前回のチェックボックスの状態を反映
			Properties.Settings setting = Properties.Settings.Default;
			isExportEnabled = setting.checkbox_EtoJson;
			Globals.Ribbons.Ribbon1.checkBox1.Checked = isExportEnabled;
		}

		// 各シートにドロップダウンを設定する
		private void SetupDropDown(IXLWorksheet xlws, string filePath)
		{
			// ドロップダウンを作成
			if (iwb == null)
				iwb = WorkbookFactory.Create(fs);

			XSSFSheet iws = (XSSFSheet)iwb.GetSheet(xlws.Name);
			foreach (var remove in iws.GetDataValidations())
			{
				iws.RemoveDataValidation(remove);
			}

			foreach (var refInfo in RefInfos)
			{
				var keyCell = xlws.Search(refInfo.Key);
				if (keyCell.Count() <= 0)
				{
					// System.Windows.Forms.MessageBox.Show($"参照リストのカラム名" + refInfo.Key + "は" + xlws.Name + "シートに存在しません");
					continue;
				}

				var columnNumber = -1;
				var keyCell2 = keyCell.Where(n => n.Value.Equals(refInfo.Key)).Select(n => n).ToList();
				if (keyCell2.Count() > 0)
				{
					columnNumber = keyCell2.FirstOrDefault().Address.ColumnNumber;
				}
				if (columnNumber < 0)
					continue;

				foreach (var row in xlws.RangeUsed().RowsUsed())
				{
					foreach (var cell in row.CellsUsed())
					{
						string value = cell.GetValue<string>();

						if (value == refInfo.Key)
							break;

						if (cell.Address.ColumnNumber == columnNumber)
						{
							var data = GetRefData(refInfo.Key);
							if (data == null)
								continue;

							// ドロップダウンに表示するリストを作成
							string[] list = new string[data.Count()];
							var count = 0;
							foreach (var item in data)
							{
								list[count] = item.Key;
								count++;
							}

							// ドロップダウンを作成
							IDataValidationHelper validationHelper = new XSSFDataValidationHelper(iws);
							CellRangeAddressList addressList = new CellRangeAddressList(cell.Address.RowNumber - 1, cell.Address.RowNumber - 1, cell.Address.ColumnNumber - 1, cell.Address.ColumnNumber - 1);
							IDataValidationConstraint constraint = validationHelper.CreateExplicitListConstraint(list);
							IDataValidation dataValidation = validationHelper.CreateValidation(constraint, addressList);
							dataValidation.SuppressDropDownArrow = true;
							iws.AddValidationData(dataValidation);
						}
					}
				}
			}
		}

		// 名前定義のセットアップ
		private void SetupRangeName(Excel.Workbook wb)
		{
			fs = new FileStream(wb.FullName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
			using (var xlwb = new XLWorkbook(fs))
			{
				// ClassFormatの設定
				var sheetName = "データ設定";
				var searchName = "型定義";
				IXLWorksheet xlws;
				if (xlwb.TryGetWorksheet(sheetName, out xlws) && (xlws.Search(searchName).Count() > 0))
				{
					IXLAddress StartAddress = xlws.Search(searchName).FirstOrDefault().Address;
					Excel.Sheets xlsheets = wb.Sheets;
					Excel.Worksheet ws = (Excel.Worksheet)xlsheets[sheetName];
					var addressName = StartAddress.ToString() + ":" + endAddressClassFormat.ToString();
					Range range = (Excel.Range)ws.get_Range(addressName, Type.Missing);
					range.Name = "ClassFormat";
				}
				else
				{
					System.Windows.Forms.MessageBox.Show($"指定されたワークシート名" + sheetName + "が存在しない為、classFormatの定義に失敗");
				}

				// ReferenceListの設定
				sheetName = "参照リスト";
				searchName = "カラム名";
				if (xlwb.TryGetWorksheet(sheetName, out xlws) && (xlws.Search(searchName).Count() > 0))
				{
					IXLAddress StartAddress = xlws.Search(searchName).FirstOrDefault().Address;
					Excel.Sheets xlsheets = wb.Sheets;
					Excel.Worksheet ws = (Excel.Worksheet)xlsheets[sheetName];
					var addressName = StartAddress.ToString() + ":" + endAddressReferenceList.ToString();
					Range range = (Excel.Range)ws.get_Range(addressName, Type.Missing);
					range.Name = "ReferenceList";
				}
				else
				{
					System.Windows.Forms.MessageBox.Show($"指定されたワークシート名" + sheetName + "が存在しない為、ReferenceListの設定の定義に失敗");
				}

				// TargetFileの設定
				sheetName = "データ設定";
				searchName = "出力先";
				if (xlwb.TryGetWorksheet(sheetName, out xlws) && (xlws.Search(searchName).Count() > 0))
				{
					IXLAddress StartAddress = xlws.Search(searchName).FirstOrDefault().Address;
					Excel.Sheets xlsheets = wb.Sheets;
					Excel.Worksheet ws = (Excel.Worksheet)xlsheets[sheetName];
					var addressName = StartAddress.ToString() + ":" + endAddressTargetFile.ToString();
					Range range = (Excel.Range)ws.get_Range(addressName, Type.Missing);
					range.Name = "TargetFile";
				}
				else
				{
					System.Windows.Forms.MessageBox.Show($"指定されたワークシート名" + sheetName + "が存在しない為、TargetFileの設定の定義に失敗");
				}

				// TableHeadの設定
				{
					searchName = "名前";
					foreach (Excel.Worksheet sheet in wb.Worksheets)
					{
						if (sheet.Name == "参照リスト")
							continue;

						if (sheet.Name == "データ設定")
							continue;

						if (xlwb.TryGetWorksheet(sheet.Name, out xlws) && (xlws.Search(searchName).Count() > 0))
						{
							var row2 = xlws.Search(searchName).FirstOrDefault().WorksheetRow();
							var start = row2.CellsUsed().First().Address;
							var end = row2.CellsUsed().Last().Address;
							wb.Worksheets[sheet.Name].Names.Add("TableHead", $"={sheet.Name}!${start.ColumnLetter}${start.ColumnNumber}:${end.ColumnLetter}${end.RowNumber + 1}");
						}
					}
				}
			}

			fs.Close();
		}

		[JsonObject]
		class ExportClass
		{
			[JsonProperty("lists")]
			public List<Equipment> Equipments { get; set; } = new List<Equipment>();
		}

		[JsonObject]
		class Equipment
		{
			[JsonProperty("name")]
			public string name { get; set; }

			[JsonProperty("list")]
			public List<Dictionary<string, object>> Data { get; set; }
		}

		// JSONにエクスポートする
		public void ExportJson(Excel.Workbook Wb)
		{
			try
			{
				// エクスポートフォルダが存在しない場合は作成する
				if (!Directory.Exists(exportFolderPath))
				{
					Directory.CreateDirectory(exportFolderPath);
				}

				// Excelファイルのフルパス
				string filePath = Wb.FullName;
				string fileName = Path.GetFileNameWithoutExtension(filePath); // ファイル名（拡張子なし）

				System.Diagnostics.Debug.WriteLine("Exporting JSON for file: " + filePath);

				// 開いているExcelを操作できるようにファイルストリームでClosedXMLを操作する
				fs = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);

				using (var xlwb = new XLWorkbook(fs))
				{
					Dictionary<string, ExportClass> exportClasses = new Dictionary<string, ExportClass>();

					// エクスポート情報の読み込み
					LoadExportInfo(xlwb);
					// 参照リスト情報の読み込み
					LoadRefInfo(xlwb);

					if (RefDatas == null)
						RefDatas = new Dictionary<string, Dictionary<string, object>>();
					foreach (var refInfo in RefInfos)
					{
						// 参照データ情報の読み込み
						LoadRefDataInfo(refInfo);
					}

					foreach (var info in ExportInfos)
					{
						System.Diagnostics.Debug.WriteLine("ExportInfos start: " + info.sheetName);

						if (info.sheetName == null)
							continue;

						// 対象シートの取得
						IXLWorksheet xlws;
						if (!xlwb.TryGetWorksheet(info.sheetName, out xlws))
						{
							System.Windows.Forms.MessageBox.Show($"指定されたワークシート名" + info.sheetName + "が存在しません。");
							continue;
						}

						// テーブルデータの読み込み
						LoadTableData(xlwb, info.sheetName, info.formatClassName);

						// 型定義に沿ってJsonを出力していく
						if ((info.formatClassName != null) && JsonFormats.ContainsKey(info.formatClassName))
						{
							Equipment equipment = new Equipment();
							equipment.name = info.sheetName;
							equipment.Data = new List<Dictionary<string, object>>();

							var formats = JsonFormats[info.formatClassName];
							List<string> error_msg = new List<string>();
							bool bAddErrorMsg = true;
							foreach (var data in tableData)
							{
								var rowData = new Dictionary<string, object>();

								foreach (var format in formats)
								{
									if ((format.cellName != null) && data.ContainsKey(format.cellName))
									{
										var value = data[format.cellName];

										// CRC32に変換
										if (format.type1 == "hash")
										{
											byte[] t = new System.Text.ASCIIEncoding().GetBytes(value.ToString());
											value = Crc32.Calculate(t);
										}

										// 参照リストのチェック
										Dictionary<string, object> ElementInfos = new Dictionary<string, object>();
										if (RefDatas != null)
										{
											foreach (var Data in RefDatas)
											{
												var num = Data.Key.ToString().IndexOf(',') - 1;
												var name = Data.Key.ToString().Substring(1, num);
												if (name == format.cellName)
												{
													ElementInfos = Data.Value;
													if ((ElementInfos.Count() > 0) && ElementInfos.ContainsKey(value.ToString()))
													{
														value = ElementInfos[value.ToString()].ToString();
													}
													break;
												}
											}
										}

										// サブクラス用のフォーマットに移行
										if (format.type1 == "class")
										{
											if (JsonFormats.ContainsKey(format.type2))
											{
												var subformats = JsonFormats[format.type2];
												var rowSubData = new Dictionary<string, object>();
												foreach (var subformat in subformats)
												{
													if ((subformat.cellName != null) && data.ContainsKey(subformat.cellName))
													{
														value = data[subformat.cellName];
														rowSubData.Add(subformat.outputName, value);
													}
												}
												if (rowSubData.Count() > 0)
												{
													rowData.Add(format.outputName, rowSubData);
												}
												continue;
											}
											else
											{
												continue;
											}
										}

										if (!rowData.ContainsKey(format.outputName))
										{
											rowData.Add(format.outputName, value);
										}
									}
									else
									{
										if (bAddErrorMsg)
										{
											// フォーマット違いをエラーメッセージとして表示する
											if ((format.cellName != null) && !data.ContainsKey(format.cellName))
											{
												error_msg.Add($"- {format.cellName}");
											}
										}
									}
								}

								if (rowData.Count() > 0)
								{
									equipment.Data.Add(rowData);
									bAddErrorMsg = false;
								}
							}

							if (error_msg.Count() > 0)
							{
								System.Windows.Forms.MessageBox.Show($"【フォーマットエラー】 シート名[{info.sheetName}] クラス名[{info.formatClassName}]\n\n" + string.Join("\n", error_msg.ToArray()) + "\n\nのセル名欠けています。");
							}

							ExportClass exportClass = null;
							if (exportClasses.ContainsKey(info.exportFileName))
							{
								exportClass = exportClasses[info.exportFileName];
							}
							else
							{
								exportClass = new ExportClass();
								exportClasses[info.exportFileName] = exportClass;
							}

							exportClass.Equipments.Add(equipment);
						}
					}

					foreach (var exportClass in exportClasses)
					{
						string jsonFilePath = Path.Combine(exportFolderPath, exportClass.Key); // 指定されたフォルダに.jsonファイルを保存するパス
						System.Diagnostics.Debug.WriteLine("Saving JSON to: " + jsonFilePath);

						string jsonStr = JsonConvert.SerializeObject(exportClass.Value, Formatting.Indented);
						System.Diagnostics.Debug.WriteLine("{0}", jsonStr);
						// 新規ファイル作成
						File.WriteAllText(jsonFilePath, jsonStr);
					}
				}

				if (fs != null)
					fs.Close();

				if (fs_ref != null)
					fs_ref.Close();
			}
			catch (Exception ex)
			{
				System.Windows.Forms.MessageBox.Show($"Exportに失敗しました" + ex.Message);

				if (fs != null)
					fs.Close();

				if (fs_ref != null)
					fs_ref.Close();
			}
		}

		// エクスポート情報の読み込み
		private void LoadExportInfo(XLWorkbook xlwb)
		{
			System.Diagnostics.Debug.WriteLine("LoadExportInfo start"); // イベント発生を確認するためのログ

			// 出力ファイル名、対象シート、型の指定データを読み込む
			ExportInfos = new List<ExportInfo>();

			var sheetName = "データ設定";
			IXLWorksheet xlws;
			if (!xlwb.TryGetWorksheet(sheetName, out xlws))
			{
				System.Diagnostics.Debug.WriteLine("指定されたワークシート名" + sheetName + "が存在しません。");
				return;
			}

			// 直近の出力先名
			string lastExportFileName = null;

			// 各セル名のカラム値を取得
			var columnExport = xlws.Search("出力先").FirstOrDefault().Address.ColumnNumber;
			var rowExport = xlws.Search("出力先").FirstOrDefault().Address.RowNumber;
			var columnSheet = xlws.Search("対象シート").FirstOrDefault().Address.ColumnNumber;
			var columnClass = xlws.Search("型").FirstOrDefault().Address.ColumnNumber;
			foreach (var row in xlws.RangeUsed().RowsUsed())
			{
				if (row.RowNumber() <= rowExport)
					continue;

				var exportInfo = new ExportInfo();
				foreach (var cell in row.CellsUsed())
				{
					string cellValue = cell.GetValue<string>();

					if (cellValue == skipRowName)
						break;

					// 出力先カラムの値か
					if (cell.Address.ColumnNumber == columnExport)
					{
						exportInfo.exportFileName = cellValue;
						lastExportFileName = cellValue;
					}

					// 対象シートカラムの値か
					if (cell.Address.ColumnNumber == columnSheet)
						exportInfo.sheetName = cellValue;

					// 型カラムの値か
					if (cell.Address.ColumnNumber == columnClass)
					{
						exportInfo.formatClassName = cellValue;
						endAddressTargetFile = cell.Address;
					}
				}

				// 名前が省略されているのであれば直近の出力先を使う
				if (exportInfo.exportFileName == null)
				{
					exportInfo.exportFileName = lastExportFileName;
				}

				// 有効なデータであれば1行分のパラメータを追加
				if (exportInfo.exportFileName != null)
				{
					ExportInfos.Add(exportInfo);
				}
			}

			// 型定義を読み込む
			JsonFormats = new Dictionary<string, List<JsonFormat>>();

			// 各セル名のカラム値を取得
			var columnDefName = xlws.Search("型（名前）").FirstOrDefault().Address.ColumnNumber;
			var columnDefType1 = xlws.Search("Type").FirstOrDefault().Address.ColumnNumber;
			var columnDefType2 = xlws.Search("SubType").FirstOrDefault().Address.ColumnNumber;
			var columnDefCellName = xlws.Search("セル名").FirstOrDefault().Address.ColumnNumber;
			var columnDefOutputName = xlws.Search("出力名").FirstOrDefault().Address.ColumnNumber;

			// 直近で読み込んだ型(名前)
			string lastDefName = null;
			foreach (var row in xlws.RangeUsed().RowsUsed())
			{
				var jsonFormat = new JsonFormat();

				foreach (var cell in row.CellsUsed())
				{
					string cellValue = cell.GetValue<string>();

					if (cellValue == skipRowName)
						break;

					// 型（名前）カラムの値か
					if (cell.Address.ColumnNumber == columnDefName)
					{
						jsonFormat.formatName = cellValue;
						// 直近で読み込んだ型(名前)を更新
						lastDefName = cellValue;
					}

					// Typeカラムの値か
					if (cell.Address.ColumnNumber == columnDefType1)
						jsonFormat.type1 = cellValue;

					// SubTypeカラムの値か
					if (cell.Address.ColumnNumber == columnDefType2)
						jsonFormat.type2 = cellValue;

					// セル名カラムの値か
					if (cell.Address.ColumnNumber == columnDefCellName)
						jsonFormat.cellName = cellValue;

					// 出力名カラムの値か
					if (cell.Address.ColumnNumber == columnDefOutputName)
					{
						jsonFormat.outputName = cellValue;
						if ((jsonFormat.type1 != null) && (jsonFormat.cellName != null))
						{
							endAddressClassFormat = cell.Address;
						}
					}
				}

				// 型（名前）が省略されていれば直近で読み込んだものを設定しておく
				if (jsonFormat.formatName == null)
					jsonFormat.formatName = lastDefName;

				// 有効なデータであれば1行分のパラメータを追加
				if ((jsonFormat.formatName != null) && (jsonFormat.type1 != null) && (jsonFormat.cellName != null) && (jsonFormat.outputName != null))
				{
					if (!JsonFormats.ContainsKey(jsonFormat.formatName))
					{
						var newFormat = new List<JsonFormat>();
						JsonFormats[jsonFormat.formatName] = newFormat;
					}
					JsonFormats[jsonFormat.formatName].Add(jsonFormat);
				}
			}
		}

		// 参照情報の読み込み
		private void LoadRefInfo(XLWorkbook xlwb)
		{
			System.Diagnostics.Debug.WriteLine("LoadRefInfo start"); // イベント発生を確認するためのログ

			// 出力ファイル名、対象シート、型の指定データを読み込む
			RefInfos = new Dictionary<string, RefInfo>();

			var sheetName = "参照リスト";
			IXLWorksheet xlws;
			if (!xlwb.TryGetWorksheet(sheetName, out xlws))
			{
				System.Diagnostics.Debug.WriteLine("指定されたワークシート名" + sheetName + "が存在しません。");
				return;
			}

			// 各セル名のカラム値を取得
			var columnName = xlws.Search("カラム名").FirstOrDefault().Address.ColumnNumber;
			var rowName = xlws.Search("カラム名").FirstOrDefault().Address.RowNumber;
			var columnRefFilePath = xlws.Search("参照ファイル").FirstOrDefault().Address.ColumnNumber;
			var columnSheet = xlws.Search("参照シート").FirstOrDefault().Address.ColumnNumber;
			var columnVisibleName = xlws.Search("参照カラム表示名").FirstOrDefault().Address.ColumnNumber;
			var columnOutputName = xlws.Search("参照カラム出力名").FirstOrDefault().Address.ColumnNumber;
			foreach (var row in xlws.RangeUsed().RowsUsed())
			{
				if (row.RowNumber() <= rowName)
					continue;

				var refInfo = new RefInfo();

				foreach (var cell in row.CellsUsed())
				{
					string cellValue = cell.GetValue<string>();

					if (cellValue == skipRowName)
						break;

					// カラム名カラムの値か
					if (cell.Address.ColumnNumber == columnName)
					{
						refInfo.columnName = cellValue;
					}

					// 参照ファイルカラムの値か
					if (cell.Address.ColumnNumber == columnRefFilePath)
						refInfo.refFilePath = cellValue;

					// 参照シートカラムの値か
					if (cell.Address.ColumnNumber == columnSheet)
						refInfo.refSheet = cellValue;

					// 参照カラム表示名カラムの値か
					if (cell.Address.ColumnNumber == columnVisibleName)
						refInfo.refColumnName = cellValue;

					// 参照カラム出力名カラムの値か
					if (cell.Address.ColumnNumber == columnOutputName)
					{
						refInfo.refColumnOutput = cellValue;
						endAddressReferenceList = cell.Address;
					}
				}

				// 有効なデータであれば1行分のパラメータを追加
				if ((refInfo.columnName != null) && (refInfo.refFilePath != null) && (refInfo.refSheet != null))
				{
					RefInfos[refInfo.columnName] = refInfo;
				}
			}
		}

		// 参照データ情報の読み込み
		private void LoadRefDataInfo(KeyValuePair<string, RefInfo> info)
		{
			System.Diagnostics.Debug.WriteLine("LoadRefDataInfo start ");

			var RefInfo = info.Value;

			// Excelファイルのフルパス
			string filePath = exportFolderPath + RefInfo.refFilePath;
			string fileName = Path.GetFileNameWithoutExtension(RefInfo.refFilePath); // ファイル名（拡張子なし）

			System.Diagnostics.Debug.WriteLine("Exporting JSON for file: " + filePath);

			if (!File.Exists(filePath))
			{
				System.Windows.Forms.MessageBox.Show($"指定されたファイルパス[" + filePath + "]は存在しません。");
				return;
			}

			// 開いているExcelを操作できるようにファイルストリームでClosedXMLを操作する
			fs_ref = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
			Dictionary<string, object> Datas = new Dictionary<string, object>();
			using (var xlwb = new XLWorkbook(fs_ref))
			{
				var sheetName = RefInfo.refSheet;
				IXLWorksheet xlws;
				if (!xlwb.TryGetWorksheet(sheetName, out xlws))
				{
					System.Windows.Forms.MessageBox.Show($"指定されたワークシート名" + RefInfo.refFilePath + " : " + sheetName + "が存在しません。");
					fs_ref.Close();
					return;
				}

				var keyCell = xlws.Search(RefInfo.refColumnName);
				// 各セル名のカラム値を取得
				if (keyCell.Count() <= 0)
				{
					fs_ref.Close();
					return;
				}
				var columnName = keyCell.FirstOrDefault().Address.ColumnNumber;

				keyCell = xlws.Search(RefInfo.refColumnOutput);
				if (keyCell.Count() <= 0)
				{
					fs_ref.Close();
					return;
				}
				var columnOutput = keyCell.FirstOrDefault().Address.ColumnNumber;
				foreach (var row in xlws.RangeUsed().RowsUsed())
				{
					foreach (var cell in row.CellsUsed())
					{
						string cellValue = cell.GetValue<string>();

						if (cellValue == skipRowName)
							break;

						// カラム名カラムの値か
						if (cell.Address.ColumnNumber == columnName)
						{
							Datas[cellValue] = row.Cell(columnOutput - 1).Value;
						}
					}
				}
			}

			if (Datas.Count() > 0)
			{
				RefDatas[info.ToString()] = Datas;
			}

			fs_ref.Close();
		}

		// 指定Keyの参照データを取得する
		private Dictionary<string, object> GetRefData(string name)
		{
			if (RefDatas != null)
			{
				foreach (var Data in RefDatas)
				{
					var num = Data.Key.ToString().IndexOf(',') - 1;
					var key = Data.Key.ToString().Substring(1, num);
					if (key == name)
					{
						return Data.Value;
					}
				}
			}

			return null;
		}

		// テーブルデータの読み込み
		private void LoadTableData(XLWorkbook xlwb, string sheetName, string formatClassName)
		{
			System.Diagnostics.Debug.WriteLine("LoadTableData start: ");

			if (sheetName == null)
				return;

			// 対象シートの取得
			IXLWorksheet xlws;
			if (!xlwb.TryGetWorksheet(sheetName, out xlws))
			{
				System.Diagnostics.Debug.WriteLine("指定されたワークシート名" + sheetName + "が存在しません。");
				return;
			}

			bool checkMainClass = false;
			bool checkSubClass = false;
			bool existSubClass = false;
			Dictionary<int, Dictionary<string, object>> tempTableData = new Dictionary<int, Dictionary<string, object>>();

			// セル名が記載されている行番号を取得して、その行の各セルのカラムをキャッシュ
			Dictionary<int, string> mainColumnNames = new Dictionary<int, string>();
			var range = xlws.Range("TableHead");
			var row_number = range.FirstRow().RowNumber();
			//var row_number = xlws.Search("名前").FirstOrDefault().Address.RowNumber;
			var row_names = xlws.Row(row_number);
			foreach (var cell in row_names.CellsUsed())
			{
				string value = cell.GetValue<string>();
				mainColumnNames[cell.Address.ColumnNumber] = value;
				if (value == "サブクラス")
				{
					existSubClass = true;
				}
			}

			foreach (var row in xlws.RangeUsed().RowsUsed())
			{
				Dictionary<string, object> rowData = new Dictionary<string, object>();
				foreach (var cell in row.CellsUsed())
				{
					var Value = cell.GetValue<string>();
					if (Value == skipRowName)
					{
						break;
					}

					// 値に対応した名前をセットでキャッシュ
					if (mainColumnNames.ContainsKey(cell.Address.ColumnNumber))
					{
						rowData[mainColumnNames[cell.Address.ColumnNumber]] = cell.GetValue<string>();
					}
				}

				if (rowData.Count() > 0)
				{
					if (!checkMainClass)
					{
						checkMainClass = true;
						checkFormatData(sheetName, formatClassName, rowData, true);
					}
					tempTableData[row.RowNumber()] = rowData;
				}
			}

			// サブクラス用の名前も同様に処理する
			Dictionary<int, string> subColumnNames = new Dictionary<int, string>();
			row_names = xlws.Row(row_number + 1);
			foreach (var cell in row_names.CellsUsed())
			{
				subColumnNames[cell.Address.ColumnNumber] = cell.GetValue<string>();
			}

			Dictionary<string, object> rowSubData = new Dictionary<string, object>();
			if (existSubClass)
			{
				foreach (var row in xlws.RangeUsed().RowsUsed())
				{
					foreach (var cell in row.CellsUsed())
					{
						var Value = cell.GetValue<string>();
						if (Value == skipRowName)
						{
							break;
						}

						if (subColumnNames.ContainsKey(cell.Address.ColumnNumber))
						{
							if (tempTableData.ContainsKey(row.RowNumber()))
							{
								tempTableData[row.RowNumber()][subColumnNames[cell.Address.ColumnNumber]] = cell.GetValue<string>();
								rowSubData[subColumnNames[cell.Address.ColumnNumber]] = cell.GetValue<string>();
							}
						}
					}
				}
			}

			if (!checkSubClass)
			{
				checkSubClass = true;
				if (JsonFormats.ContainsKey(formatClassName))
				{
					var format = JsonFormats[formatClassName];
					foreach (var item in format)
					{
						if ((item.type1 == "class") && (item.type2 != null))
						{
							checkFormatData(sheetName, item.type2, rowSubData, false);
							break;
						}
					}
				}
			}

			tableData = new List<Dictionary<string, object>>();
			foreach (var temp in tempTableData)
			{
				tableData.Add(temp.Value);
			}
		}

		// パラメーターシートにある項目がデータ設定シートに存在しない場合のエラーメッセージ対応
		private void checkFormatData(string sheetName, string formatClassName, Dictionary<string, object> rowData, bool checkExportMsg)
		{
			if (checkExportMsg && !JsonFormats.ContainsKey(formatClassName))
			{
				System.Windows.Forms.MessageBox.Show($"{sheetName}シートはExportに失敗しました");
				return;
			}

			if (!JsonFormats.ContainsKey(formatClassName))
			{
				System.Windows.Forms.MessageBox.Show($"{formatClassName}クラスは存在しません");
				return;
			}
#if false  // データ設定シートに無いカラムが存在しても問題ない       
			foreach (var data in rowData)
			{
				var bisFind = false;
				foreach (var format in JsonFormats[formatClassName])
				{
					if (data.Key == format.cellName)
					{
						bisFind = true;
						break;
					}
				}
                if (!bisFind)
                {
                    System.Windows.Forms.MessageBox.Show($"{sheetName}シートの{data.Key}は{formatClassName}に存在しません。");
                }
			}
#endif
		}

		private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
		{
			// チェックボックスの状態を保存
			Properties.Settings setting = Properties.Settings.Default;
			setting.checkbox_EtoJson = isExportEnabled;
			setting.Save();
		}

		#region VSTO で生成されたコード

		/// <summary>
		/// デザイナーのサポートに必要なメソッドです。
		/// コード エディターで変更しないでください。
		/// </summary>
		private void InternalStartup()
		{
			this.Startup += new System.EventHandler(ThisAddIn_Startup);
			this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
		}

		#endregion
	}
}
