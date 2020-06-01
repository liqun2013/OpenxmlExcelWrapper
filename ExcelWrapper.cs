using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace CommonFunction
{
	public class ExcelWrapper
	{
		public ExcelWrapper()
		{
		}

		#region 写入到Excel相关
		/// <summary>
		/// 生成一个含单个sheet的Excel文件
		/// </summary>
		/// <param name="sheetRowsData">sheet data</param>
		/// <param name="sheetName">sheet name</param>
		/// <param name="filename">excel文件名</param>
		public void GenerateSpreadSheetDoc(SheetRows sheetRowsData, string sheetName, string filename)
		{
			using (SpreadsheetDocument spreadsheetDoc = SpreadsheetDocument.Create(filename, SpreadsheetDocumentType.Workbook))
			{
				WorkbookPart workbookPart = spreadsheetDoc.AddWorkbookPart();
				workbookPart.Workbook = new Workbook();
				WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
				worksheetPart.Worksheet = new Worksheet(CreateSheetData(sheetRowsData));

				Sheets sheets = spreadsheetDoc.WorkbookPart.Workbook.AppendChild(new Sheets());

				Sheet sheet = new Sheet { Id = spreadsheetDoc.WorkbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = sheetName };
				sheets.AppendChild(sheet);

				workbookPart.Workbook.Save();
			}
		}

		/// <summary>
		/// 生成一个含多个sheet的Excel文件
		/// </summary>
		/// <param name="sheetRowsData">多个sheet data的数组</param>
		/// <param name="sheetName">多个sheet name的数组</param>
		/// <param name="filename">excel文件名</param>
		public void GenerateSpreadSheetDoc(List<SheetRows> sheetRowsData, string[] sheetName, string filename)
		{
			using (SpreadsheetDocument spreadsheetDoc = SpreadsheetDocument.Create(filename, SpreadsheetDocumentType.Workbook))
			{
				WorkbookPart workbookPart = spreadsheetDoc.AddWorkbookPart();
				workbookPart.Workbook = new Workbook();
				Sheets sheets = spreadsheetDoc.WorkbookPart.Workbook.AppendChild(new Sheets());
				for (uint i = 0; i < sheetRowsData.Count; i++)
				{
					WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
					worksheetPart.Worksheet = new Worksheet(CreateSheetData(sheetRowsData[(int)i]));
					UInt32Value id = UInt32Value.FromUInt32(i + 1);
					Sheet sheet = new Sheet { Id = spreadsheetDoc.WorkbookPart.GetIdOfPart(worksheetPart), SheetId = id, Name = sheetName[i] };

					sheets.AppendChild(sheet);
				}
				workbookPart.Workbook.Save();
			}
		}


		//public MemoryStream GenerateSpreadSheetDocStream(SheetRows sheetRowsData, string sheetName)
		//{
		//	MemoryStream ms = new MemoryStream();
		//	SpreadsheetDocument spreadsheetDoc = null;
		//	try
		//	{
		//		spreadsheetDoc = SpreadsheetDocument.Create(ms, SpreadsheetDocumentType.Workbook);
		//		WorkbookPart workbookPart = spreadsheetDoc.AddWorkbookPart();
		//		workbookPart.Workbook = new Workbook();
		//		WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
		//		worksheetPart.Worksheet = new Worksheet(CreateSheetData(sheetRowsData));

		//		Sheets sheets = spreadsheetDoc.WorkbookPart.Workbook.AppendChild(new Sheets());

		//		Sheet sheet = new Sheet { Id = spreadsheetDoc.WorkbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = sheetName };
		//		sheets.AppendChild(sheet);

		//		workbookPart.Workbook.Save();
		//	}
		//	catch (InvalidOperationException exc)
		//	{
		//	}
		//	catch (OpenXmlPackageException exc)
		//	{
		//	}
		//	catch (IOException exc)
		//	{
		//	}
		//	catch (Exception exc)
		//	{
		//	}
		//	finally
		//	{
		//		spreadsheetDoc.Dispose();
		//	}

		//	return ms;
		//}

		protected SheetData CreateSheetData(SheetRows sheetRowsData)
		{
			SheetData sheetData = null;

			if (sheetRowsData != null)
			{
				sheetData = new SheetData();
				Row headRow = CreateHeadRow(sheetRowsData.HeadRowData);
				sheetData.AppendChild(headRow);

				if (sheetRowsData.DataRows != null)
				{
					foreach (RowData itm in sheetRowsData.DataRows.OrderBy(x => x.RowIndex))
					{
						Row dataRow = CreateDataRow(itm);
						sheetData.AppendChild(dataRow);
					}
				}
			}

			return sheetData;
		}

		protected Row CreateHeadRow(RowData rowData)
		{
			Row result = null;
			if (rowData != null && rowData.DataItems != null && rowData.DataItems.Any())
			{
				result = new Row { RowIndex = rowData.RowIndex };
				List<DataItem> lstData = rowData.DataItems.OrderBy(x => x.ColIndex).ToList();
				foreach (DataItem itm in lstData)
				{
					result.AppendChild(CreateCell(ColumnLetter(itm.ColIndex), rowData.RowIndex, itm.DataText, DataTypes.String));
				}
			}

			return result;
		}

		protected Row CreateDataRow(RowData rowData)
		{
			Row result = null;
			if (rowData != null && rowData.DataItems != null && rowData.DataItems.Any())
			{
				result = new Row { RowIndex = rowData.RowIndex };
				List<DataItem> lstData = rowData.DataItems.OrderBy(x => x.ColIndex).ToList();
				foreach (DataItem itm in lstData)
				{
					Cell cell = CreateCell(ColumnLetter(itm.ColIndex), rowData.RowIndex, itm.DataText, itm.DataType);
					result.AppendChild(cell);
				}
			}

			return result;
		}

		private Cell CreateCell(string col, uint row, string text, DataTypes dataType)
		{
			switch (dataType)
			{
				case DataTypes.Number:
					{
						return CreateNumberCell(col, row, text);
					}
				case DataTypes.String:
					{
						return CreateTextCell(col, row, text);
					}
				default:
					{
						return null;
					}
			}
		}

		private Cell CreateNumberCell(string col, uint row, string v)
		{
			return new Cell { DataType = CellValues.Number, CellReference = col + row.ToString(), CellValue = new CellValue(v) };
		}

		private Cell CreateTextCell(string col, uint row, string text)
		{
			Cell cell = new Cell { DataType = CellValues.InlineString, CellReference = col + row.ToString() };

			InlineString istring = new InlineString();
			Text t = new Text { Text = string.IsNullOrWhiteSpace(text) ? string.Empty : text };
			istring.AppendChild(t);
			cell.AppendChild(istring);
			return cell;
		}

		private string ColumnLetter(int intCol)
		{
			int intFirstLetter = ((intCol) / 676) + 64;
			int intSecondLetter = ((intCol % 676) / 26) + 64;
			int intThirdLetter = (intCol % 26) + 65;

			char firstLetter = (intFirstLetter > 64) ? (char)intFirstLetter : ' ';
			char secondLetter = (intSecondLetter > 64) ? (char)intSecondLetter : ' ';
			char thirdLetter = (char)intThirdLetter;

			return string.Concat(firstLetter, secondLetter, thirdLetter).Trim();
		}
		#endregion

		#region 读取Excel相关

		public SheetRows ReadSpreadSheetDoc(string filename, bool hasHeadRow)
		{
			using (var stream = File.Open(filename, FileMode.Open))
			{
				return ReadSpreadSheetDoc(stream, hasHeadRow);
			}
		}

		public SheetRows ReadSpreadSheetDoc(Stream stream, bool hasHeadRow)
		{
			SheetRows result = new SheetRows();
			using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(stream, false))
			{
				WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart;
				var theSheet = workbookPart.Workbook.Sheets.FirstChild as Sheet;
				var workSheetPart = workbookPart.GetPartById(theSheet.Id) as WorksheetPart;
				//SheetData sheetData = workSheetPart.Worksheet.FirstOrDefault() as SheetData;
				string cellValue = string.Empty;
				uint rowIndex = 1;
				var lstDataRows = new List<RowData>();
				var rows = workSheetPart.Worksheet.Descendants<Row>();
				foreach (Row r in rows)
				{
					var lstDataItems = new List<DataItem>();
					int colIndex = 1;
					foreach (Cell theCell in r.Elements<Cell>())
					{
						if (theCell.InnerText.Length > 0)
						{
							cellValue = GetCellValue(workbookPart, theCell);
						}
						DataItem headItem = new DataItem { ColIndex = colIndex++, DataText = cellValue, DataType = DataTypes.String };

						lstDataItems.Add(headItem);
					}
					if (rowIndex == 1 && hasHeadRow)
					{
						result.HeadRowData = new RowData { DataItems = lstDataItems, RowIndex = rowIndex };
					}
					else
					{
						lstDataRows.Add(new RowData { DataItems = lstDataItems, RowIndex = rowIndex });
					}
					rowIndex++;
				}
				result.DataRows = lstDataRows;
			}

			return result;
		}

		private string GetCellValue(WorkbookPart workbookPart, Cell theCell)
		{
			string cellValue = theCell.InnerText;
			if (theCell.DataType != null)
			{
				switch (theCell.DataType.Value)
				{
					case CellValues.SharedString:
						var stringTable = workbookPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();
						if (stringTable != null)
						{
							cellValue = stringTable.SharedStringTable.ElementAt(int.Parse(cellValue)).InnerText;
						}
						break;
					case CellValues.Boolean:
						switch (cellValue)
						{
							case "0":
								cellValue = "FALSE";
								break;
							default:
								cellValue = "TRUE";
								break;
						}
						break;
				}
			}

			return cellValue;
		}

		#endregion
	}

	public class SheetRows
	{
		public SheetRows()
		{
		}
		public RowData HeadRowData { get; set; }
		public IEnumerable<RowData> DataRows { get; set; }
	}

	public class RowData
	{
		public uint RowIndex { get; set; }
		public IEnumerable<DataItem> DataItems { get; set; }
	}

	public class DataItem
	{
		public string DataText { get; set; }
		public int ColIndex { get; set; }
		//public int RowIndex { get; set; }
		public DataTypes DataType { get; set; }
	}

	public enum DataTypes
	{
		Number = 1,
		String = 2
	}

}
