using System;
using System.IO;
using System.Data;
using System.Linq;
using System.Drawing;
using System.Reflection;
using System.Collections.Generic;
using OfficeOpenXml;
using OfficeOpenXml.Table;
using OfficeOpenXml.Style;

namespace ITGrass.Common.Utility.Excel
{
	public static class CollectionExcelExtensions
	{
		/// <summary>
		/// 输出Excel, 超链接设置方式为遍历所有单元格
		/// </summary>
		/// <typeparam name="T"></typeparam>
		/// <param name="datas"></param>
		/// <param name="mapDictionary">key为excel模型字段名，value为标题栏</param>
		/// <param name="hasHyperlink">是否有超链接</param>
		/// <param name="optionsAction">选项设置</param>
		/// <returns></returns>
		public static byte[] ToExcel<T>(this IEnumerable<T> datas, Dictionary<string, string> mapDictionary = null,
			bool hasHyperlink = false, Action<ToExcelOptions> optionsAction = null)
		{
			var options = new ToExcelOptions();
			optionsAction?.Invoke(options);

			using (var package = new ExcelPackage())
			{
				var ws = package.Workbook.Worksheets.Add("数据");
				ws.Cells[1, 1].LoadFromCollection(datas, true, options.TableStyles);
				var hiddenCols = new List<int>();
				if (mapDictionary != null)
				{
					//设置excel列标题, 并获取需要隐藏的列
					for (int i = 1; i <= ws.Dimension.Columns; i++)
					{
						var cell = ws.Cells[1, i];
						if (mapDictionary.TryGetValue(cell.GetValue<string>().Replace(' ', '_'), out string name))
						{
							cell.Value = name;
						}
						else
						{
							hiddenCols.Add(i);
						}
					}

					//设置列隐藏
					foreach (var item in hiddenCols)
					{
						ws.Column(item).Hidden = true;
					}
				}

				//设置单元格格式
				var type = typeof(T);
				var index = 0;
				var rowCount = datas.Count();
				foreach (var prop in type.GetProperties())
				{
					index++;

					//设置日期显示格式
					if (prop.PropertyType == typeof(DateTime))
						ws.Column(index).Style.Numberformat.Format = options.DatetimeFormate;

					//设置单元格超链接
					if (hasHyperlink)
					{
						for (var count = 1; count <= rowCount; count++)
						{
							var value = ws.Cells[count + 1, index].Value?.ToString();
							if (!string.IsNullOrWhiteSpace(value) &&
							    Uri.TryCreate(value, UriKind.Absolute, out Uri uri))
							{
								ws.Cells[count + 1, index].Hyperlink = uri;
							}
						}
					}
				}

				ws.Cells[ws.Dimension.Address].AutoFitColumns();
				return package.GetAsByteArray();
			}
		}

		/// <summary>
		/// 输出Excel, 超链接设置方式为遍历指定的列
		/// </summary>
		/// <typeparam name="T"></typeparam>
		/// <param name="datas"></param>
		/// <param name="mapDictionary">key为excel模型字段名，value为标题栏</param>
		/// <param name="hyperlinkFields">超链接字段集合</param>
		/// <param name="optionsAction">选项设置</param>
		/// <returns></returns>
		public static byte[] ToExcelAppoint<T>(this IEnumerable<T> datas,
			Dictionary<string, string> mapDictionary = null, List<string> hyperlinkFields = null,
			Action<ToExcelOptions> optionsAction = null)
		{
			var options = new ToExcelOptions();
			optionsAction?.Invoke(options);

			using (var package = new ExcelPackage())
			{
				var ws = package.Workbook.Worksheets.Add("数据");
				ws.Cells[1, 1].LoadFromCollection(datas, true, options.TableStyles);
				var hiddenCols = new List<int>();
				if (mapDictionary != null)
				{
					//设置excel列标题, 并获取需要隐藏的列
					for (int i = 1; i <= ws.Dimension.Columns; i++)
					{
						var cell = ws.Cells[1, i];
						if (mapDictionary.TryGetValue(cell.GetValue<string>().Replace(' ', '_'), out string name))
						{
							cell.Value = name;
						}
						else
						{
							hiddenCols.Add(i);
						}
					}

					//设置列隐藏
					foreach (var item in hiddenCols)
					{
						ws.Column(item).Hidden = true;
					}
				}

				//设置单元格格式 
				var type = typeof(T);
				var index = 0;
				var rowCount = datas.Count();
				foreach (var prop in type.GetProperties())
				{
					index++;

					//设置日期显示格式
					if (prop.PropertyType == typeof(DateTime))
						ws.Column(index).Style.Numberformat.Format = options.DatetimeFormate;

					//设置单元格超链接
					if (hyperlinkFields != null && hyperlinkFields.Any(a => a.ToLower() == prop.Name.ToLower()))
					{
						for (var count = 1; count <= rowCount; count++)
						{
							var value = ws.Cells[count + 1, index].Value?.ToString();
							if (!string.IsNullOrWhiteSpace(value) &&
							    Uri.TryCreate(value, UriKind.Absolute, out Uri uri))
							{
								ws.Cells[count + 1, index].Hyperlink = uri;
							}
						}
					}
				}

				ws.Cells[ws.Dimension.Address].AutoFitColumns();
				return package.GetAsByteArray();
			}
		}

		/// <summary>
		/// 输出Excel, 根据字典顺序排序, 超链接设置方式为遍历所有单元格
		/// </summary>
		/// <typeparam name="T"></typeparam>
		/// <param name="datas"></param>
		/// <param name="mapDictionary">key为excel模型字段名，value为标题栏</param>
		/// <param name="hasHyperlink">是否有超链接</param>
		/// <param name="optionsAction">选项设置</param>
		/// <returns></returns>
		public static byte[] ToExcelSort<T>(this IEnumerable<T> datas, Dictionary<string, string> mapDictionary,
			bool hasHyperlink = false, Action<ToExcelOptions> optionsAction = null)
		{
			var options = new ToExcelOptions();
			optionsAction?.Invoke(options);

			//创建DataTable
			var dTable = CreateDataTable(datas, mapDictionary);

			using (var package = new ExcelPackage())
			{
				var ws = package.Workbook.Worksheets.Add("数据");
				ws.Cells[1, 1].LoadFromDataTable(dTable, true, options.TableStyles);

				//设置单元格格式 
				var type = typeof(T);
				var index = 0;
				var rowCount = datas.Count();
				foreach (var item in mapDictionary)
				{
					var prop = type.GetProperty(item.Key,
						BindingFlags.Instance | BindingFlags.Public | BindingFlags.IgnoreCase);
					if (prop == null)
						continue;

					index++;

					//设置日期显示格式
					if (prop.PropertyType == typeof(DateTime))
						ws.Column(index).Style.Numberformat.Format = options.DatetimeFormate;

					//设置单元格超链接
					if (hasHyperlink)
					{
						for (var count = 1; count <= rowCount; count++)
						{
							var value = ws.Cells[count + 1, index].Value?.ToString();
							if (!string.IsNullOrWhiteSpace(value) &&
							    Uri.TryCreate(value, UriKind.Absolute, out Uri uri))
							{
								ws.Cells[count + 1, index].Hyperlink = uri;
							}
						}
					}
				}

				ws.Cells[ws.Dimension.Address].AutoFitColumns();
				return package.GetAsByteArray();
			}
		}

		/// <summary>
		/// 输出Excel, 根据字典顺序排序, 超链接设置方式为遍历所有单元格
		/// 根据指定唯一列计算合并列, 按照字典顺序, 从第一列开始, 计算合并mergeLength指定长度的单元格
		/// </summary>
		/// <typeparam name="T"></typeparam>
		/// <param name="datas"></param>
		/// <param name="mapDictionary">key为excel模型字段名，value为标题栏</param>
		/// <param name="uniqueColumnName">合并唯一标识列</param>
		/// <param name="mergeLength">要合并列的长度</param>
		/// <param name="hasHyperlink">是否有超链接</param>
		/// <param name="optionsAction">选项设置</param>
		/// <returns></returns>
		public static byte[] ToExcelSortMerge<T>(this IEnumerable<T> datas, Dictionary<string, string> mapDictionary,
			string uniqueColumnName, int mergeLength, bool hasHyperlink = false,
			Action<ToExcelOptions> optionsAction = null)
		{
			var options = new ToExcelOptions();
			optionsAction?.Invoke(options);

			//创建DataTable
			var dTable = CreateDataTable(datas, mapDictionary);

			using (var package = new ExcelPackage())
			{
				var ws = package.Workbook.Worksheets.Add("数据");
				ws.Cells[1, 1].LoadFromDataTable(dTable, true);

				//设置单元格格式 
				var type = typeof(T);
				var index = 0;
				var rowCount = datas.Count();
				foreach (var item in mapDictionary)
				{
					var prop = type.GetProperty(item.Key,
						BindingFlags.Instance | BindingFlags.Public | BindingFlags.IgnoreCase);
					if (prop == null)
						continue;

					index++;

					//设置日期显示格式
					if (prop.PropertyType == typeof(DateTime))
						ws.Column(index).Style.Numberformat.Format = options.DatetimeFormate;

					//设置单元格超链接
					if (hasHyperlink)
					{
						for (var count = 1; count <= rowCount; count++)
						{
							var value = ws.Cells[count + 1, index].Value?.ToString();
							if (!string.IsNullOrWhiteSpace(value) &&
							    Uri.TryCreate(value, UriKind.Absolute, out Uri uri))
							{
								ws.Cells[count + 1, index].Hyperlink = uri;
							}
						}
					}
				}

				//合并单元格
				MergeSetting(ws, dTable, mapDictionary, uniqueColumnName, mergeLength);

				ws.Cells[ws.Dimension.Address].AutoFitColumns();
				return package.GetAsByteArray();
			}
		}

		/// <summary>
		/// 输出Excel, 根据字典顺序排序, 超链接设置方式为遍历指定的列
		/// </summary>
		/// <typeparam name="T"></typeparam>
		/// <param name="datas"></param>
		/// <param name="mapDictionary">key为excel模型字段名，value为标题栏</param>
		/// <param name="hyperlinkFields">超链接字段集合</param>
		/// <param name="optionsAction">选项设置</param>
		/// <returns></returns>
		public static byte[] ToExcelAppointSort<T>(this IEnumerable<T> datas, Dictionary<string, string> mapDictionary,
			List<string> hyperlinkFields = null, Action<ToExcelOptions> optionsAction = null)
		{
			var options = new ToExcelOptions();
			optionsAction?.Invoke(options);

			//创建DataTable
			var dTable = CreateDataTable(datas, mapDictionary);

			using (var package = new ExcelPackage())
			{
				var ws = package.Workbook.Worksheets.Add("数据");
				ws.Cells[1, 1].LoadFromDataTable(dTable, true, options.TableStyles);

				//设置单元格格式 
				var type = typeof(T);
				var index = 0;
				var rowCount = datas.Count();
				foreach (var item in mapDictionary)
				{
					var prop = type.GetProperty(item.Key,
						BindingFlags.Instance | BindingFlags.Public | BindingFlags.IgnoreCase);
					if (prop == null)
						continue;

					index++;

					//设置日期显示格式
					if (prop.PropertyType == typeof(DateTime))
						ws.Column(index).Style.Numberformat.Format = options.DatetimeFormate;

					//设置单元格超链接
					if (hyperlinkFields != null && hyperlinkFields.Any(a => a.ToLower() == prop.Name.ToLower()))
					{
						for (var count = 1; count <= rowCount; count++)
						{
							var value = ws.Cells[count + 1, index].Value?.ToString();
							if (!string.IsNullOrWhiteSpace(value) &&
							    Uri.TryCreate(value, UriKind.Absolute, out Uri uri))
							{
								ws.Cells[count + 1, index].Hyperlink = uri;
							}
						}
					}
				}

				ws.Cells[ws.Dimension.Address].AutoFitColumns();
				return package.GetAsByteArray();
			}
		}

		/// <summary>
		/// 输出Excel, 根据字典顺序排序, 超链接设置方式为遍历指定的列
		/// 根据指定唯一列计算合并列, 按照字典顺序, 从第一列开始, 计算合并mergeLength指定长度的单元格
		/// </summary>
		/// <typeparam name="T"></typeparam>
		/// <param name="datas"></param>
		/// <param name="mapDictionary">key为excel模型字段名，value为标题栏</param>
		/// <param name="uniqueColumnName">合并唯一标识列</param>
		/// <param name="mergeLength">要合并列的长度</param>
		/// <param name="hyperlinkFields">超链接字段集合</param>
		/// <param name="optionsAction">选项设置</param>
		/// <returns></returns>
		public static byte[] ToExcelAppointSortMerge<T>(this IEnumerable<T> datas,
			Dictionary<string, string> mapDictionary, string uniqueColumnName, int mergeLength,
			List<string> hyperlinkFields = null, Action<ToExcelOptions> optionsAction = null)
		{
			var options = new ToExcelOptions();
			optionsAction?.Invoke(options);

			//创建DataTable
			var dTable = CreateDataTable(datas, mapDictionary);

			using (var package = new ExcelPackage())
			{
				var ws = package.Workbook.Worksheets.Add("数据");
				ws.Cells[1, 1].LoadFromDataTable(dTable, true);

				//设置单元格格式 
				var type = typeof(T);
				var index = 0;
				var rowCount = datas.Count();
				foreach (var item in mapDictionary)
				{
					var prop = type.GetProperty(item.Key,
						BindingFlags.Instance | BindingFlags.Public | BindingFlags.IgnoreCase);
					if (prop == null)
						continue;

					index++;

					//设置日期显示格式
					if (prop.PropertyType == typeof(DateTime))
						ws.Column(index).Style.Numberformat.Format = options.DatetimeFormate;

					//设置单元格超链接
					if (hyperlinkFields != null && hyperlinkFields.Any(a => a.ToLower() == prop.Name.ToLower()))
					{
						for (var count = 1; count <= rowCount; count++)
						{
							var value = ws.Cells[count + 1, index].Value?.ToString();
							if (!string.IsNullOrWhiteSpace(value) &&
							    Uri.TryCreate(value, UriKind.Absolute, out Uri uri))
							{
								ws.Cells[count + 1, index].Hyperlink = uri;
							}
						}
					}
				}

				//合并单元格
				MergeSetting(ws, dTable, mapDictionary, uniqueColumnName, mergeLength);

				ws.Cells[ws.Dimension.Address].AutoFitColumns();
				return package.GetAsByteArray();
			}
		}

		/// <summary>
		/// 输出Excel, 根据字典顺序排序, 支持单属性字段为(字符串)集合的对象集合, 超链接设置方式为遍历所有单元格
		/// </summary>
		/// <typeparam name="T"></typeparam>
		/// <param name="datas"></param>
		/// <param name="collectionFieldName">集合属性字段名称</param>
		/// <param name="collectionFieldMaxLength">集合属性字段的最大元素个数</param>
		/// <param name="mapDictionary">key为excel模型字段名，value为标题栏</param>
		/// <param name="hasHyperlink">是否有超链接</param>
		/// <param name="optionsAction">选项设置</param>
		/// <returns></returns>
		public static byte[] ToExcelWithCollection<T>(this IEnumerable<T> datas, string collectionFieldName,
			int collectionFieldMaxLength, Dictionary<string, string> mapDictionary, bool hasHyperlink = false,
			Action<ToExcelOptions> optionsAction = null)
		{
			var options = new ToExcelOptions();
			optionsAction?.Invoke(options);

			//创建DataTable
			var dTable = CreateDataTable(datas, collectionFieldName, collectionFieldMaxLength, mapDictionary);

			using (var package = new ExcelPackage())
			{
				var ws = package.Workbook.Worksheets.Add("数据");
				ws.Cells[1, 1].LoadFromDataTable(dTable, true, options.TableStyles);

				//设置单元格格式 
				var type = typeof(T);
				var index = 0;
				var rowCount = datas.Count();
				var cycleCount = 0;
				foreach (var item in mapDictionary)
				{
					var prop = type.GetProperty(item.Key,
						BindingFlags.Instance | BindingFlags.Public | BindingFlags.IgnoreCase);
					if (prop == null)
						continue;

					if (prop.Name.ToLower() == collectionFieldName.ToLower())
					{
						index += collectionFieldMaxLength;
						cycleCount = collectionFieldMaxLength - 1;
					}
					else
					{
						index++;
						cycleCount = 0;
					}

					//设置日期显示格式
					if (prop.PropertyType == typeof(DateTime))
						ws.Column(index).Style.Numberformat.Format = options.DatetimeFormate;

					//设置单元格超链接
					if (hasHyperlink)
					{
						for (var count = 1; count <= rowCount; count++)
						{
							for (var i = cycleCount; i >= 0; i--)
							{
								var value = ws.Cells[count + 1, index - i].Value?.ToString();
								if (!string.IsNullOrWhiteSpace(value) &&
								    Uri.TryCreate(value, UriKind.Absolute, out Uri uri))
								{
									ws.Cells[count + 1, index - i].Hyperlink = uri;
								}
							}
						}
					}
				}

				ws.Cells[ws.Dimension.Address].AutoFitColumns();
				return package.GetAsByteArray();
			}
		}

		/// <summary>
		/// 输出Excel, 根据字典顺序排序, 支持单属性字段为(字符串)集合的对象集合, 超链接设置方式为遍历所有单元格
		/// 根据指定唯一列计算合并列, 按照字典顺序, 从第一列开始, 计算合并mergeLength指定长度的单元格
		/// </summary>
		/// <typeparam name="T"></typeparam>
		/// <param name="datas"></param>
		/// <param name="collectionFieldName">集合属性字段名称</param>
		/// <param name="collectionFieldMaxLength">集合属性字段的最大元素个数</param>
		/// <param name="mapDictionary">key为excel模型字段名，value为标题栏</param>
		/// <param name="uniqueColumnName">合并唯一标识列</param>
		/// <param name="mergeLength">要合并列的长度</param>
		/// <param name="hasHyperlink">是否有超链接</param>
		/// <param name="optionsAction">选项设置</param>
		/// <returns></returns>
		public static byte[] ToExcelWithCollectionMerge<T>(this IEnumerable<T> datas, string collectionFieldName,
			int collectionFieldMaxLength, Dictionary<string, string> mapDictionary, string uniqueColumnName,
			int mergeLength, bool hasHyperlink = false, Action<ToExcelOptions> optionsAction = null)
		{
			var options = new ToExcelOptions();
			optionsAction?.Invoke(options);

			//创建DataTable
			var dTable = CreateDataTable(datas, collectionFieldName, collectionFieldMaxLength, mapDictionary);

			using (var package = new ExcelPackage())
			{
				var ws = package.Workbook.Worksheets.Add("数据");
				ws.Cells[1, 1].LoadFromDataTable(dTable, true);

				//设置单元格格式 
				var type = typeof(T);
				var index = 0;
				var rowCount = datas.Count();
				var cycleCount = 0;
				foreach (var item in mapDictionary)
				{
					var prop = type.GetProperty(item.Key,
						BindingFlags.Instance | BindingFlags.Public | BindingFlags.IgnoreCase);
					if (prop == null)
						continue;

					if (prop.Name.ToLower() == collectionFieldName.ToLower())
					{
						index += collectionFieldMaxLength;
						cycleCount = collectionFieldMaxLength - 1;
					}
					else
					{
						index++;
						cycleCount = 0;
					}

					//设置日期显示格式
					if (prop.PropertyType == typeof(DateTime))
						ws.Column(index).Style.Numberformat.Format = options.DatetimeFormate;

					//设置单元格超链接
					if (hasHyperlink)
					{
						for (var count = 1; count <= rowCount; count++)
						{
							for (var i = cycleCount; i >= 0; i--)
							{
								var value = ws.Cells[count + 1, index - i].Value?.ToString();
								if (!string.IsNullOrWhiteSpace(value) &&
								    Uri.TryCreate(value, UriKind.Absolute, out Uri uri))
								{
									ws.Cells[count + 1, index - i].Hyperlink = uri;
								}
							}
						}
					}
				}

				//合并单元格
				MergeSetting(ws, dTable, mapDictionary, uniqueColumnName, mergeLength, 1, collectionFieldMaxLength);

				ws.Cells[ws.Dimension.Address].AutoFitColumns();
				return package.GetAsByteArray();
			}
		}

		/// <summary>
		/// 输出Excel, 根据字典顺序排序, 支持单属性字段为(字符串)集合的对象集合, 超链接设置方式为遍历指定的列
		/// </summary>
		/// <typeparam name="T"></typeparam>
		/// <param name="datas"></param>
		/// <param name="collectionFieldName">集合属性字段名称</param>
		/// <param name="collectionFieldMaxLength">集合属性字段的最大元素个数</param>
		/// <param name="mapDictionary">key为excel模型字段名，value为标题栏</param>
		/// <param name="hyperlinkFields">超链接字段集合</param>
		/// <param name="optionsAction">选项设置</param>
		/// <returns></returns>
		public static byte[] ToExcelWithCollectionAppoint<T>(this IEnumerable<T> datas, string collectionFieldName,
			int collectionFieldMaxLength, Dictionary<string, string> mapDictionary, List<string> hyperlinkFields = null,
			Action<ToExcelOptions> optionsAction = null)
		{
			var options = new ToExcelOptions();
			optionsAction?.Invoke(options);

			//创建DataTable
			var dTable = CreateDataTable(datas, collectionFieldName, collectionFieldMaxLength, mapDictionary);

			using (var package = new ExcelPackage())
			{
				var ws = package.Workbook.Worksheets.Add("数据");
				ws.Cells[1, 1].LoadFromDataTable(dTable, true, options.TableStyles);

				//设置单元格格式 
				var type = typeof(T);
				var index = 0;
				var rowCount = datas.Count();
				var cycleCount = 0;
				foreach (var item in mapDictionary)
				{
					var prop = type.GetProperty(item.Key,
						BindingFlags.Instance | BindingFlags.Public | BindingFlags.IgnoreCase);
					if (prop == null)
						continue;

					if (prop.Name.ToLower() == collectionFieldName.ToLower())
					{
						index += collectionFieldMaxLength;
						cycleCount = collectionFieldMaxLength - 1;
					}
					else
					{
						index++;
						cycleCount = 0;
					}

					//设置日期显示格式
					if (prop.PropertyType == typeof(DateTime))
						ws.Column(index).Style.Numberformat.Format = options.DatetimeFormate;

					//设置单元格超链接
					if (hyperlinkFields != null && hyperlinkFields.Any(a => a.ToLower() == prop.Name.ToLower()))
					{
						for (var count = 1; count <= rowCount; count++)
						{
							for (var i = cycleCount; i >= 0; i--)
							{
								var value = ws.Cells[count + 1, index - i].Value?.ToString();
								if (!string.IsNullOrWhiteSpace(value) &&
								    Uri.TryCreate(value, UriKind.Absolute, out Uri uri))
								{
									ws.Cells[count + 1, index - i].Hyperlink = uri;
								}
							}
						}
					}
				}

				ws.Cells[ws.Dimension.Address].AutoFitColumns();
				return package.GetAsByteArray();
			}
		}

		/// <summary>
		/// 输出Excel, 根据字典顺序排序, 支持单属性字段为(字符串)集合的对象集合, 超链接设置方式为遍历指定的列
		/// 根据指定唯一列计算合并列, 按照字典顺序, 从第一列开始, 计算合并mergeLength指定长度的单元格
		/// </summary>
		/// <typeparam name="T"></typeparam>
		/// <param name="datas"></param>
		/// <param name="collectionFieldName">集合属性字段名称</param>
		/// <param name="collectionFieldMaxLength">集合属性字段的最大元素个数</param>
		/// <param name="mapDictionary">key为excel模型字段名，value为标题栏</param>
		/// <param name="uniqueColumnName">合并唯一标识列</param>
		/// <param name="mergeLength">要合并列的长度</param>
		/// <param name="hyperlinkFields">超链接字段集合</param>
		/// <param name="optionsAction">选项设置</param>
		/// <returns></returns>
		public static byte[] ToExcelWithCollectionAppointMerge<T>(this IEnumerable<T> datas, string collectionFieldName,
			int collectionFieldMaxLength, Dictionary<string, string> mapDictionary, string uniqueColumnName,
			int mergeLength, List<string> hyperlinkFields = null, Action<ToExcelOptions> optionsAction = null)
		{
			var options = new ToExcelOptions();
			optionsAction?.Invoke(options);

			//创建DataTable
			var dTable = CreateDataTable(datas, collectionFieldName, collectionFieldMaxLength, mapDictionary);

			using (var package = new ExcelPackage())
			{
				var ws = package.Workbook.Worksheets.Add("数据");
				ws.Cells[1, 1].LoadFromDataTable(dTable, true);

				//设置单元格格式 
				var type = typeof(T);
				var index = 0;
				var rowCount = datas.Count();
				var cycleCount = 0;
				foreach (var item in mapDictionary)
				{
					var prop = type.GetProperty(item.Key,
						BindingFlags.Instance | BindingFlags.Public | BindingFlags.IgnoreCase);
					if (prop == null)
						continue;

					if (prop.Name.ToLower() == collectionFieldName.ToLower())
					{
						index += collectionFieldMaxLength;
						cycleCount = collectionFieldMaxLength - 1;
					}
					else
					{
						index++;
						cycleCount = 0;
					}

					//设置日期显示格式
					if (prop.PropertyType == typeof(DateTime))
						ws.Column(index).Style.Numberformat.Format = options.DatetimeFormate;

					//设置单元格超链接
					if (hyperlinkFields != null && hyperlinkFields.Any(a => a.ToLower() == prop.Name.ToLower()))
					{
						for (var count = 1; count <= rowCount; count++)
						{
							for (var i = cycleCount; i >= 0; i--)
							{
								var value = ws.Cells[count + 1, index - i].Value?.ToString();
								if (!string.IsNullOrWhiteSpace(value) &&
								    Uri.TryCreate(value, UriKind.Absolute, out Uri uri))
								{
									ws.Cells[count + 1, index - i].Hyperlink = uri;
								}
							}
						}
					}
				}

				//合并单元格
				MergeSetting(ws, dTable, mapDictionary, uniqueColumnName, mergeLength, 1, collectionFieldMaxLength);

				ws.Cells[ws.Dimension.Address].AutoFitColumns();
				return package.GetAsByteArray();
			}
		}

		/// <summary>
		/// 输出Excel, 根据字典顺序排序, 支持多属性字段为(字符串)集合的对象集合, 超链接设置方式为遍历所有单元格
		/// </summary>
		/// <typeparam name="T"></typeparam>
		/// <param name="datas"></param>
		/// <param name="collectionFields">集合属性字段字典</param>
		/// <param name="mapDictionary">key为excel模型字段名，value为标题栏</param>
		/// <param name="hasHyperlink">是否有超链接</param>
		/// <param name="optionsAction">选项设置</param>
		/// <returns></returns>
		public static byte[] ToExcelWithCollection<T>(this IEnumerable<T> datas,
			Dictionary<string, int> collectionFields,
			Dictionary<string, string> mapDictionary, bool hasHyperlink = false,
			Action<ToExcelOptions> optionsAction = null)
		{
			var options = new ToExcelOptions();
			optionsAction?.Invoke(options);

			//创建DataTable
			var dTable = CreateDataTable(datas, collectionFields, mapDictionary);

			using (var package = new ExcelPackage())
			{
				var ws = package.Workbook.Worksheets.Add("数据");
				ws.Cells[1, 1].LoadFromDataTable(dTable, true, options.TableStyles);

				//设置单元格格式 
				var type = typeof(T);
				var index = 0;
				var rowCount = datas.Count();
				var cycleCount = 0;
				foreach (var item in mapDictionary)
				{
					var prop = type.GetProperty(item.Key,
						BindingFlags.Instance | BindingFlags.Public | BindingFlags.IgnoreCase);
					if (prop == null)
						continue;

					var fieldKey = collectionFields.Where(w => w.Key.ToLower() == prop.Name.ToLower())
						.Select(s => s.Key).FirstOrDefault();
					if (!string.IsNullOrWhiteSpace(fieldKey))
					{
						index += collectionFields[fieldKey];
						cycleCount = collectionFields[fieldKey] - 1;
					}
					else
					{
						index++;
						cycleCount = 0;
					}

					//设置日期显示格式
					if (prop.PropertyType == typeof(DateTime))
						ws.Column(index).Style.Numberformat.Format = options.DatetimeFormate;

					//设置单元格超链接
					if (hasHyperlink)
					{
						for (var count = 1; count <= rowCount; count++)
						{
							for (var i = cycleCount; i >= 0; i--)
							{
								var value = ws.Cells[count + 1, index - i].Value?.ToString();
								if (!string.IsNullOrWhiteSpace(value) &&
								    Uri.TryCreate(value, UriKind.Absolute, out Uri uri))
								{
									ws.Cells[count + 1, index - i].Hyperlink = uri;
								}
							}
						}
					}
				}

				ws.Cells[ws.Dimension.Address].AutoFitColumns();
				return package.GetAsByteArray();
			}
		}

		/// <summary>
		/// 输出Excel, 根据字典顺序排序, 支持多属性字段为(字符串)集合的对象集合, 超链接设置方式为遍历所有单元格
		/// 根据指定唯一列计算合并列, 按照字典顺序, 从第一列开始, 计算合并mergeLength指定长度的单元格
		/// </summary>
		/// <typeparam name="T"></typeparam>
		/// <param name="datas"></param>
		/// <param name="collectionFields">集合属性字段字典</param>
		/// <param name="mapDictionary">key为excel模型字段名，value为标题栏</param>
		/// <param name="uniqueColumnName">合并唯一标识列</param>
		/// <param name="mergeLength">要合并列的长度</param>
		/// <param name="hasHyperlink">是否有超链接</param>
		/// <param name="optionsAction">选项设置</param>
		/// <returns></returns>
		public static byte[] ToExcelWithCollectionMerge<T>(this IEnumerable<T> datas,
			Dictionary<string, int> collectionFields,
			Dictionary<string, string> mapDictionary, string uniqueColumnName, int mergeLength,
			bool hasHyperlink = false, Action<ToExcelOptions> optionsAction = null)
		{
			var options = new ToExcelOptions();
			optionsAction?.Invoke(options);

			//创建DataTable
			var dTable = CreateDataTable(datas, collectionFields, mapDictionary);

			using (var package = new ExcelPackage())
			{
				var ws = package.Workbook.Worksheets.Add("数据");
				ws.Cells[1, 1].LoadFromDataTable(dTable, true);

				//设置单元格格式 
				var type = typeof(T);
				var index = 0;
				var rowCount = datas.Count();
				var cycleCount = 0;
				foreach (var item in mapDictionary)
				{
					var prop = type.GetProperty(item.Key,
						BindingFlags.Instance | BindingFlags.Public | BindingFlags.IgnoreCase);
					if (prop == null)
						continue;

					var fieldKey = collectionFields.Where(w => w.Key.ToLower() == prop.Name.ToLower())
						.Select(s => s.Key).FirstOrDefault();
					if (!string.IsNullOrWhiteSpace(fieldKey))
					{
						index += collectionFields[fieldKey];
						cycleCount = collectionFields[fieldKey] - 1;
					}
					else
					{
						index++;
						cycleCount = 0;
					}

					//设置日期显示格式
					if (prop.PropertyType == typeof(DateTime))
						ws.Column(index).Style.Numberformat.Format = options.DatetimeFormate;

					//设置单元格超链接
					if (hasHyperlink)
					{
						for (var count = 1; count <= rowCount; count++)
						{
							for (var i = cycleCount; i >= 0; i--)
							{
								var value = ws.Cells[count + 1, index - i].Value?.ToString();
								if (!string.IsNullOrWhiteSpace(value) &&
								    Uri.TryCreate(value, UriKind.Absolute, out Uri uri))
								{
									ws.Cells[count + 1, index - i].Hyperlink = uri;
								}
							}
						}
					}
				}

				//合并单元格
				var extendColumnTotalLength = collectionFields.Sum(s => s.Value);
				MergeSetting(ws, dTable, mapDictionary, uniqueColumnName, mergeLength, collectionFields.Count,
					extendColumnTotalLength);

				ws.Cells[ws.Dimension.Address].AutoFitColumns();
				return package.GetAsByteArray();
			}
		}

		/// <summary>
		/// 输出Excel, 根据字典顺序排序, 支持多属性字段为(字符串)集合的对象集合, 超链接设置方式为遍历指定的列
		/// </summary>
		/// <typeparam name="T"></typeparam>
		/// <param name="datas"></param>
		/// <param name="collectionFields">集合属性字段字典</param>
		/// <param name="mapDictionary">key为excel模型字段名，value为标题栏</param>
		/// <param name="hyperlinkFields">超链接字段集合</param>
		/// <param name="optionsAction">选项设置</param>
		/// <returns></returns>
		public static byte[] ToExcelWithCollectionAppoint<T>(this IEnumerable<T> datas,
			Dictionary<string, int> collectionFields,
			Dictionary<string, string> mapDictionary, List<string> hyperlinkFields = null,
			Action<ToExcelOptions> optionsAction = null)
		{
			var options = new ToExcelOptions();
			optionsAction?.Invoke(options);

			//创建DataTable
			var dTable = CreateDataTable(datas, collectionFields, mapDictionary);

			using (var package = new ExcelPackage())
			{
				var ws = package.Workbook.Worksheets.Add("数据");
				ws.Cells[1, 1].LoadFromDataTable(dTable, true, options.TableStyles);

				//设置单元格格式 
				var type = typeof(T);
				var index = 0;
				var rowCount = datas.Count();
				var cycleCount = 0;
				foreach (var item in mapDictionary)
				{
					var prop = type.GetProperty(item.Key,
						BindingFlags.Instance | BindingFlags.Public | BindingFlags.IgnoreCase);
					if (prop == null)
						continue;

					var fieldKey = collectionFields.Where(w => w.Key.ToLower() == prop.Name.ToLower())
						.Select(s => s.Key).FirstOrDefault();
					if (!string.IsNullOrWhiteSpace(fieldKey))
					{
						index += collectionFields[fieldKey];
						cycleCount = collectionFields[fieldKey] - 1;
					}
					else
					{
						index++;
						cycleCount = 0;
					}

					//设置日期显示格式
					if (prop.PropertyType == typeof(DateTime))
						ws.Column(index).Style.Numberformat.Format = options.DatetimeFormate;

					//设置单元格超链接
					if (hyperlinkFields != null && hyperlinkFields.Any(a => a.ToLower() == prop.Name.ToLower()))
					{
						for (var count = 1; count <= rowCount; count++)
						{
							for (var i = cycleCount; i >= 0; i--)
							{
								var value = ws.Cells[count + 1, index - i].Value?.ToString();
								if (!string.IsNullOrWhiteSpace(value) &&
								    Uri.TryCreate(value, UriKind.Absolute, out Uri uri))
								{
									ws.Cells[count + 1, index - i].Hyperlink = uri;
								}
							}
						}
					}
				}

				ws.Cells[ws.Dimension.Address].AutoFitColumns();
				return package.GetAsByteArray();
			}
		}

		/// <summary>
		/// 输出Excel, 根据字典顺序排序, 支持多属性字段为(字符串)集合的对象集合, 超链接设置方式为遍历指定的列
		/// 根据指定唯一列计算合并列, 按照字典顺序, 从第一列开始, 计算合并mergeLength指定长度的单元格
		/// </summary>
		/// <typeparam name="T"></typeparam>
		/// <param name="datas"></param>
		/// <param name="collectionFields">集合属性字段字典</param>
		/// <param name="mapDictionary">key为excel模型字段名，value为标题栏</param>
		/// <param name="uniqueColumnName">合并唯一标识列</param>
		/// <param name="mergeLength">要合并列的长度</param>
		/// <param name="hyperlinkFields">超链接字段集合</param>
		/// <param name="optionsAction">选项设置</param>
		/// <returns></returns>
		public static byte[] ToExcelWithCollectionAppointMerge<T>(this IEnumerable<T> datas,
			Dictionary<string, int> collectionFields,
			Dictionary<string, string> mapDictionary, string uniqueColumnName, int mergeLength,
			List<string> hyperlinkFields = null, Action<ToExcelOptions> optionsAction = null)
		{
			var options = new ToExcelOptions();
			optionsAction?.Invoke(options);

			//创建DataTable
			var dTable = CreateDataTable(datas, collectionFields, mapDictionary);

			using (var package = new ExcelPackage())
			{
				var ws = package.Workbook.Worksheets.Add("数据");
				ws.Cells[1, 1].LoadFromDataTable(dTable, true);

				//设置单元格格式 
				var type = typeof(T);
				var index = 0;
				var rowCount = datas.Count();
				var cycleCount = 0;
				foreach (var item in mapDictionary)
				{
					var prop = type.GetProperty(item.Key,
						BindingFlags.Instance | BindingFlags.Public | BindingFlags.IgnoreCase);
					if (prop == null)
						continue;

					var fieldKey = collectionFields.Where(w => w.Key.ToLower() == prop.Name.ToLower())
						.Select(s => s.Key).FirstOrDefault();
					if (!string.IsNullOrWhiteSpace(fieldKey))
					{
						index += collectionFields[fieldKey];
						cycleCount = collectionFields[fieldKey] - 1;
					}
					else
					{
						index++;
						cycleCount = 0;
					}

					//设置日期显示格式
					if (prop.PropertyType == typeof(DateTime))
						ws.Column(index).Style.Numberformat.Format = options.DatetimeFormate;

					//设置单元格超链接
					if (hyperlinkFields != null && hyperlinkFields.Any(a => a.ToLower() == prop.Name.ToLower()))
					{
						for (var count = 1; count <= rowCount; count++)
						{
							for (var i = cycleCount; i >= 0; i--)
							{
								var value = ws.Cells[count + 1, index - i].Value?.ToString();
								if (!string.IsNullOrWhiteSpace(value) &&
								    Uri.TryCreate(value, UriKind.Absolute, out Uri uri))
								{
									ws.Cells[count + 1, index - i].Hyperlink = uri;
								}
							}
						}
					}
				}

				//合并单元格
				var extendColumnTotalLength = collectionFields.Sum(s => s.Value);
				MergeSetting(ws, dTable, mapDictionary, uniqueColumnName, mergeLength, collectionFields.Count,
					extendColumnTotalLength);

				ws.Cells[ws.Dimension.Address].AutoFitColumns();
				return package.GetAsByteArray();
			}
		}

		/// <summary>
		/// 把excel数据填充到集合
		/// </summary>
		/// <typeparam name="T"></typeparam>
		/// <param name="model"></param>
		/// <param name="dataStream">数据流</param>
		/// <param name="mapDictionary">key为excel标题栏，value为模型字段名</param>
		/// <param name="funcDictionary">委托执行字体，，key为属性字段</param>
		public static void FillFromExcel<T>(this IList<T> model, Stream dataStream,
			Dictionary<string, string> mapDictionary = null,
			Dictionary<string, Func<string, object>> funcDictionary = null) where T : new()
		{
			mapDictionary = mapDictionary ?? new Dictionary<string, string>();
			funcDictionary = funcDictionary ?? new Dictionary<string, Func<string, object>>();
			using (var excel = new ExcelPackage(dataStream))
			{
				var sheet = excel.Workbook.Worksheets.First();
				var rows = sheet.Dimension.End.Row;
				var cols = sheet.Dimension.End.Column;

				var colMapDictionary = new Dictionary<int, string>();
				for (int i = 1; i <= cols; i++)
				{
					var title = sheet.Cells[1, i].GetValue<string>();
					if (string.IsNullOrWhiteSpace(title))
						continue;

					if (mapDictionary.ContainsKey(title))
						colMapDictionary.Add(i, mapDictionary[title]);
				}

				ReadData(model, dataStream, colMapDictionary, sheet, rows, cols, funcDictionary);
			}
		}

		#region 内部方法

		private static void ReadData<T>(IList<T> model, Stream dataStream, Dictionary<int, string> colMapDictionary,
			ExcelWorksheet sheet, int rows, int cols, Dictionary<string, Func<string, object>> funcDictionary = null)
			where T : new()
		{
			var firstValue = string.Empty;
			for (int row = 2; row <= rows; row++)
			{
				//第一列值为空时忽略该行
				firstValue = sheet.Cells[row, 1]?.Value?.ToString();
				if (string.IsNullOrWhiteSpace(firstValue))
					continue;

				var data = new T();

				foreach (var key in colMapDictionary.Keys)
				{
					var excelValue = sheet.Cells[row, key]?.Value?.ToString();
					var propertyName = colMapDictionary[key];

					var property = data.GetType().GetProperty(propertyName);
					object value = null;
					if (funcDictionary.ContainsKey(propertyName))
					{
						value = funcDictionary[propertyName].Invoke(excelValue);
					}
					else
					{
						if (typeof(Enum).IsAssignableFrom(property.PropertyType))
							value = Enum.Parse(property.PropertyType, excelValue);
						else
							value = Convert.ChangeType(excelValue, property.PropertyType);
					}

					property.SetValue(data, value);
				}

				model.Add(data);
			}
		}

		/// <summary>
		/// 创建DataTable
		/// </summary>
		/// <typeparam name="T"></typeparam>
		/// <param name="datas"></param>
		/// <param name="mapDictionary"></param>
		/// <returns></returns>
		private static DataTable CreateDataTable<T>(this IEnumerable<T> datas, Dictionary<string, string> mapDictionary)
		{
			//创建表结构
			var type = typeof(T);
			var dTable = new DataTable();
			foreach (var key in mapDictionary.Keys)
			{
				var columnName = mapDictionary[key];

				var prop = type.GetProperty(key, BindingFlags.Instance | BindingFlags.Public | BindingFlags.IgnoreCase);
				if (prop == null)
				{
					dTable.Columns.Add(new DataColumn(columnName));
				}
				else
				{
					if (prop.PropertyType == typeof(DateTime))
						dTable.Columns.Add(new DataColumn(columnName, typeof(DateTime)));
					else
						dTable.Columns.Add(new DataColumn(columnName));
				}
			}

			//datatable赋值
			foreach (var data in datas)
			{
				var row = dTable.NewRow();
				foreach (var key in mapDictionary.Keys)
				{
					var columnName = mapDictionary[key];
					row[columnName] = type.GetProperty(key)?.GetValue(data);
				}

				dTable.Rows.Add(row);
			}

			return dTable;
		}

		/// <summary>
		/// 创建DataTable
		/// </summary>
		/// <typeparam name="T"></typeparam>
		/// <param name="datas"></param>
		/// <param name="collectionFieldName"></param>
		/// <param name="collectionFieldMaxLength"></param>
		/// <param name="mapDictionary"></param>
		/// <returns></returns>
		private static DataTable CreateDataTable<T>(this IEnumerable<T> datas, string collectionFieldName,
			int collectionFieldMaxLength, Dictionary<string, string> mapDictionary)
		{
			//创建表结构
			var type = typeof(T);
			var dTable = new DataTable();
			foreach (var key in mapDictionary.Keys)
			{
				var columnName = mapDictionary[key];

				var prop = type.GetProperty(key, BindingFlags.Instance | BindingFlags.Public | BindingFlags.IgnoreCase);
				if (prop == null)
				{
					if (key == collectionFieldName && collectionFieldMaxLength > 0)
					{
						for (int i = 1; i <= collectionFieldMaxLength; i++)
						{
							dTable.Columns.Add(new DataColumn(columnName + i));
						}
					}
					else
					{
						dTable.Columns.Add(new DataColumn(columnName));
					}
				}
				else
				{
					if (key == collectionFieldName && collectionFieldMaxLength > 0)
					{
						for (int i = 1; i <= collectionFieldMaxLength; i++)
						{
							if (prop.PropertyType == typeof(DateTime))
								dTable.Columns.Add(new DataColumn(columnName + i, typeof(DateTime)));
							else
								dTable.Columns.Add(new DataColumn(columnName + i));
						}
					}
					else
					{
						if (prop.PropertyType == typeof(DateTime))
							dTable.Columns.Add(new DataColumn(columnName, typeof(DateTime)));
						else
							dTable.Columns.Add(new DataColumn(columnName));
					}
				}
			}

			//datatable赋值
			foreach (var data in datas)
			{
				var row = dTable.NewRow();
				foreach (var key in mapDictionary.Keys)
				{
					var columnName = mapDictionary[key];
					if (key == collectionFieldName)
					{
						if (collectionFieldMaxLength > 0)
						{
							var items = (IEnumerable<string>) type.GetProperty(key).GetValue(data);
							if (items != null)
							{
								var index = 0;
								foreach (var item in items)
								{
									index++;
									row[columnName + index] = item;
								}
							}
						}
						else
						{
							row[columnName] = "";
						}
					}
					else
					{
						row[columnName] = type.GetProperty(key)?.GetValue(data);
					}
				}

				dTable.Rows.Add(row);
			}

			return dTable;
		}

		/// <summary>
		/// 创建DataTable
		/// </summary>
		/// <typeparam name="T"></typeparam>
		/// <param name="datas"></param>
		/// <param name="collectionFields"></param>
		/// <param name="mapDictionary"></param>
		/// <returns></returns>
		private static DataTable CreateDataTable<T>(this IEnumerable<T> datas, Dictionary<string, int> collectionFields,
			Dictionary<string, string> mapDictionary)
		{
			//创建表结构
			var type = typeof(T);
			var dTable = new DataTable();
			foreach (var key in mapDictionary.Keys)
			{
				var columnName = mapDictionary[key];

				var prop = type.GetProperty(key, BindingFlags.Instance | BindingFlags.Public | BindingFlags.IgnoreCase);
				if (prop == null)
				{
					if (collectionFields.ContainsKey(key) && collectionFields[key] > 0)
					{
						for (int i = 1; i <= collectionFields[key]; i++)
						{
							dTable.Columns.Add(new DataColumn(columnName + i));
						}
					}
					else
					{
						dTable.Columns.Add(new DataColumn(columnName));
					}
				}
				else
				{
					if (collectionFields.ContainsKey(key) && collectionFields[key] > 0)
					{
						for (int i = 1; i <= collectionFields[key]; i++)
						{
							if (prop.PropertyType == typeof(DateTime))
								dTable.Columns.Add(new DataColumn(columnName + i, typeof(DateTime)));
							else
								dTable.Columns.Add(new DataColumn(columnName + i));
						}
					}
					else
					{
						if (prop.PropertyType == typeof(DateTime))
							dTable.Columns.Add(new DataColumn(columnName, typeof(DateTime)));
						else
							dTable.Columns.Add(new DataColumn(columnName));
					}
				}
			}

			//datatable赋值
			foreach (var data in datas)
			{
				var row = dTable.NewRow();
				foreach (var key in mapDictionary.Keys)
				{
					var columnName = mapDictionary[key];
					if (collectionFields.ContainsKey(key))
					{
						if (collectionFields[key] > 0)
						{
							var items = (IEnumerable<string>) type.GetProperty(key).GetValue(data);
							if (items != null)
							{
								var index = 0;
								foreach (var item in items)
								{
									index++;
									row[columnName + index] = item;
								}
							}
						}
						else
						{
							row[columnName] = "";
						}
					}
					else
					{
						row[columnName] = type.GetProperty(key)?.GetValue(data);
					}
				}

				dTable.Rows.Add(row);
			}

			return dTable;
		}

		/// <summary>
		/// 合并设置
		/// </summary>
		/// <param name="ws">sheet</param>
		/// <param name="dTable">数据table</param>
		/// <param name="mapDictionary">导出列字典</param>
		/// <param name="uniqueColumnName">合并唯一列key</param>
		/// <param name="mergeLength">合并列长度</param>
		/// <param name="extendColumnCount">扩展列数量</param>
		/// <param name="extendColumnTotalLength">扩展列总长度</param>
		private static void MergeSetting(ExcelWorksheet ws, DataTable dTable, Dictionary<string, string> mapDictionary,
			string uniqueColumnName, int mergeLength, int extendColumnCount = 0, int extendColumnTotalLength = 0)
		{
			var rowCount = dTable.Rows.Count;
			var mergeColumnCount = mapDictionary.Count;
			if (extendColumnCount > 0)
				mergeColumnCount = mergeColumnCount - extendColumnCount + extendColumnTotalLength;

			//合并单元格
			if (mapDictionary.ContainsKey(uniqueColumnName))
			{
				var columnName = mapDictionary[uniqueColumnName];
				var columnIndex = dTable.Columns[columnName].Ordinal;

				var mergeStartIndex = 0;
				var mergeRowCount = 0;
				var mergeRowTotal = 0;
				for (var i = 0; i < rowCount; i++)
				{
					if (i + 1 < rowCount && dTable.Rows[i][columnIndex].ToString() ==
					    dTable.Rows[i + 1][columnIndex].ToString())
					{
						//记录合并起始位置
						if (mergeStartIndex == 0)
							mergeStartIndex = i;

						mergeRowCount++;
						continue;
					}
					else
					{
						mergeRowTotal++;
						if (mergeRowCount > 0)
						{
							for (var n = 1; n <= mergeLength; n++)
							{
								ws.Cells[mergeStartIndex + 2, n, i + 2, n].Merge = true;

								if (mergeRowTotal % 2 == 0)
								{
									ws.Cells[mergeStartIndex + 2, n, i + 2, n].Style.Fill.PatternType =
										ExcelFillStyle.Solid;
									ws.Cells[mergeStartIndex + 2, n, i + 2, n].Style.Fill.BackgroundColor
										.SetColor(Color.FromArgb(189, 215, 238));
									ws.Cells[mergeStartIndex + 2, n, i + 2, n].Style.Border
										.BorderAround(ExcelBorderStyle.Thin, Color.FromArgb(156, 156, 156));
								}
								else
								{
									ws.Cells[mergeStartIndex + 2, n, i + 2, n].Style.Fill.PatternType =
										ExcelFillStyle.Solid;
									ws.Cells[mergeStartIndex + 2, n, i + 2, n].Style.Fill.BackgroundColor
										.SetColor(Color.FromArgb(221, 235, 247));
									ws.Cells[mergeStartIndex + 2, n, i + 2, n].Style.Border
										.BorderAround(ExcelBorderStyle.Thin, Color.FromArgb(156, 156, 156));
								}
							}

							for (var m = mergeStartIndex; m <= i; m++)
							{
								for (var n = mergeLength + 1; n <= mergeColumnCount; n++)
								{
									if (mergeRowTotal % 2 == 0)
									{
										ws.Cells[m + 2, n].Style.Fill.PatternType = ExcelFillStyle.Solid;
										ws.Cells[m + 2, n].Style.Fill.BackgroundColor
											.SetColor(Color.FromArgb(189, 215, 238));
										ws.Cells[m + 2, n].Style.Border.BorderAround(ExcelBorderStyle.Thin,
											Color.FromArgb(156, 156, 156));
									}
									else
									{
										ws.Cells[m + 2, n].Style.Fill.PatternType = ExcelFillStyle.Solid;
										ws.Cells[m + 2, n].Style.Fill.BackgroundColor
											.SetColor(Color.FromArgb(221, 235, 247));
										ws.Cells[m + 2, n].Style.Border.BorderAround(ExcelBorderStyle.Thin,
											Color.FromArgb(156, 156, 156));
									}
								}
							}
						}
						else
						{
							for (var n = 1; n <= mergeColumnCount; n++)
							{
								if (mergeRowTotal % 2 == 0)
								{
									ws.Cells[i + 2, n].Style.Fill.PatternType = ExcelFillStyle.Solid;
									ws.Cells[i + 2, n].Style.Fill.BackgroundColor
										.SetColor(Color.FromArgb(189, 215, 238));
									ws.Cells[i + 2, n].Style.Border.BorderAround(ExcelBorderStyle.Thin,
										Color.FromArgb(156, 156, 156));
								}
								else
								{
									ws.Cells[i + 2, n].Style.Fill.PatternType = ExcelFillStyle.Solid;
									ws.Cells[i + 2, n].Style.Fill.BackgroundColor
										.SetColor(Color.FromArgb(221, 235, 247));
									ws.Cells[i + 2, n].Style.Border.BorderAround(ExcelBorderStyle.Thin,
										Color.FromArgb(156, 156, 156));
								}
							}
						}

						mergeStartIndex = 0;
						mergeRowCount = 0;
					}
				}
			}

			//设置行列样式
			for (var m = 0; m <= rowCount; m++)
			{
				ws.Row(m + 1).Height = 28; //行高

				for (var n = 1; n <= mergeColumnCount; n++)
				{
					ws.Cells[m + 1, n].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left; //水平居左
					ws.Cells[m + 1, n].Style.VerticalAlignment = ExcelVerticalAlignment.Center; //水平居中
				}
			}

			//设置标题行样式
			for (var n = 1; n <= mergeColumnCount; n++)
			{
				ws.Cells[1, n].Style.Font.Bold = true;
				ws.Cells[1, n].Style.Font.Color.SetColor(Color.FromArgb(255, 255, 255));
				ws.Cells[1, n].Style.Font.Name = "微软雅黑";
				ws.Cells[1, n].Style.Font.Size = 12;
				ws.Cells[1, n].Style.Fill.PatternType = ExcelFillStyle.Solid;
				ws.Cells[1, n].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(91, 155, 213));
				ws.Cells[1, n].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.FromArgb(156, 156, 156));
			}
		}

		#endregion
	}

	public class ToExcelOptions
	{
		/// <summary>
		/// 导出表格样式
		/// </summary>
		public TableStyles TableStyles { get; set; } = TableStyles.Medium9;

		/// <summary>
		/// 时间格式
		/// </summary>
		public string DatetimeFormate { get; set; } = "yyyy/dd/MM HH:mm:ss";
	}
}