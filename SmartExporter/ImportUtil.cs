using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Globalization;
using System.IO;
using System.Text.RegularExpressions;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;

namespace SmartExporter
{
    /// <summary>
    /// 导入工具类，提供同步异步导入功能
    /// 同步导入功能是静态的，异步需要实例化本类
    /// 在异步过程中需要进度信息的，需要实现BackgroundWorker的相应方法
    /// 异步的启动方法一定要使用ImportFromExcelAsync方法
    /// </summary>
    public class ImportUtil
    {
        public BackgroundWorker bgBackgroundWorker { get; }
        public delegate void ProgressChangedDelegate(int i, string msg);
        /// <summary>
        /// 导出进度变化的通知
        /// </summary>
        public ProgressChangedDelegate ProgressChanged;

        public delegate void ProgressCompletedDelegate(bool compliete, Exception msg);
        /// <summary>
        /// 导出完成后的事件通知
        /// </summary>
        public ProgressCompletedDelegate ProgressCompleted;
        public ImportUtil()
        {
            bgBackgroundWorker = new BackgroundWorker();
            bgBackgroundWorker.DoWork += BgBackgroundWorkerOnDoWork;

            bgBackgroundWorker.ProgressChanged += (sender, args) =>
            {
                ProgressChanged?.Invoke(args.ProgressPercentage, args.UserState.ToString());
            };
            bgBackgroundWorker.WorkerReportsProgress = true;
            bgBackgroundWorker.RunWorkerCompleted += (sender, args) =>
            {
                ProgressCompleted?.Invoke((bool)args.Result, args.Error);
            };
        }

        /// <summary>
        /// 启动异步导入方法
        /// </summary>
        /// <param name="filePath">文件路径</param>
        /// <param name="T">需要转换成的Model的类型</param>
        public void ImportFromExcelAsync(string filePath, Type T)
        {
            bgBackgroundWorker.RunWorkerAsync(new object[] { filePath, T });
        }

        private void BgBackgroundWorkerOnDoWork(object sender, DoWorkEventArgs e)
        {
            //判断参数是否正确
            if (!(e.Argument is object[] args) || args.Length != 2)
            {
                throw new Exception("传入参数出错。");
            }
            var filePath = args[0] as string;
            var T = args[1] as Type;
            if (T == null || string.IsNullOrEmpty(filePath))
            {
                throw new Exception("传入参数存在空值。");
            }
            //保存当前类的属性与特性对应，<属性名，特性名>
            const string match = "^(\\d{2}|\\d{4})[/|-]([0][1-9]|(1[0-2]))[/|-]([1-9]|([012]\\d)|(3[01]))([ \\t\\n\\x0B\\f\\r])*(([0-1]{1}[0-9]{1})|([2]{1}[0-4]{1}))([:])(([0-5]{1}[0-9]{1}|[6]{1}[0]{1}))((([:])((([0-5]{1}[0-9]{1}|[6]{1}[0]{1}))))|())$";
            var propDic = new Dictionary<string, string>();
            foreach (var m in T.GetProperties())
            {
                foreach (Attribute a in m.GetCustomAttributes(true))
                {
                    if (a is ExcelDisplayInfo dbi)
                    {
                        if (null != dbi)
                        {
                            propDic.Add(dbi.Name, m.Name);
                        }
                    }

                }
            }
            //创建对应类型的List
            var result = Activator.CreateInstance(typeof(List<>).MakeGenericType(T)) as IList;
            try
            {
                var Workbook = new HSSFWorkbook(File.OpenRead(filePath));
                var sheetAt = Workbook.GetSheetAt(0);
                var columns = new List<string>();
                var row0 = sheetAt.GetRow(0);
                bgBackgroundWorker.ReportProgress(1, "正在获取文件");
                for (var i = 0; i < sheetAt.GetRow(0).LastCellNum; i++)
                {
                    if (propDic.TryGetValue(row0.GetCell(i).StringCellValue, out var propName))
                    {
                        columns.Add(propName);
                    }
                }
                for (var i = 1; i <= sheetAt.LastRowNum; i++)
                {
                    var t = Activator.CreateInstance(T);
                    var row = sheetAt.GetRow(i);
                    for (var j = 0; j < row.LastCellNum; j++)
                    {
                        var cell = row.GetCell(j);
                        var propName = columns[j];
                        var prop = t.GetType().GetProperty(propName);
                        if (prop == null) return;
                        if (cell == null)
                        {
                            prop.SetValue(t, null, null);
                        }
                        else
                        {
                            var cellValue = string.Empty;
                            switch (cell.CellType)
                            {
                                case CellType.Unknown:
                                    cellValue = cell.StringCellValue;
                                    break;
                                case CellType.Numeric:
                                    //var s = cell.ToString();
                                    //var m = Regex.IsMatch(cell.ToString(), match);
                                    cellValue = Regex.IsMatch(cell.ToString(), match) ? cell.DateCellValue.ToString("yyyy-MM-dd HH:mm:ss") : cell.NumericCellValue.ToString();
                                    break;
                                case CellType.String:
                                    cellValue = cell.StringCellValue;
                                    break;
                                case CellType.Formula:
                                    break;
                                case CellType.Blank:
                                    cellValue = "";
                                    break;
                                case CellType.Boolean:
                                    cellValue = cell.BooleanCellValue.ToString();
                                    break;
                                case CellType.Error:
                                    break;
                            }
                            if (prop.PropertyType == typeof(DateTime?))
                            {
                                if (string.IsNullOrEmpty(cellValue))
                                {
                                    prop.SetValue(t, null, null);
                                }
                                else
                                {
                                    var tryParse = DateTime.TryParse(cellValue, out var time);
                                    if (!tryParse) throw new TypeCastErrorException($"{i}行{j}列数据格式不正确，应当为日期类型。");
                                    prop.SetValue(t, time, null);
                                }

                            }
                            else if (prop.PropertyType == typeof(int?))
                            {
                                var tryParse = int.TryParse(cellValue, out var value);
                                if (!tryParse && !string.IsNullOrEmpty(cellValue)) throw new TypeCastErrorException($"{i}行{j}列数据格式不正确，应当为整数类型。");
                                prop.SetValue(t, value, null);
                            }
                            else if (prop.PropertyType == typeof(float?))
                            {
                                var tryParse = float.TryParse(cellValue, out var value);
                                if (!tryParse && !string.IsNullOrEmpty(cellValue)) throw new TypeCastErrorException($"{i}行{j}列数据格式不正确，应当为数字类型。");
                                prop.SetValue(t, value, null);
                            }
                            else if (prop.PropertyType == typeof(double?))
                            {
                                var tryParse = double.TryParse(cellValue, out var value);
                                if (!tryParse && !string.IsNullOrEmpty(cellValue)) throw new TypeCastErrorException($"{i}行{j}列数据格式不正确，应当为数字类型。");

                                prop.SetValue(t, value, null);
                            }
                            else
                            {
                                prop.SetValue(t, cell.StringCellValue, null);
                            }
                        }
                    }

                    result?.Add(t);
                    var progress = (double)i / sheetAt.LastRowNum * 100;
                    bgBackgroundWorker.ReportProgress(progress > 100 ? 100 : (int)progress, $"获取文件第{i}行数据，共{sheetAt.LastRowNum}行。");

                }
                bgBackgroundWorker.ReportProgress(99, "正在进行合成。");
                e.Result = result;
                bgBackgroundWorker.ReportProgress(100, "数据读取完成。");
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
                throw ex;
                //MessageBox.Show(ex.ToString());
            }
        }

        public static List<T> ImportFromExcel<T>(string filePath)
        {
            var result = new List<T>();
            try
            {
                var hssfWorkbook = new HSSFWorkbook(File.OpenRead(filePath));
                var sheetAt = hssfWorkbook.GetSheetAt(0);
                var columns = new List<string>();
                var row0 = sheetAt.GetRow(0);
                for (var i = 0; i < row0.LastCellNum; i++)
                {
                    columns.Add(row0.GetCell(i).StringCellValue);
                }
                for (var i = 1; i <= sheetAt.LastRowNum; i++)
                {
                    var t = Activator.CreateInstance<T>();
                    var row = sheetAt.GetRow(i);
                    for (var j = 0; j < row.LastCellNum; j++)
                    {
                        var cell = row.GetCell(j);
                        var propName = columns[j];
                        var prop = t.GetType().GetProperty(propName);
                        if (prop == null) return null;
                        if (cell == null)
                        {
                            prop.SetValue(t, "", null);
                        }
                        else
                        {
                            if (prop.PropertyType == typeof(DateTime))
                            {
                                prop.SetValue(t, DateTime.Parse(cell.StringCellValue), null);
                            }
                            else if (prop.PropertyType == typeof(int))
                            {
                                int.TryParse(cell.StringCellValue, out var value);
                                prop.SetValue(t, value, null);
                            }
                            else if (prop.PropertyType == typeof(float))
                            {
                                float.TryParse(cell.StringCellValue, out var value);
                                prop.SetValue(t, value, null);
                            }
                            else if (prop.PropertyType == typeof(double))
                            {
                                double.TryParse(cell.StringCellValue, out var value);
                                prop.SetValue(t, value, null);
                            }
                            else
                            {
                                prop.SetValue(t, cell.StringCellValue, null);
                            }
                        }
                    }
                    result.Add(t);
                }

                return result;
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
                throw;
            }
        }
    }
}
