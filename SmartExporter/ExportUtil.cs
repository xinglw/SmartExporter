using System;
using System.Collections.Generic;
using System.ComponentModel;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using System.Linq;
using System.IO;
using NPOI.SS.Util;

namespace SmartExporter
{
    /// <summary>
    /// 导出对应的Excel
    /// </summary>
    public class ExportUtil
    {
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
        public BackgroundWorker backgroundWoker { get; private set; }

        /// <summary>
        /// 普通单元格字体名称
        /// </summary>
        public string NormalFontName { get; set; } = "微软雅黑";

        /// <summary>
        /// 普通单元格字体大小
        /// </summary>
        public double NormalFontSize { get; set; } = 100;
        /// <summary>
        /// 普通单元格内容对齐样式，横向
        /// </summary>
        public HorizontalAlignment NormalHorizontalAlignment { get; set; } = HorizontalAlignment.Center;
        /// <summary>
        /// 普通单元格内容纵向对齐样式
        /// </summary>
        public VerticalAlignment NormalVerticalAlignment { get; set; } = VerticalAlignment.Center;

        /// <summary>
        /// 标题单元格字体名称
        /// </summary>
        public string TitleFontName { get; set; } = "微软雅黑";

        /// <summary>
        /// 标题单元格字体大小
        /// </summary>
        public double TitleFontSize { get; set; } = 450;
        /// <summary>
        /// 标题单元格内容对齐样式，横向
        /// </summary>
        public HorizontalAlignment TitleHorizontalAlignment { get; set; } = HorizontalAlignment.Center;
        /// <summary>
        /// 标题单元格内容纵向对齐样式
        /// </summary>
        public VerticalAlignment TitleVerticalAlignment { get; set; } = VerticalAlignment.Center;

        /// <summary>
        /// 统计单元格字体名称
        /// </summary>
        public string StatisticsFontName { get; set; } = "微软雅黑";

        /// <summary>
        /// 统计单元格字体大小
        /// </summary>
        public double StatisticsFontSize { get; set; } = 100;
        /// <summary>
        /// 统计单元格内容对齐样式，横向
        /// </summary>
        public HorizontalAlignment StatisticsHorizontalAlignment { get; set; } = HorizontalAlignment.Center;
        /// <summary>
        /// 统计单元格内容纵向对齐样式
        /// </summary>
        public VerticalAlignment StatisticsVerticalAlignment { get; set; } = VerticalAlignment.Center;
        public ExportUtil()
        {
            backgroundWoker = new BackgroundWorker();
            backgroundWoker.ProgressChanged += (sender, args) =>
            {
                ProgressChanged?.Invoke(args.ProgressPercentage, args.UserState.ToString());
            };
            backgroundWoker.WorkerReportsProgress = true;
            backgroundWoker.RunWorkerCompleted += (sender, args) =>
            {
                ProgressCompleted?.Invoke((bool)args.Result, args.Error);
            };
        }
        ///// <summary>
        ///// 
        ///// </summary>
        ///// <typeparam name="T"></typeparam>
        ///// <param name="filePath">包含文件名的文件路径，绝对路径，相对路径暂不支持</param>
        ///// <param name="list"></param>
        //public void ExportAsync<T>(string filePath, List<T> list)
        //{
        //    if (list == null || list.Count == 0) return;
        //    if (string.IsNullOrEmpty(filePath)) return;
        //    var type = typeof(T);
        //    backgroundWoker.DoWork += (a, b) =>
        //    {
        //        try
        //        {
        //            BackgroundWoker_DoWork<T>(filePath, list);
        //            b.Result = true;
        //        }
        //        catch (Exception e)
        //        {
        //            b.Result = false;
        //            throw;
        //        }

        //    };
        //    backgroundWoker.RunWorkerAsync();
        //}



        //private void BackgroundWoker_DoWork<T>(string filePath, List<T> list)
        //{
        //    var propDic = new Dictionary<string, string>();
        //    var propList = new List<ExcelDisplayInfo>();
        //    var type = typeof(T);
        //    foreach (var m in type.GetProperties())
        //    {
        //        foreach (Attribute a in m.GetCustomAttributes(true))
        //        {
        //            if (a is ExcelDisplayInfo dbi)
        //            {
        //                if (null != dbi)
        //                {
        //                    propDic.Add(dbi.Name, m.Name);
        //                    propList.Add(dbi);
        //                }
        //            }
        //        }
        //    }

        //    var array = (from items in propList orderby items.ExportIndex select items.Name).ToArray();

        //    var hssfWorkbook = new HSSFWorkbook();
        //    var sheet = (HSSFSheet)hssfWorkbook.CreateSheet("sheet1");
        //    var cellStyle = hssfWorkbook.CreateCellStyle();
        //    cellStyle.Alignment = HorizontalAlignment.Center;
        //    cellStyle.VerticalAlignment = VerticalAlignment.Center;
        //    var header = (HSSFRow)sheet.CreateRow(0);
        //    for (int i = 0; i < array.Length; i++)
        //    {
        //        var cell = header.CreateCell(i);
        //        cell.SetCellValue(array[i]);
        //        cell.CellStyle = cellStyle;
        //    }
        //    backgroundWoker?.ReportProgress(1, "初始化列名");
        //    FileHelper.CheckPath(filePath);
        //    for (int i = 0; i < list.Count; i++)
        //    {
        //        var rowsCount = (double)i / list.Count;
        //        backgroundWoker?.ReportProgress(rowsCount > 100 ? 100 : (int)(rowsCount * 100), $"正在填充第{i}行数据，共{list.Count}行数据。");
        //        var item = list[i];
        //        var row = (HSSFRow)sheet.CreateRow(i + 1);
        //        for (int j = 0; j < array.Length; j++)
        //        {
        //            var cell = row.CreateCell(j);
        //            var prop = item.GetType().GetProperty(propDic[array[j]]);
        //            if (prop == null) continue;
        //            cell.SetCellValue(prop.GetValue(item).ToString());
        //            cell.CellStyle = cellStyle;
        //        }
        //    }
        //    backgroundWoker?.ReportProgress(100, "正在合并文件，请稍后。。。");
        //    var file = new FileStream(filePath, FileMode.Create);
        //    hssfWorkbook.Write(file);
        //    file.Close();
        //}
        #region 临时测试用方法

        /// <summary>
        /// 导出数据，可以导出带标题，带最后统计的Excel文件
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="filePath">包含文件名的文件路径，绝对路径，相对路径暂不支持</param>
        /// <param name="list">需要导出的实体list，需要导出的属性需要有ExcelDisplayInfo特性修饰，无ExcelDisplayInfo修饰的属性不会导出</param>
        /// <param name="title">导出文件的标题，如不需要可以不传</param>
        /// <param name="Statistics">导出文件最后一行的统计信息，为了保持更多的适配，由调用者传入，如不需要可以不传</param>
        public void ExportAsync<T>(string filePath, List<T> list, string title, string Statistics)
        {
            if (list == null || list.Count == 0) return;
            if (string.IsNullOrEmpty(filePath)) return;
            var type = typeof(T);
            backgroundWoker.DoWork += (a, b) =>
            {
                try
                {
                    ExportWithTitleAndStatistics_DoWork<T>(filePath, list, title,Statistics);
                    b.Result = true;
                }
                catch (Exception e)
                {
                    b.Result = false;
                }

            };
            backgroundWoker.RunWorkerAsync();
        }



        private void ExportWithTitleAndStatistics_DoWork<T>(string filePath, List<T> list,string titlestring,string StatisticsString)
        {
            var propDic = new Dictionary<string, string>();
            var propList = new List<ExcelDisplayInfo>();
            var type = typeof(T);
            foreach (var m in type.GetProperties())
            {
                foreach (Attribute a in m.GetCustomAttributes(true))
                {
                    if (!(a is ExcelDisplayInfo dbi)) continue;
                    propDic.Add(dbi.Name, m.Name);
                    propList.Add(dbi);
                }
            }

            var array = (from items in propList orderby items.ExportIndex select items.Name).ToArray();

            var hssfWorkbook = new HSSFWorkbook();
            var sheet = (HSSFSheet)hssfWorkbook.CreateSheet("sheet1");
            var cellStyle = hssfWorkbook.CreateCellStyle();
            cellStyle.Alignment = NormalHorizontalAlignment;
            cellStyle.VerticalAlignment = NormalVerticalAlignment;
            var normalFont = hssfWorkbook.CreateFont();
            normalFont.FontHeight = NormalFontSize;
            normalFont.FontName = NormalFontName;
            cellStyle.SetFont(normalFont);
            var RowIndexOffset = 0;
            if (!string.IsNullOrEmpty(titlestring))
            {
                RowIndexOffset = 2;
                var title = (HSSFRow)sheet.CreateRow(0);
                var cellc = title.CreateCell(0);
                cellc.SetCellValue(list.Count);
                CellRangeAddress region = new CellRangeAddress(0, 1, 0, array.Length - 1);
                sheet.AddMergedRegion(region);
                var titleStyle = hssfWorkbook.CreateCellStyle();
                titleStyle.Alignment = TitleHorizontalAlignment;
                titleStyle.VerticalAlignment = TitleVerticalAlignment;
                IFont font = hssfWorkbook.CreateFont();
                font.FontName = TitleFontName;
                font.IsBold = true;
                font.FontHeight = TitleFontSize;
                titleStyle.SetFont(font);
                cellc.CellStyle = titleStyle;
                cellc.SetCellValue(titlestring);
            }
            
            var header = (HSSFRow)sheet.CreateRow(RowIndexOffset+ 0);
            //var array = propDic.Keys.ToArray();
            for (int i = 0; i < array.Length; i++)
            {
                var cell = header.CreateCell(i);
                cell.SetCellValue(array[i]);
                cell.CellStyle = cellStyle;
            }
            backgroundWoker?.ReportProgress(1, "初始化列名");
            FileHelper.CheckPath(filePath);
            for (int i = 0; i < list.Count; i++)
            {
                var rowsCount = (double)i / list.Count;
                backgroundWoker?.ReportProgress(rowsCount > 100 ? 100 : (int)(rowsCount * 100), $"正在填充第{i}行数据，共{list.Count}行数据。");
                var item = list[i];
                var row = (HSSFRow)sheet.CreateRow(i + 1 + RowIndexOffset);
                for (int j = 0; j < array.Length; j++)
                {
                    var cell = row.CreateCell(j);
                    var prop = item.GetType().GetProperty(propDic[array[j]]);
                    if (prop == null) continue;
                    cell.SetCellValue(prop.GetValue(item)?.ToString());
                    cell.CellStyle = cellStyle;
                }
            }
            if (!string.IsNullOrEmpty(StatisticsString))
            {
                var lastRow1 = (HSSFRow)sheet.CreateRow(list.Count + 1);
                var cell22 = lastRow1.CreateCell(0);
                cell22.SetCellValue(list.Count);
                CellRangeAddress region1 = new CellRangeAddress(list.Count + 1, list.Count + 1, 0, array.Length - 1);
                sheet.AddMergedRegion(region1);
                var StatisticsStyle = hssfWorkbook.CreateCellStyle();
                cellStyle.Alignment = StatisticsHorizontalAlignment;
                cellStyle.VerticalAlignment = StatisticsVerticalAlignment;
                var StatisticsFont = hssfWorkbook.CreateFont();
                normalFont.FontHeight = StatisticsFontSize;
                normalFont.FontName = StatisticsFontName;
                cellStyle.SetFont(StatisticsFont);
                cell22.CellStyle = StatisticsStyle;
                cell22.SetCellValue(StatisticsString);
            }
            
            backgroundWoker?.ReportProgress(100, "正在合并文件，请稍后。。。");
            var file = new FileStream(filePath, FileMode.Create);
            hssfWorkbook.Write(file);
            file.Close();
        }
        #endregion


    }
}
