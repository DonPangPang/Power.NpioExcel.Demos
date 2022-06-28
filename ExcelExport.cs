using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Threading.Tasks;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.SS.Util;

#nullable disable

namespace TestApp1
{
    public class ExcelExport<T> where T : IExcelStruct
    {
        private HSSFWorkbook _workbook;
        private ISheet _sheet;

        private IRow _headerRow { get; set; }

        private ICellStyle _headerStyle { get; set; }
        private ICellStyle _contentStyle { get; set; }

        private IEnumerable<T> _list { get; set; }

        public ExcelExport(ICellStyle headerStyle = null, ICellStyle contentStyle = null)
        {
            _workbook = new HSSFWorkbook();
            _sheet = _workbook.CreateSheet("Sheet1");

            _headerStyle = headerStyle is null ? ExcelHelper.GetHeaderStyle(_workbook) : headerStyle;
            _contentStyle = contentStyle is null ? ExcelHelper.GetCommonStyle(_workbook) : contentStyle;
        }

        public ExcelExport<T> Export(IEnumerable<T> data)
        {
            _headerRow = _sheet.CreateRow(0);

            _list = data;

            return this;
        }

        private void SetHeader()
        {
            var type = typeof(T);
            var properties = typeof(T).GetProperties().Where(x => Attribute.IsDefined(x, typeof(ExcelColumnAttribute))).OrderBy(x => x.GetCustomAttribute<ExcelColumnAttribute>().Index);

            foreach (var property in properties)
            {
                var attr = property.GetCustomAttribute<ExcelColumnAttribute>();
                var cell = _headerRow.CreateCell(attr.Index);
                cell.SetCellValue(attr.Name);
                cell.CellStyle = _headerStyle;
                _sheet.SetColumnWidth(attr.Index, attr.Width * 256);
            }

            _headerRow.Height = 30 * 20;
        }

        private void SetBody()
        {
            foreach (var item in _list)
            {
                var row = _sheet.CreateRow(1 + _sheet.LastRowNum);
                var properties = typeof(T).GetProperties().Where(x => Attribute.IsDefined(x, typeof(ExcelColumnAttribute))).OrderBy(x => x.GetCustomAttribute<ExcelColumnAttribute>().Index);
                foreach (var property in properties)
                {
                    var attr = property.GetCustomAttribute<ExcelColumnAttribute>();
                    var cell = row.CreateCell(attr.Index);
                    var value = property.GetValue(item);
                    if (value != null)
                    {
                        cell.SetCellValue(value.ToString());
                        cell.CellStyle = _contentStyle;
                    }
                }
            }
        }

        private void SetTemp()
        {
            var properties = typeof(T).GetProperties().Where(x => Attribute.IsDefined(x, typeof(ExcelColumnAttribute))).OrderBy(x => x.GetCustomAttribute<ExcelColumnAttribute>().Index);
            foreach (var property in properties)
            {
                var attr = property.GetCustomAttribute<ExcelColumnAttribute>();

                if (attr.StructType == ExcelStructType.String)
                {
                    var row = _sheet.CreateRow(1 + _sheet.LastRowNum);
                    var cell = row.CreateCell(attr.Index);
                    cell.CellStyle = _contentStyle;
                }

                if (attr.StructType == ExcelStructType.DropList)
                {
                    SetCellDropdownList(_sheet, attr.Index, attr.Index, attr.DropList.Split(","));
                }
            }
        }

        public static void SetCellDropdownList(ISheet sheet, int firstcol, int lastcol, string[] vals)
        {
            //设置生成下拉框的行和列
            var cellRegions = new CellRangeAddressList(1, 65535, firstcol, lastcol);

            //设置 下拉框内容
            DVConstraint constraint = DVConstraint.CreateExplicitListConstraint(vals);

            //绑定下拉框和作用区域，并设置错误提示信息
            HSSFDataValidation dataValidate = new HSSFDataValidation(cellRegions, constraint);
            dataValidate.CreateErrorBox("输入不合法", "请输入或选择下拉列表中的值。");
            dataValidate.ShowPromptBox = true;

            sheet.AddValidationData(dataValidate);
        }

        public void ToFile(string filePath)
        {
            SetHeader();
            SetBody();

            using (var fs = new FileStream(filePath, FileMode.Create, FileAccess.Write))
            {
                _workbook.Write(fs);
            }
        }

        public void ToTemp(string filePath)
        {
            SetHeader();
            SetTemp();

            using (var fs = new FileStream(filePath, FileMode.Create, FileAccess.Write))
            {
                _workbook.Write(fs);
            }
        }
    }
}