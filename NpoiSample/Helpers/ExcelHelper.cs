using NPOI.HSSF.Converter;
using NPOI.HSSF.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Web;

namespace NpoiSample.Helpers
{
    public class ExcelHelper
    {
        public List<PropertyInfo> GetProperties(Type type, IEnumerable<string> propertyNames)
        {
            var allProperties = type.GetProperties();
            var properties = new List<PropertyInfo>();

            foreach (var propertyName in propertyNames)
            {
                var property = allProperties.First(p => p.Name == propertyName);

                properties.Add(property);
            }

            return properties;
        }

        public List<T> ReadData<T>(Stream stream, string fileName, List<PropertyInfo> columns)
        {
            var list = new List<T>();

            if (fileName.EndsWith("xls"))
            {
                var workbook = new HSSFWorkbook(stream);
                var sheet = workbook.GetSheetAt(0);

                if (sheet.PhysicalNumberOfRows <= 1)
                    return list;

                for (int i = 1; i < sheet.PhysicalNumberOfRows; i++)
                {
                    var item = Activator.CreateInstance<T>();

                    var row = sheet.GetRow(i);

                    var j = 0;

                    foreach (var column in columns)
                    {
                        var val = row.GetCell(j).StringCellValue;

                        if (string.IsNullOrWhiteSpace(val))
                        {
                            continue;
                        }
                        Type t = Nullable.GetUnderlyingType(column.PropertyType) ?? column.PropertyType;
                        column.SetValue(item, Convert.ChangeType(val, t), null);

                        j++;
                    }

                    list.Add(item);
                }
            }
            else if (fileName.EndsWith("xlsx"))
            {
                var workbook = new XSSFWorkbook(stream);
                var sheet = workbook.GetSheetAt(0);

                for (int i = 0; i < sheet.PhysicalNumberOfRows; i++)
                {
                    var item = Activator.CreateInstance<T>();

                    var row = sheet.GetRow(i);

                    var j = 0;

                    foreach (var column in columns)
                    {
                        var val = row.GetCell(j).StringCellValue;

                        if (string.IsNullOrWhiteSpace(val))
                        {
                            continue;
                        }
                        Type t = Nullable.GetUnderlyingType(column.PropertyType) ?? column.PropertyType;
                        column.SetValue(item, Convert.ChangeType(val, t), null);

                        j++;
                    }

                    list.Add(item);
                }
            }

            return list;
        }

        public HSSFWorkbook CreateXls<T>(List<T> items, List<PropertyInfo> properties)
        {
            // Create the workbook, sheet, and row
            var workbook = new HSSFWorkbook();
            var sheet = workbook.CreateSheet();

            var i = 0;

            // Header row
            {
                var j = 0;

                var row = sheet.CreateRow(i);

                foreach (var property in properties)
                {
                    row.CreateCell(j).SetCellValue(property.Name);

                    j++;
                }

                i++;
            }

            foreach (var item in items)
            {
                var j = 0;

                var row = sheet.CreateRow(i);

                foreach (var property in properties)
                {
                    row.CreateCell(j).SetCellValue(property.GetValue(item).ToString());

                    j++;
                }

                i++;
            }

            return workbook;
        }

        public XSSFWorkbook CreateXlsx<T>(List<T> items, List<PropertyInfo> properties)
        {
            // Create the workbook, sheet, and row
            var workbook = new XSSFWorkbook();
            var sheet = workbook.CreateSheet();

            var i = 0;

            // Header row
            {
                var j = 0;

                var row = sheet.CreateRow(i);

                foreach (var property in properties)
                {
                    row.CreateCell(j).SetCellValue(property.Name);

                    j++;
                }

                i++;
            }

            foreach (var item in items)
            {
                var j = 0;

                var row = sheet.CreateRow(i);

                foreach (var property in properties)
                {
                    row.CreateCell(j).SetCellValue(property.GetValue(item).ToString());

                    j++;
                }

                i++;
            }

            return workbook;
        }
    }
}
