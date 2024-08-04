using ExcelDataReader;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.IO;
using System.Linq;
using System.Web;
using System.Xml.Serialization;

namespace UtilityProject
{
    public class Helper
    {
        public T DeserializeToObject<T>(string filepath) where T : class
        {
            System.Xml.Serialization.XmlSerializer ser = new System.Xml.Serialization.XmlSerializer(typeof(T));

            using (StreamReader sr = new StreamReader(filepath))
            {
                return (T)ser.Deserialize(sr);
            }
        }

        public void SerializeToXml<T>(T anyobject, string xmlFilePath)
        {
            XmlSerializer xmlSerializer = new XmlSerializer(anyobject.GetType());

            using (StreamWriter writer = new StreamWriter(xmlFilePath))
            {
                xmlSerializer.Serialize(writer, anyobject);
            }
        }

        public DataTable ConvertExcelToDataTable(string filePath, bool isXlsx = false)
        {
            try
            {
                FileStream stream = null;
                DataTable dataTable = null;

                IExcelDataReader excelReader = null;

                using (stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
                {
                    excelReader = isXlsx ? ExcelReaderFactory.CreateOpenXmlReader(stream) : ExcelReaderFactory.CreateBinaryReader(stream);
                    DataSet result = excelReader.AsDataSet(new ExcelDataSetConfiguration()
                    {
                        ConfigureDataTable = (_) => new ExcelDataTableConfiguration()
                        {
                            UseHeaderRow = true
                        }
                    });

                    if (stream != null)
                        stream.Dispose();
                    if (excelReader != null)
                        excelReader.Dispose();

                    if (result != null && result.Tables.Count > 0)
                        dataTable = result.Tables[0];
                }

                return dataTable;
            }
            catch (Exception)
            {

                throw;
            }

        }

        public DataTable ConvertToDataTable<T>(IList<T> data)
        {
            PropertyDescriptorCollection properties = TypeDescriptor.GetProperties(typeof(T));
            DataTable table = new DataTable();
            foreach (PropertyDescriptor prop in properties)
                table.Columns.Add(prop.Name, Nullable.GetUnderlyingType(prop.PropertyType) ?? prop.PropertyType);
            foreach (T item in data)
            {
                DataRow row = table.NewRow();
                foreach (PropertyDescriptor prop in properties)
                    row[prop.Name] = prop.GetValue(item) ?? DBNull.Value;
                table.Rows.Add(row);
            }
            return table;
        }


    }
}
