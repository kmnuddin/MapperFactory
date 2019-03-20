using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Data.Odbc;
using System.Collections;
using System.Reflection;
using System.ComponentModel;
using ExcelFactory;

namespace AutoMapperFactory
{
    public static class MapperFactory<T> where T : class
    {

        private static List<T> objlist;
        private static string _filename;

        #region Private Methods
        private static T createInstance() => (T)Activator.CreateInstance(typeof(T));
        private static PropertyInfo[] GetPropertyInfo<T>() => typeof(T).GetProperties();
        private static DataTable Load_Work_Sheet_XLS()
        {
            DataTable dt = new DataTable();
            using (OdbcConnection conn = new OdbcConnection(getConnectionString()))
            {
                string query = "select * from [Main$]";
                OdbcCommand cmd = new OdbcCommand(query, conn);
                conn.Open();
                OdbcDataReader dataReader = cmd.ExecuteReader();

                dt.Load(dataReader);
                conn.Close();
            }
            return dt;
        }
        private static object convertType(string obj, Type sourceType, Type destinationType)
        {
            TypeConverter converter = TypeDescriptor.GetConverter(sourceType);
            if (converter.CanConvertTo(destinationType))
            {
                return converter.ConvertTo(obj, destinationType);
            }
            converter = TypeDescriptor.GetConverter(destinationType);
            if (converter.CanConvertFrom(sourceType))
            {
                return converter.ConvertFrom(obj);
            }
            return null;
        }
        
        #region Conn_String
        private static string getConnectionString()
        {
            Dictionary<string, string> props = new Dictionary<string, string>();

            // XLSX - Excel 2007, 2010, 2012, 2013
            props["Driver"] = "{Microsoft Excel Driver (*.xls, *.xlsx, *.xlsm, *.xlsb)}";
            //props["DriverId"] = "790";
            // props["DSN"] = "''";
            props["DBQ"] = _filename;



            // XLS - Excel 2003 and Older
            //props["Provider"] = "Microsoft.Jet.OLEDB.4.0";
            //props["Extended Properties"] = "Excel 8.0";
            //props["Data Source"] = "C:\\MyExcel.xls";

            StringBuilder sb = new StringBuilder();

            foreach (KeyValuePair<string, string> prop in props)
            {
                sb.Append(prop.Key);
                sb.Append('=');
                sb.Append(prop.Value);
                sb.Append(';');
            }

            return sb.ToString();
        }
        #endregion
        #endregion


        #region xls Mapping
        public static IEnumerable<T> Map_XLS(string filename)
        {
            _filename = filename;
            objlist = new List<T>();

            DataTable dt = Load_Work_Sheet_XLS();
            var prop_info = GetPropertyInfo<T>();

            for (int i = 0; i < dt.Rows.Count; i++)
            {
                T obj = createInstance();
                for (int j = 0; j < dt.Columns.Count; j++)
                {
                    prop_info[j].SetValue(obj, dt.Rows[i][j]);
                }
                objlist.Add(obj);
            }
            return objlist;
        }



        #endregion

        #region csv Mapping
        public static IEnumerable<T> Map_CSV(string[] data, MapperEnums.Inputs inputEnums)
        {
            objlist = new List<T>();
            var prop_info = GetPropertyInfo<T>();
            if (inputEnums == MapperEnums.Inputs.Regions)
            {
                for (int i = 1; i < data.Length; i++)
                {
                    string[] values = data[i].Split(new char[] { ',' });
                    T obj = createInstance();
                    for (int j = 0; j < values.Length; j++)
                    {
                        //var data_type = prop_info[j].GetType();
                        //var data_t = Cast(data[j], data_type);
                        var data_t = convertType(values[j], typeof(string), prop_info[j].PropertyType);
                        prop_info[j].SetValue(obj, data_t);
                    }

                    objlist.Add(obj);

                } 
            }
            if(inputEnums == MapperEnums.Inputs.Correlations)
            {
                string[,] valuePairs = new string[data.Length -1 , data.Length - 1];
                for(int i = 0; i < data.Length - 1; i++)
                {
                    string[] values = data[i].Split(new char[] { ',' });
                    
                    for(int j = 0; j < values.Length; j++)
                    {
                        if (values[j] == "0\r")
                            values[j] = "0";
                        valuePairs[i, j] = values[j];                        
                    }                  

                }

                for(int i = 1; i < data.Length - 1; i++)
                {
                    for(int j = 1; j < data.Length - 1; j++)
                    {
                        if (valuePairs[i,j] != "0")
                        {
                            T obj = createInstance();
                            prop_info[0].SetValue(obj, valuePairs[0, j]);
                            prop_info[1].SetValue(obj, valuePairs[i, 0]);
                            prop_info[2].SetValue(obj, convertType(valuePairs[i, j], typeof(string), prop_info[2].PropertyType));
                            objlist.Add(obj);
                        }
                    }


                }
            }
            return objlist;
        }
        #endregion




    }

}
