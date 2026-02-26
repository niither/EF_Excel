using Microsoft.Data.SqlClient;
using OfficeOpenXml;
using System.Reflection;
using System.Reflection.Emit;
using System.Text;
using Z.Dapper.Plus;

namespace EF_DapperHW
{
    public class DynamicModelGenerator
    {
        public static void Main()
        {
            Console.OutputEncoding = Encoding.UTF8;
            string connectionString = "Server=localhost\\SQLEXPRESS;Database=Bookstore_network;Trusted_Connection=True;TrustServerCertificate=True;";
            string filePath = "C:/Users/user/Desktop/100mb.xlsx";

            var records = ReadExcel(filePath);
            var modelType = CreateDynamicModel(filePath);
            var dataList = MapDataToModel(records, modelType);

            string tableName = "ExcelTable";
            CreateTableInDatabase(tableName, modelType, connectionString);

            using (var connection = new SqlConnection(connectionString))
            {
                connection.Open();
                connection.BulkInsert(dataList);
                Console.WriteLine("Data added to SQl table");
            }
        }

        static List<Dictionary<string, string>> ReadExcel(string filePath)
        {
            ExcelPackage.License.SetNonCommercialPersonal("niither");

            var records = new List<Dictionary<string, string>>();

            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                Console.WriteLine($"Всього аркушів: {package.Workbook.Worksheets.Count}");

                foreach (var sheet in package.Workbook.Worksheets)
                {
                    Console.WriteLine($"Аркуш: {sheet.Name}");
                }

                if (package.Workbook.Worksheets.Count == 0)
                {
                    Console.WriteLine("Аркуші у файлі відсутні!");
                    return records;
                }

                var worksheet = package.Workbook.Worksheets[0];

                int rowCount = worksheet.Dimension.Rows;
                int colCount = worksheet.Dimension.Columns;

                var headers = new List<string>();
                for (int col = 1; col <= colCount; col++)
                {
                    headers.Add(worksheet.Cells[1, col].Text.Trim());
                }

                for (int row = 2; row <= rowCount; row++)
                {
                    var rowData = new Dictionary<string, string>();
                    for (int col = 1; col <= colCount; col++)
                    {
                        rowData[headers[col - 1]] = worksheet.Cells[row, col].Text.Trim();
                    }
                    records.Add(rowData);
                }
            }

            return records;
        }

        static Type CreateDynamicModel(string filePath)
        {
            var columns = new List<string>();

            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                var worksheet = package.Workbook.Worksheets[0];
                int colCount = worksheet.Dimension.Columns;

                for (int col = 1; col <= colCount; col++)
                {
                    columns.Add(worksheet.Cells[1, col].Text.Trim());
                }
            }

            if (columns.Count == 0)
            {
                throw new Exception("Файл Excel не містить заголовків!");
            }

            Console.WriteLine("Заголовки Excel: " + string.Join(", ", columns));

            var assemblyName = new AssemblyName("ExcelTableAssembly");
            var assemblyBuilder = AssemblyBuilder.DefineDynamicAssembly(assemblyName, AssemblyBuilderAccess.Run);
            var moduleBuilder = assemblyBuilder.DefineDynamicModule("MainModule");
            var typeBuilder = moduleBuilder.DefineType("ExcelTable", TypeAttributes.Public | TypeAttributes.Class);

            foreach (var column in columns)
            {
                var fieldBuilder = typeBuilder.DefineField("_" + column, typeof(string), FieldAttributes.Private);

                var propertyBuilder = typeBuilder.DefineProperty(column, PropertyAttributes.HasDefault, typeof(string), null);

                var getterMethodBuilder = typeBuilder.DefineMethod("get_" + column, MethodAttributes.Public, typeof(string), Type.EmptyTypes);
                var getterIl = getterMethodBuilder.GetILGenerator();
                getterIl.Emit(OpCodes.Ldarg_0);
                getterIl.Emit(OpCodes.Ldfld, fieldBuilder);
                getterIl.Emit(OpCodes.Ret);

                var setterMethodBuilder = typeBuilder.DefineMethod("set_" + column, MethodAttributes.Public, null, new[] { typeof(string) });
                var setterIl = setterMethodBuilder.GetILGenerator();
                setterIl.Emit(OpCodes.Ldarg_0);
                setterIl.Emit(OpCodes.Ldarg_1);
                setterIl.Emit(OpCodes.Stfld, fieldBuilder);
                setterIl.Emit(OpCodes.Ret);

                propertyBuilder.SetGetMethod(getterMethodBuilder);
                propertyBuilder.SetSetMethod(setterMethodBuilder);
            }

            return typeBuilder.CreateType();
        }
        static List<object> MapDataToModel(List<Dictionary<string, string>> records, Type modelType)
        {
            var dataList = new List<object>();

            foreach (var row in records)
            {
                if (row == null || row.Count == 0) continue;

                var instance = Activator.CreateInstance(modelType);

                foreach (var kvp in row)
                {
                    var property = modelType.GetProperty(kvp.Key);
                    if (property != null)
                    {
                        property.SetValue(instance, kvp.Value);
                    }
                }

                dataList.Add(instance);
            }

            return dataList;
        }

        static void CreateTableInDatabase(string tableName, Type modelType, string connectionString)
        {
            using (var connection = new SqlConnection(connectionString))
            {
                connection.Open();

                var columns = modelType.GetProperties()
                    .Select(prop => $"[{prop.Name}] NVARCHAR(MAX)")
                    .ToList();

                string createTableQuery = $@"
                    IF OBJECT_ID('{tableName}', 'U') IS NULL
                    CREATE TABLE {tableName} (
                        {string.Join(", ", columns)}
                    )";

                using (var command = new SqlCommand(createTableQuery, connection))
                {
                    command.ExecuteNonQuery();
                }
            }
        }
    }
}