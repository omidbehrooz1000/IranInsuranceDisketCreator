using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Linq;


namespace Dbf
{
    public static class DbfFile
    {
        private static Func<string, string> _defaultCorrection = (i) => { return i.ArabicToFarsi(); };

        private static bool _correctFarsiCharacters = true;

        private static Dictionary<char, Func<byte[], object>> _mapFunctions;

        private const int _arabicWindowsEncoding = 1256;

        public static Encoding _arabicEncoding = Encoding.GetEncoding(_arabicWindowsEncoding);

        private static Encoding _currentEncoding = Encoding.GetEncoding(_arabicWindowsEncoding);

        private static Encoding _fieldsEncoding = Encoding.GetEncoding(_arabicWindowsEncoding);

        static DbfFile()
        {
            InitializeMapFunctions();
        }

        private static void InitializeMapFunctions()
        {
            _mapFunctions = new Dictionary<char, Func<byte[], object>>();

            _mapFunctions.Add('F', ToDouble);

            _mapFunctions.Add('O', ToDouble);
            _mapFunctions.Add('+', ToDouble);
            _mapFunctions.Add('I', ToInt);
            _mapFunctions.Add('Y', ToDecimal);

            _mapFunctions.Add('L', (input) => ConvertDbfLogicalValueToBoolean(input));

            //MapFunction.Add('D', (input) => new DateTime(int.Parse(new string(Encoding.ASCII.GetChars(input, 0, 4))),
            //                                                int.Parse(new string(Encoding.ASCII.GetChars(input, 4, 2))),
            //                                                int.Parse(new string(Encoding.ASCII.GetChars(input, 6, 2)))));

            _mapFunctions.Add('D', (input) =>
            {
                var first = new string(Encoding.ASCII.GetChars(input, 0, 4));
                var second = new string(Encoding.ASCII.GetChars(input, 4, 2));
                var third = new string(Encoding.ASCII.GetChars(input, 6, 2));

                if (string.IsNullOrEmpty(first + second + third) || string.IsNullOrWhiteSpace(first + second + third))
                {
                    return null;
                }
                else
                {
                    var year = int.Parse(first);
                    var month = int.Parse(second);
                    var day = int.Parse(third);

                    if (year + month + day == 0) //in the case; first : 0000, second : 00, third : 00
                    {
                        return null;
                    }

                    return new DateTime(year, month, day);
                }
            });


            _mapFunctions.Add('M', (input) => input);

            _mapFunctions.Add('B', (input) => input);

            _mapFunctions.Add('P', (input) => input);

            _mapFunctions.Add('N', ToDouble);

            _mapFunctions.Add('C', (input) => ConvertDbfGeneralToString(input, _currentEncoding));

            _mapFunctions.Add('G', (input) => ConvertDbfGeneralToString(input, _currentEncoding));

            _mapFunctions.Add('V', (input) => ConvertDbfGeneralToString(input, _currentEncoding));
        }


        private static byte[] GetIranSystemEncodingBytes(string value, byte[] array)
        {
            var truncatedString = value;

            var length = array.Length;

            IranSystemConvertor.ConvertWindowsPersianToDOS convertorClass = new IranSystemConvertor.ConvertWindowsPersianToDOS();
            List<byte> resultList = convertorClass.get_Unicode_To_IranSystem(value);
            if (resultList.Count > length)
            {
                //Truncate Scenario
                resultList.RemoveRange(length, resultList.Count - length);

                System.Diagnostics.Trace.WriteLine("Truncation occured in writing the dbf file");
                System.Diagnostics.Trace.WriteLine($"Original String: {value}");
                System.Diagnostics.Trace.WriteLine($"Truncated String: {truncatedString}");
                System.Diagnostics.Trace.WriteLine($"Lost String: {value.Replace(truncatedString, string.Empty)}");
                System.Diagnostics.Trace.WriteLine(string.Empty);
            }




            //Encoder en = encoding.GetEncoder().Convert(, 0, 0, null, 0, 0,, 0);
            //Consider using the Encoder.Convert method instead of GetByteCount.
            //The conversion method converts as much data as possible, and does 
            //throw an exception if the output buffer is too small.For continuous 
            //encoding of a stream, this method is often the best choice.
            resultList.ToArray().CopyTo(array, 0);
            return array;
        }

        private static byte[] GetBytes(string value, byte[] array, Encoding encoding)
        {
            var truncatedString = value;

            var length = array.Length;

            if (encoding.GetByteCount(value) > length)
            {
                //Truncate Scenario
                truncatedString = new string(value.TakeWhile((c, i) => encoding.GetByteCount(value.Substring(0, i + 1)) < length).ToArray());

                System.Diagnostics.Trace.WriteLine("Truncation occured in writing the dbf file");
                System.Diagnostics.Trace.WriteLine($"Original String: {value}");
                System.Diagnostics.Trace.WriteLine($"Truncated String: {truncatedString}");
                System.Diagnostics.Trace.WriteLine($"Lost String: {value.Replace(truncatedString, string.Empty)}");
                System.Diagnostics.Trace.WriteLine(string.Empty);
            }

            

            encoding.GetBytes(truncatedString, 0, truncatedString.Length, array, 0);

            //Encoder en = encoding.GetEncoder().Convert(, 0, 0, null, 0, 0,, 0);
            //Consider using the Encoder.Convert method instead of GetByteCount.
            //The conversion method converts as much data as possible, and does 
            //throw an exception if the output buffer is too small.For continuous 
            //encoding of a stream, this method is often the best choice.

            return array;
        }

        private static bool? ConvertDbfLogicalValueToBoolean(byte[] buffer)
        {
            string tempValue = Encoding.ASCII.GetString(buffer);

            if (tempValue.ToUpper().Equals("T") || tempValue.ToUpper().Equals("Y"))
            {
                return true;
            }
            else if (tempValue.ToUpper().Equals("F") || tempValue.ToUpper().Equals("N"))
            {
                return false;
            }
            else
            {
                return null;
            }
        }

        private static string ConvertDbfGeneralToString(byte[] buffer, Encoding encoding)
        {
            if (_correctFarsiCharacters)
            {
                return _defaultCorrection(encoding.GetString(buffer).Replace('\0', ' ').Trim());
            }
            else
            {

                return encoding.GetString(buffer).Replace('\0', ' ').Trim();
            }

        }

        private static readonly Func<byte[], object> ToDouble =
            (input) =>
            {
                //string value = Encoding.ASCII.GetString(input).Trim();
                //return string.IsNullOrEmpty(value) ? DBNull.Value : (object)double.Parse(value);
                double value;
                return double.TryParse(Encoding.ASCII.GetString(input), out value) ? (object)value : DBNull.Value;
            };

        private static readonly Func<byte[], object> ToInt =
            (input) =>
            {
                string value = Encoding.ASCII.GetString(input);
                return string.IsNullOrEmpty(value) ? DBNull.Value : (object)int.Parse(value);
            };

        private static readonly Func<byte[], object> ToDecimal =
            (input) =>
            {
                string value = Encoding.ASCII.GetString(input);
                return string.IsNullOrEmpty(value) ? DBNull.Value : (object)decimal.Parse(value);
            };

        private static short GetRecordLength(List<DbfFieldDescriptor> columns)
        {
            short result = 0;

            foreach (var item in columns)
            {
                result += item.Length;
            }

            result += 1; //Deletion Flag

            return result;
        }

        public static void ChangeEncoding(Encoding newEncoding)
        {
            _currentEncoding = newEncoding;

            InitializeMapFunctions();
        }

        public static List<DbfFieldDescriptor> GetDbfSchema(string dbfFileName)
        {
            System.IO.Stream stream = new System.IO.FileStream(dbfFileName, System.IO.FileMode.Open);

            System.IO.BinaryReader reader = new System.IO.BinaryReader(stream);

            byte[] buffer = reader.ReadBytes(Marshal.SizeOf(typeof(DbfHeader)));

            DbfHeader header = StreamHelper.ByteArrayToStructure<DbfHeader>(buffer);

            List<DbfFieldDescriptor> columns = new List<DbfFieldDescriptor>();

            if ((header.LengthOfHeader - 33) % 32 != 0) { throw new NotImplementedException(); }

            int numberOfFields = (header.LengthOfHeader - 33) / 32;

            for (int i = 0; i < numberOfFields; i++)
            {
                buffer = reader.ReadBytes(Marshal.SizeOf(typeof(DbfFieldDescriptor)));

                columns.Add(StreamHelper.ParseToStructure<DbfFieldDescriptor>(buffer));
            }

            reader.Close();

            stream.Close();

            return columns;
        }

        public static List<DbfFieldDescriptor> GetDbfSchema(string dbfFileName, Encoding encoding)
        {
            System.IO.Stream stream = new System.IO.FileStream(dbfFileName, System.IO.FileMode.Open);

            System.IO.BinaryReader reader = new System.IO.BinaryReader(stream);

            byte[] buffer = reader.ReadBytes(Marshal.SizeOf(typeof(DbfHeader)));

            DbfHeader header = StreamHelper.ByteArrayToStructure<DbfHeader>(buffer);

            List<DbfFieldDescriptor> columns = new List<DbfFieldDescriptor>();

            if ((header.LengthOfHeader - 33) % 32 != 0) { throw new NotImplementedException(); }

            int numberOfFields = (header.LengthOfHeader - 33) / 32;

            for (int i = 0; i < numberOfFields; i++)
            {
                buffer = reader.ReadBytes(Marshal.SizeOf(typeof(DbfFieldDescriptor)));

                columns.Add(DbfFieldDescriptor.Parse(buffer, encoding));
            }

            reader.Close();

            stream.Close();

            return columns;
        }


        public static Encoding TryDetectEncoding(string dbfFileName)
        {
            var cpgFile = GetCpgFileName(dbfFileName);

            if (!System.IO.File.Exists(cpgFile))
            {
                return null;
            }

            var encodingText = System.IO.File.ReadAllText(cpgFile);

            if (encodingText?.ToUpper()?.Trim() == "UTF-8" || encodingText?.ToUpper()?.Trim() == "UTF8")
            {
                return Encoding.UTF8;
            }
            else if (encodingText?.Contains("1256") == true)
            {
                return Dbf.DbfFile._arabicEncoding;
            }
            else
            {
                return null;
            }
        }

        //public static List<Dictionary<string, object>> Read(string dbfFileName, bool correctFarsiCharacters = true, Encoding dataEncoding = null, Encoding fieldHeaderEncoding = null)
        public static EsriAttributeDictionary Read(string dbfFileName, bool correctFarsiCharacters = true, Encoding dataEncoding = null, Encoding fieldHeaderEncoding = null)
        {
            dataEncoding = dataEncoding ?? (TryDetectEncoding(dbfFileName) ?? Encoding.UTF8);

            ChangeEncoding(dataEncoding);

            //if (tryDetectEncoding)
            //{
            //    Encoding encoding = TryDetectEncoding(dbfFileName) ?? dataEncoding;

            //    ChangeEncoding(encoding);
            //}
            //else
            //{
            //    ChangeEncoding(dataEncoding);
            //}

            DbfFile._fieldsEncoding = fieldHeaderEncoding ?? _arabicEncoding;

            DbfFile._correctFarsiCharacters = correctFarsiCharacters;

            System.IO.Stream stream = new System.IO.FileStream(dbfFileName, System.IO.FileMode.Open);

            System.IO.BinaryReader reader = new System.IO.BinaryReader(stream);

            byte[] buffer = reader.ReadBytes(Marshal.SizeOf(typeof(DbfHeader)));

            DbfHeader header = StreamHelper.ByteArrayToStructure<DbfHeader>(buffer);

            List<DbfFieldDescriptor> fields = new List<DbfFieldDescriptor>();

            if ((header.LengthOfHeader - 33) % 32 != 0) { throw new NotImplementedException(); }

            int numberOfFields = (header.LengthOfHeader - 33) / 32;

            for (int i = 0; i < numberOfFields; i++)
            {
                buffer = reader.ReadBytes(Marshal.SizeOf(typeof(DbfFieldDescriptor)));

                fields.Add(DbfFieldDescriptor.Parse(buffer, DbfFile._fieldsEncoding));
            }


            //System.Data.DataTable result = MakeTableSchema(tableName, columns);

            var attributes = new List<Dictionary<string, object>>(header.NumberOfRecords);

            ((FileStream)reader.BaseStream).Seek(header.LengthOfHeader, SeekOrigin.Begin);

            for (int i = 0; i < header.NumberOfRecords; i++)
            {
                // First we'll read the entire record into a buffer and then read each field from the buffer
                // This helps account for any extra space at the end of each record and probably performs better
                buffer = reader.ReadBytes(header.LengthOfEachRecord);
                BinaryReader recordReader = new BinaryReader(new MemoryStream(buffer));

                // All dbf field records begin with a deleted flag field. Deleted - 0x2A (asterisk) else 0x20 (space)
                if (recordReader.ReadChar() == '*')
                {
                    continue;
                }

                Dictionary<string, object> values = new Dictionary<string, object>();

                for (int j = 0; j < fields.Count; j++)
                {
                    int fieldLenth = fields[j].Length;

                    //values[j] = MapFunction[columns[j].Type](recordReader.ReadBytes(fieldLenth));
                    values.Add(fields[j].Name, _mapFunctions[fields[j].Type](recordReader.ReadBytes(fieldLenth)));
                }

                recordReader.Close();

                attributes.Add(values);
            }

            reader.Close();

            stream.Close();

            return new EsriAttributeDictionary(attributes, fields);
        }

        public static object[][] ReadToObject(string dbfFileName, string tableName, bool correctFarsiCharacters = true, Encoding dataEncoding = null, Encoding fieldHeaderEncoding = null)
        {
            dataEncoding = dataEncoding ?? (TryDetectEncoding(dbfFileName) ?? Encoding.UTF8);

            ChangeEncoding(dataEncoding);

            DbfFile._fieldsEncoding = fieldHeaderEncoding ?? _arabicEncoding;

            DbfFile._correctFarsiCharacters = correctFarsiCharacters;


            System.IO.Stream stream = new System.IO.FileStream(dbfFileName, System.IO.FileMode.Open);

            System.IO.BinaryReader reader = new System.IO.BinaryReader(stream);

            byte[] buffer = reader.ReadBytes(Marshal.SizeOf(typeof(DbfHeader)));

            DbfHeader header = StreamHelper.ByteArrayToStructure<DbfHeader>(buffer);

            List<DbfFieldDescriptor> columns = new List<DbfFieldDescriptor>();

            if ((header.LengthOfHeader - 33) % 32 != 0) { throw new NotImplementedException(); }

            int numberOfFields = (header.LengthOfHeader - 33) / 32;

            for (int i = 0; i < numberOfFields; i++)
            {
                buffer = reader.ReadBytes(Marshal.SizeOf(typeof(DbfFieldDescriptor)));

                columns.Add(DbfFieldDescriptor.Parse(buffer, DbfFile._fieldsEncoding));
            }

            //System.Data.DataTable result = MakeTableSchema(tableName, columns);
            var result = new object[header.NumberOfRecords][];

            ((FileStream)reader.BaseStream).Seek(header.LengthOfHeader, SeekOrigin.Begin);

            for (int i = 0; i < header.NumberOfRecords; i++)
            {
                // First we'll read the entire record into a buffer and then read each field from the buffer
                // This helps account for any extra space at the end of each record and probably performs better
                buffer = reader.ReadBytes(header.LengthOfEachRecord);
                BinaryReader recordReader = new BinaryReader(new MemoryStream(buffer));

                // All dbf field records begin with a deleted flag field. Deleted - 0x2A (asterisk) else 0x20 (space)
                if (recordReader.ReadChar() == '*')
                {
                    continue;
                }

                object[] values = new object[columns.Count];

                for (int j = 0; j < columns.Count; j++)
                {
                    int fieldLenth = columns[j].Length;

                    values[j] = _mapFunctions[columns[j].Type](recordReader.ReadBytes(fieldLenth));
                }

                recordReader.Close();

                result[i] = values;
            }

            reader.Close();

            stream.Close();

            return result;

            //ChangeEncoding(dataEncoding);

            //DbfFile._fieldsEncoding = fieldHeaderEncoding;

            //DbfFile._correctFarsiCharacters = correctFarsiCharacters;

            //return ReadToObject(dbfFileName, tableName);
        }

        //public static object[][] ReadToObject(string dbfFileName, string tableName)
        //{

        //}

        public static void Write(string fileName, int numberOfRecords, bool overwrite = false)
        {
            List<int> attributes = Enumerable.Range(0, numberOfRecords).ToList();

            List<ObjectToDbfTypeMap<int>> mapping = new List<ObjectToDbfTypeMap<int>>() { new ObjectToDbfTypeMap<int>(DbfFieldDescriptors.GetIntegerField("Id"), i => i) };

            Write(fileName,
                attributes,
                mapping,
                Encoding.ASCII,
                overwrite);

            //Write(fileName,
            //    attributes,
            //    new List<Func<int, object>>() { i => i },
            //    new List<DbfFieldDescriptor>() { DbfFieldDescriptors.GetIntegerField("Id") },
            //    Encoding.ASCII,
            //    overwrite);

        }

        //public static void Write<T>(string dbfFileName,
        //                                IEnumerable<T> values,
        //                                List<Func<T, object>> mapping,
        //                                List<DbfFieldDescriptor> columns,
        //                                Encoding encoding,
        //                                bool overwrite = false)
        //{

        //}

        public static void Write<T>(string dbfFileName,
                                        IEnumerable<T> values,
                                        List<ObjectToDbfTypeMap<T>> mapping,
                                        Encoding encoding,
                                        bool overwrite = false)
        {
            //Write(dbfFileName, values, mapping.Select(m => m.MapFunction).ToList(), mapping.Select(m => m.FieldType).ToList(), encoding, overwrite);

            var columns = mapping.Select(m => m.FieldType).ToList();

            int control = 0;
            try
            {
                //if (columns.Count != mapping.Count)
                //{
                //    throw new NotImplementedException();
                //}

                //var mode = overwrite ? System.IO.FileMode.Create : System.IO.FileMode.CreateNew;
                var mode = GetMode(dbfFileName, overwrite);

                System.IO.Stream stream = new System.IO.FileStream(dbfFileName, mode);

                System.IO.BinaryWriter writer = new System.IO.BinaryWriter(stream);

                DbfHeader header = new DbfHeader(values.Count(), mapping.Count, GetRecordLength(columns), encoding);

                writer.Write(StreamHelper.StructureToByteArray(header));

                foreach (var item in columns)
                {
                    writer.Write(StreamHelper.StructureToByteArray(item));
                }

                //Terminator
                writer.Write(byte.Parse("0D", System.Globalization.NumberStyles.HexNumber));

                for (int i = 0; i < values.Count(); i++)
                {
                    control = i;
                    // All dbf field records begin with a deleted flag field. Deleted - 0x2A (asterisk) else 0x20 (space)
                    writer.Write(byte.Parse("20", System.Globalization.NumberStyles.HexNumber));

                    for (int j = 0; j < mapping.Count; j++)
                    {
                        byte[] temp = new byte[columns[j].Length];

                        object value = mapping[j].MapFunction(values.ElementAt(i));

                        if (value is DateTime dt)
                        {
                            value = dt.ToString("yyyyMMdd");
                        }

                        if (value != null)
                        {
                            //encoding.GetBytes(value.ToString(), 0, value.ToString().Length, temp, 0);
                            temp = GetBytes(value.ToString(), temp, encoding);
                        }

                        //string tt = encoding.GetString(temp);
                        //var le = tt.Length;
                        writer.Write(temp);
                    }
                }

                //End of file
                writer.Write(byte.Parse("1A", System.Globalization.NumberStyles.HexNumber));

                writer.Close();

                stream.Close();

                System.IO.File.WriteAllText(GetCpgFileName(dbfFileName), encoding.BodyName);

            }
            catch (Exception ex)
            {
                string message = ex.Message;

                string m2 = message + " " + control.ToString();

            }
        }

        public static void Write<T>(string dbfFileName,
                                       IEnumerable<T> values,
                                       ObjectToDfbFields<T> mapping,
                                       Encoding encoding,
                                       bool overwrite = false)
        {
            //Write(dbfFileName, values, mapping.Select(m => m.MapFunction).ToList(), mapping.Select(m => m.FieldType).ToList(), encoding, overwrite);

            var columns = mapping.Fields;

            int control = 0;

            try
            {
                //if (columns.Count != mapping.Count)
                //{
                //    throw new NotImplementedException();
                //}

                //var mode = overwrite ? System.IO.FileMode.Create : System.IO.FileMode.CreateNew;
                var mode = GetMode(dbfFileName, overwrite);

                System.IO.Stream stream = new System.IO.FileStream(dbfFileName, mode);

                System.IO.BinaryWriter writer = new System.IO.BinaryWriter(stream);

                DbfHeader header = new DbfHeader(values.Count(), columns.Count, GetRecordLength(columns), encoding);

                writer.Write(StreamHelper.StructureToByteArray(header));

                foreach (var item in columns)
                {
                    writer.Write(StreamHelper.StructureToByteArray(item));
                }

                //Terminator
                writer.Write(byte.Parse("0D", System.Globalization.NumberStyles.HexNumber));

                for (int i = 0; i < values.Count(); i++)
                {
                    control = i;
                    // All dbf field records begin with a deleted flag field. Deleted - 0x2A (asterisk) else 0x20 (space)
                    writer.Write(byte.Parse("20", System.Globalization.NumberStyles.HexNumber));

                    var fieldValues = mapping.ExtractAttributesFunc(values.ElementAt(i));

                    for (int j = 0; j < columns.Count; j++)
                    {
                        byte[] temp = new byte[columns[j].Length];

                        if (fieldValues[j] != null)
                        {
                            //encoding.GetBytes(value.ToString(), 0, value.ToString().Length, temp, 0);
                            temp = GetBytes(fieldValues[j]?.ToString(), temp, encoding);
                        }

                        //string tt = encoding.GetString(temp);
                        //var le = tt.Length;
                        writer.Write(temp);
                    }
                }

                //End of file
                writer.Write(byte.Parse("1A", System.Globalization.NumberStyles.HexNumber));

                writer.Close();

                stream.Close();

                System.IO.File.WriteAllText(GetCpgFileName(dbfFileName), encoding.BodyName);

            }
            catch (Exception ex)
            {
                string message = ex.Message;

                string m2 = message + " " + control.ToString();

            }
        }


        public static void Write(string dbfFileName,
                                    List<Dictionary<string, object>> attributes,
                                    Encoding encoding,
                                    bool overwirte = false)
        {
            if (attributes == null || attributes.Count < 1)
            {
                return;
            }

            //make schema
            var columns = MakeDbfFields(attributes.First());


            List<ObjectToDbfTypeMap<Dictionary<string, object>>> mapping = new List<ObjectToDbfTypeMap<Dictionary<string, object>>>();

            var counter = 0;

            foreach (var item in attributes.First())
            {
                mapping.Add(new ObjectToDbfTypeMap<Dictionary<string, object>>(columns[counter], d => d[item.Key]));
            }

            Write(dbfFileName, attributes, mapping, encoding, overwirte);


            //1397.08.27
            //List<Func<Dictionary<string, object>, object>> mappings = new List<Func<Dictionary<string, object>, object>>();
            //foreach (var item in attributes.First())
            //{
            //    mappings.Add(d => d[item.Key]);
            //}
            //Write(dbfFileName, attributes, mappings, columns, encoding, overwirte);
        }

        public static List<DbfFieldDescriptor> MakeDbfFields(Dictionary<string, object> dictionary)
        {
            List<DbfFieldDescriptor> result = new List<DbfFieldDescriptor>();

            foreach (var item in dictionary)
            {
                result.Add(new DbfFieldDescriptor(item.Key, 'C', 255, 0));
            }

            return result;
        }

        public static string GetCpgFileName(string shpFileName)
        {
            return System.IO.Path.ChangeExtension(shpFileName, "cpg");
        }

        public static System.IO.FileMode GetMode(string fileName, bool overwrite)
        {
            return System.IO.File.Exists(fileName) && overwrite ? System.IO.FileMode.Create : System.IO.FileMode.CreateNew;
        }

        #region DataTable

        private static List<DbfFieldDescriptor> MakeDbfFields(System.Data.DataColumnCollection columns)
        {
            List<DbfFieldDescriptor> result = new List<DbfFieldDescriptor>();

            foreach (System.Data.DataColumn item in columns)
            {
                result.Add(new DbfFieldDescriptor(item.ColumnName, 'C', (byte)item.MaxLength, 0));
            }

            return result;
        }

        public static System.Data.DataTable MakeTableSchema(string tableName, List<DbfFieldDescriptor> columns)
        {
            System.Data.DataTable result = new System.Data.DataTable(tableName);

            foreach (DbfFieldDescriptor item in columns)
            {
                switch (char.ToUpper(item.Type))
                {
                    case 'F':
                    case 'O':
                    case '+':
                        result.Columns.Add(item.Name, typeof(double));
                        break;

                    case 'I':
                        result.Columns.Add(item.Name, typeof(int));
                        break;

                    case 'Y':
                        result.Columns.Add(item.Name, typeof(decimal));
                        break;

                    case 'L':
                        result.Columns.Add(item.Name, typeof(bool));
                        break;

                    case 'D':
                    case 'T':
                    case '@':
                        result.Columns.Add(item.Name, typeof(DateTime));
                        break;

                    case 'M':
                    case 'B':
                    case 'P':
                        result.Columns.Add(item.Name, typeof(byte[]));
                        break;

                    case 'N':
                        if (item.DecimalCount == 0)
                            result.Columns.Add(item.Name, typeof(int));
                        else
                            result.Columns.Add(item.Name, typeof(double));
                        break;

                    case 'C':
                    case 'G':
                    case 'V':
                    case 'X':
                    default:
                        result.Columns.Add(item.Name, typeof(string));
                        break;
                }
            }

            return result;
        }

        //Read
        public static System.Data.DataTable Read(string dbfFileName, string tableName, Encoding dataEncoding, Encoding fieldHeaderEncoding, bool correctFarsiCharacters)
        {
            ChangeEncoding(dataEncoding);

            DbfFile._fieldsEncoding = fieldHeaderEncoding;

            DbfFile._correctFarsiCharacters = correctFarsiCharacters;

            return Read(dbfFileName, tableName);
        }

        public static System.Data.DataTable Read(string dbfFileName, string tableName)
        {
            Stream stream = new FileStream(dbfFileName, System.IO.FileMode.Open);

            BinaryReader reader = new BinaryReader(stream);

            byte[] buffer = reader.ReadBytes(Marshal.SizeOf(typeof(DbfHeader)));

            DbfHeader header = StreamHelper.ByteArrayToStructure<DbfHeader>(buffer);

            List<DbfFieldDescriptor> columns = new List<DbfFieldDescriptor>();

            if ((header.LengthOfHeader - 33) % 32 != 0) { throw new NotImplementedException(); }

            int numberOfFields = (header.LengthOfHeader - 33) / 32;

            for (int i = 0; i < numberOfFields; i++)
            {
                buffer = reader.ReadBytes(Marshal.SizeOf(typeof(DbfFieldDescriptor)));

                columns.Add(DbfFieldDescriptor.Parse(buffer, DbfFile._fieldsEncoding));
            }

            System.Data.DataTable result = MakeTableSchema(tableName, columns);

            ((FileStream)reader.BaseStream).Seek(header.LengthOfHeader, SeekOrigin.Begin);

            for (int i = 0; i < header.NumberOfRecords; i++)
            {
                // First we'll read the entire record into a buffer and then read each field from the buffer
                // This helps account for any extra space at the end of each record and probably performs better
                buffer = reader.ReadBytes(header.LengthOfEachRecord);

                BinaryReader recordReader = new BinaryReader(new MemoryStream(buffer));

                // All dbf field records begin with a deleted flag field. Deleted - 0x2A (asterisk) else 0x20 (space)
                if (recordReader.ReadChar() == '*')
                {
                    continue;
                }

                object[] values = new object[columns.Count];

                for (int j = 0; j < columns.Count; j++)
                {
                    int fieldLenth = columns[j].Length;

                    values[j] = _mapFunctions[columns[j].Type](recordReader.ReadBytes(fieldLenth));
                }

                recordReader.Close();

                result.Rows.Add(values);
            }

            reader.Close();

            stream.Close();

            return result;
        }

        //Write


        public static void Write(string fileName, System.Data.DataTable table, Encoding encoding, bool overwrite = false, bool? useIranSystemEncoding = false)
        {
            var mode = GetMode(fileName, overwrite);

            Stream stream = new FileStream(fileName, mode);

            BinaryWriter writer = new BinaryWriter(stream);

            List<DbfFieldDescriptor> columns = MakeDbfFields(table.Columns);

            DbfHeader header = new DbfHeader(table.Rows.Count, table.Columns.Count, GetRecordLength(columns), encoding);

            writer.Write(StreamHelper.StructureToByteArray(header));

            foreach (var item in columns)
            {
                writer.Write(StreamHelper.StructureToByteArray(item));
            }

            //Terminator
            writer.Write(byte.Parse("0D", System.Globalization.NumberStyles.HexNumber));

            for (int i = 0; i < table.Rows.Count; i++)
            {
                // All dbf field records begin with a deleted flag field. Deleted - 0x2A (asterisk) else 0x20 (space)
                writer.Write(byte.Parse("20", System.Globalization.NumberStyles.HexNumber));

                for (int j = 0; j < table.Columns.Count; j++)
                {
                    byte[] temp = new byte[columns[j].Length];

                    string value = table.Rows[i][j].ToString().Trim();

                    //encoding.GetBytes(value, 0, value.Length, temp, 0);
                    //writer.Write(temp);

                    if (useIranSystemEncoding.HasValue && Convert.ToBoolean(useIranSystemEncoding))
                    {
                        GetIranSystemEncodingBytes(value, temp);
                        writer.Write(temp);
                    }
                    else
                    {
                        writer.Write(GetBytes(value, temp, encoding));
                    }
                }
            }

            //End of file
            writer.Write(byte.Parse("1A", System.Globalization.NumberStyles.HexNumber));

            writer.Close();

            stream.Close();
        }

        #endregion
    }
}


namespace IranSystemConvertor
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Text;
    using System.Text.RegularExpressions;

    public enum TextEncoding
    {
        Arabic1256 = 1256,
        CP1252 = 1252
    }

    public static class ConvertTo
    {
        #region private Members (3)

        // متغیری برای نگهداری اعدادی که در رشته ایران سیستم وجود دارند
        static Stack<string> NumbersInTheString;

        // کد کاراکترها در ایران سیستم و معادل آنها در عربی 1256
        static Dictionary<byte, byte> CharactersMapper = new Dictionary<byte, byte>
        {
        {128,48}, // 0
        {129,49}, // 1
        {130,50}, // 2
        {131,51}, // 3
        {132,52}, // 4
        {133,53}, // 5
        {134,54}, // 6
        {135,55}, // 7
        {136,56}, // 8
        {137,57}, // 9
        {138,161}, // ،
        {139,220}, // -
        {140,191}, // ؟
        {141,194}, // آ
        {142,196}, // ﺋ
        {143,154}, // ء
        {144,199}, // ﺍ
        {145,199}, // ﺎ
        {146,200}, // ﺏ
        {147,200}, // ﺑ
        {148,129}, // ﭖ
        {149,129}, // ﭘ
        {150,202}, // ﺕ
        {151,202}, // ﺗ
        {152,203}, // ﺙ
        {153,203}, // ﺛ
        {154,204}, //ﺝ
        {155,204},// ﺟ
        {156,141},//ﭼ
        {157,141},//ﭼ
        {158,205},//ﺡ
        {159,205},//ﺣ
        {160,206},//ﺥ
        {161,206},//ﺧ
        {162,207},//د
        {163,208},//ذ
        {164,209},//ر
        {165,210},//ز
        {166,142},//ژ
        {167,211},//ﺱ
        {168,211},//ﺳ
        {169,212},//ﺵ
        {170,212},//ﺷ
        {171,213},//ﺹ
        {172,213},//ﺻ
        {173,214},//ﺽ
        {174,214},//ﺿ
        {175,216},//ط
        {224,217},//ظ
        {225,218},//ﻉ
        {226,218},//ﻊ
        {227,218},//ﻌ
        {228,218},//ﻋ
        {229,219},//ﻍ
        {230,219},//ﻎ
        {231,219},//ﻐ
        {232,219},//ﻏ
        {233,221},//ﻑ
        {234,221},//ﻓ
        {235,222},//ﻕ
        {236,222},//ﻗ
        {237,152},//ﮎ
        {238,152},//ﮐ
        {239,144},//ﮒ
        {240,144},//ﮔ
        {241,225},//ﻝ
        {242,225},//ﻻ
        {243,225},//ﻟ
        {244,227},//ﻡ
        {245,227},//ﻣ
        {246,228},//ﻥ
        {247,228},//ﻧ
        {248,230},//و
        {249,229},//ﻩ
        {250,229},//ﻬ
        {251,170},//ﻫ
        {252,236},//ﯽ
        {253,237},//ﯼ
        {254,237},//ﯾ
        {255,160} // فاصله
        };

        /// <summary>
        /// لیست کاراکترهایی که بعد از آنها باید یک فاصله اضافه شود
        /// </summary>
        static byte[] charactersWithSpaceAfter = {
                                             146, // ب
                                             148, // پ
                                             150, // ت
                                             152, // ث
                                             154, // ج
                                             156, // چ
                                             158, // ح
                                             160, // خ
                                             167, // س
                                             169, // ش
                                             171, // ص
                                             173, // ض
                                             225, // ع
                                             229, // غ
                                             233, // ف
                                             235, // ق
                                             237, // ک
                                             239, // گ
                                             241, // ل
                                             244, // م
                                             246, // ن
                                             249, // ه
                                             252, //ﯽ
                                             253 // ی
                                         };


        #endregion

        /// <summary>
        /// تبدیل یک رشته ایران سیستم به یونیکد با استفاده از عربی 1256
        /// </summary>
        /// <param name="iranSystemEncodedString">رشته ایران سیستم</param>
        /// <returns></returns>
        [Obsolete("بهتر است از UnicodeFrom استفاده کنید")]
        public static string Unicode(string iranSystemEncodedString)
        {
            return UnicodeFrom(TextEncoding.Arabic1256, iranSystemEncodedString);
        }

        /// <summary>
        /// تبدیل یک رشته ایران سیستم به یونیکد
        /// </summary>
        /// <param name="textEncoding">کدپیج رشته ایران سیستم</param>
        /// <param name="iranSystemEncodedString">رشته ایران سیستم</param>
        /// <returns></returns>
        public static string UnicodeFrom(TextEncoding textEncoding, string iranSystemEncodedString)
        {
            // حذف فاصله های موجود در رشته
            iranSystemEncodedString = iranSystemEncodedString.Replace(" ", "");


            /// بازگشت در صورت خالی بودن رشته
            if (string.IsNullOrWhiteSpace(iranSystemEncodedString))
            {
                return string.Empty;
            }

            // در صورتی که رشته تماماً عدد نباشد
            if (!IsNumber(iranSystemEncodedString))
            {
                /// تغییر ترتیب کاراکترها از آخر به اول 
                iranSystemEncodedString = Reverse(iranSystemEncodedString);

                /// خارج کردن اعداد درون رشته
                iranSystemEncodedString = ExcludeNumbers(iranSystemEncodedString);
            }

            // وهله سازی از انکودینگ صحیح برای تبدیل رشته ایران سیستم به بایت
            Encoding encoding = Encoding.GetEncoding((int)textEncoding);

            // تبدیل رشته به بایت
            byte[] stringBytes = encoding.GetBytes(iranSystemEncodedString.Trim());


            // آرایه ای که بایت های معادل را در آن قرار می دهیم
            // مجموع تعداد بایت های رشته + بایت های اضافی محاسبه شده 
            byte[] newStringBytes = new byte[stringBytes.Length + CountCharactersRequireTwoBytes(stringBytes)];

            int index = 0;

            // بررسی هر بایت و پیدا کردن بایت (های) معادل آن
            for (int i = 0; i < stringBytes.Length; ++i)
            {
                byte charByte = stringBytes[i];

                // اگر جز 128 بایت اول باشد که نیازی به تبدیل ندارد چون کد اسکی است
                if (charByte < 128)
                {
                    newStringBytes[index] = charByte;
                }
                else
                {
                    // اگر جز حروف یا اعداد بود معادلش رو قرار می دیم
                    if (CharactersMapper.ContainsKey(charByte))
                    {
                        newStringBytes[index] = CharactersMapper[charByte];
                    }
                }

                // اگر کاراکتر ایران سیستم "لا" بود چون کاراکتر متناظرش در عربی 1256 "ل" است و باید یک "ا" هم بعدش اضافه کنیم
                if (charByte == 242)
                {
                    newStringBytes[++index] = 199;
                }

                // اگر کاراکتر یکی از انواعی بود که بعدشان باید یک فاصله باشد
                // و در عین حال آخرین کاراکتر رشته نبود
                if (charactersWithSpaceAfter.Contains(charByte) && Array.IndexOf(stringBytes, charByte) != stringBytes.Length - 1)
                {
                    // یک فاصله بعد ان اضافه می کنیم
                    newStringBytes[++index] = 32;
                }

                index += 1;
            }

            // تبدیل به رشته و ارسال به فراخواننده
            string convertedString = Encoding.GetEncoding(1256).GetString(newStringBytes);


            return IncludeNumbers(convertedString);
        }

        #region Private Methods (4)

        /// <summary>
        /// رشته ارسال شده تنها حاوی اعداد است یا نه
        /// </summary>
        /// <param name="str"></param>
        /// <returns></returns>
        static bool IsNumber(string str)
        {
            return Regex.IsMatch(str, @"^[\d]+$");
        }

        /// <summary>
        ///  محاسبه تعداد کاراکترهایی که بعد از آنها یک کاراکتر باید اضافه شود
        ///  شامل کاراکتر لا
        ///  و کاراکترهای غیرچسبان تنها در صورتی که کاراکتر پایانی رشته نباشند
        /// </summary>
        /// <param name="content"></param>
        /// <returns></returns>
        static int CountCharactersRequireTwoBytes(byte[] irTextBytes)
        {
            return (from b in irTextBytes
                    where (
                    charactersWithSpaceAfter.Contains(b) // یکی از حروف غیرچسبان باشد
                    && Array.IndexOf(irTextBytes, b) != irTextBytes.Length - 1) // و کاراکتر آخر هم نباشد
                    || b == 242 // یا کاراکتر لا باشد
                    select b).Count();
        }

        /// <summary>
        /// خارج کردن اعدادی که در رشته ایران سیستم قرار دارند
        /// </summary>
        /// <param name="iranSystemString"></param>
        /// <returns></returns>
        static string ExcludeNumbers(string iranSystemString)
        {
            /// گرفتن لیستی از اعداد درون رشته
            NumbersInTheString = new Stack<string>(Regex.Split(iranSystemString, @"\D+"));

            /// جایگزین کردن اعداد با یک علامت جایگزین
            /// در نهایت بعد از تبدیل رشته اعداد به رشته اضافه می شوند
            return Regex.Replace(iranSystemString, @"\d+", "#");
        }

        /// <summary>
        /// اضافه کردن اعداد جدا شده پس از تبدیل رشته
        /// </summary>
        /// <param name="convertedString"></param>
        /// <returns></returns>
        static string IncludeNumbers(string convertedString)
        {
            while (convertedString.IndexOf("#") >= 0)
            {
                string number = Reverse(NumbersInTheString.Pop());
                if (!string.IsNullOrWhiteSpace(number))
                {
                    int index = convertedString.IndexOf("#");

                    convertedString = convertedString.Remove(index, 1);
                    convertedString = convertedString.Insert(index, number);
                }
            }

            return convertedString;
        }

        static string Reverse(string s)
        {
            char[] charArray = s.ToCharArray();
            Array.Reverse(charArray);
            return new string(charArray);
        }

        #endregion

    }

    //public class IranianSystemEncoding
    //{
    //    static char[] ByteToChar;
    //    static Byte[][] CharToByte;

    //    static IranianSystemEncoding()
    //    {
    //        InitializeData();
    //    }

    //    static void InitializeData()
    //    {
    //        var iranSystem = new int[] { 0x06F0, 0x06F1, 0x06F2, 0x06F3, 0x06F4, 0x06F5, 0x06F6, 0x06F7, 0x06F8, 0x06F9, 0x060C, 0x0640, 0x061F, 0xFE81, 0xFE8B, 0x0621, 0xFE8D, 0xFE8E, 0xFE8F, 0xFE91, 0xFB56, 0xFB58, 0xFE95, 0xFE97, 0xFE99, 0xFE9B, 0xFE9D, 0xFE9F, 0xFB7C, 0xFB7C, 0xFEA1, 0xFEA3, 0xFEA5, 0xFEA7, 0x062F, 0x0630, 0x0631, 0x0632, 0x0698, 0xFEB1, 0xFEB3, 0xFEB5, 0xFEB7, 0xFEB9, 0xFEBB, 0xFEBD, 0xFEBF, 0x0637, 0x2591, 0x2592, 0x2593, 0x2502, 0x2524, 0x2561, 0x2562, 0x2556, 0x2555, 0x2563, 0x2551, 0x2557, 0x255D, 0x255C, 0x255B, 0x2510, 0x2514, 0x2534, 0x252C, 0x251C, 0x2500, 0x253C, 0x255E, 0x255F, 0x255A, 0x2554, 0x2569, 0x2566, 0x2560, 0x2550, 0x256C, 0x2567, 0x2568, 0x2564, 0x2565, 0x2559, 0x2558, 0x2552, 0x2553, 0x256B, 0x256A, 0x2518, 0x250C, 0x2588, 0x2584, 0x258C, 0x2590, 0x2580, 0x0638, 0xFEC9, 0xFECA, 0xFECC, 0xFECB, 0xFECD, 0xFECE, 0xFED0, 0xFECF, 0xFED1, 0xFED3, 0xFED5, 0xFED7, 0xFB8E, 0xFB90, 0xFB92, 0xFB94, 0xFEDD, 0xFEFB, 0xFEDF, 0xFEE1, 0xFEE3, 0xFEE5, 0xFEE7, 0x0648, 0xFEE9, 0xFEEC, 0xFEEB, 0xFBFD, 0xFBFC, 0xFBFE, 0x00A0 };
    //        ByteToChar = new char[256];
    //        // ascii first
    //        for (int i = 0; i < 128; i++) ByteToChar[i] = (char)i;
    //        // non-ascii
    //        for (int i = 128; i < 256; i++) ByteToChar[i] = (char)iranSystem[i - 128];

    //        // ok now reverse
    //        CharToByte = new Byte[256][];
    //        for (int i = 0; i < 256; i++)
    //        {
    //            char ch = (char)ByteToChar[i];
    //            var low = ch & 0xff;
    //            var high = ch >> 8 & 0xff;

    //            var lowCharToByte = CharToByte[high];
    //            if (lowCharToByte == null)
    //            {
    //                lowCharToByte = new Byte[256];
    //                CharToByte[high] = lowCharToByte;
    //            }
    //            lowCharToByte[low] = (byte)(i);
    //        }
    //    }

    //    public static String GetString(byte[] bytes)
    //    {
    //        var sb = new System.Text.StringBuilder();
    //        foreach (var b in bytes)
    //        {
    //            sb.Append(ByteToChar[b]);
    //        }
    //        return sb.ToString();
    //    }

    //    public static Byte[] GetBytes(string str)
    //    {
    //        var mem = new System.IO.MemoryStream();
    //        foreach (var ch in str)
    //        {
    //            var high = ch >> 8 & 0xff;
    //            var lowCharToByte = CharToByte[high];
    //            Byte res = 0;
    //            if (lowCharToByte != null)
    //            {
    //                var low = ch & 0xff;
    //                res = lowCharToByte[low];
    //            }
    //            if (res == 0) res = 0xff;
    //            mem.WriteByte(res);
    //        }
    //        return mem.ToArray();
    //    }
    //}

    public class ConvertWindowsPersianToDOS
    {
        public Dictionary<byte, byte> CharachtersMapper_Group1 = new Dictionary<byte, byte>
        {

        {48 , 128}, // 0
        {49 , 129}, // 1
        {50 , 130}, // 2
        {51 , 131}, // 3
        {52 , 132}, // 4
        {53 , 133}, // 5
        {54 , 134}, // 6
        {55 , 135}, // 7
        {56 , 136}, // 8
        {57 , 137}, // 9
        {161, 138}, // ،
        {191, 140}, // ؟
        {193, 143}, //ء 
        {194, 141}, // آ
        {195, 144}, // أ
        {196, 248}, //ؤ  
        {197, 144}, //إ
        {200, 146}, //ب 
        {201, 249}, //ة
        {202, 150}, //ت
        {203, 152}, //ث 
        {204, 154}, //ﺝ
        {205, 158}, //ﺡ
        {206, 160}, //ﺥ
        {207, 162}, //د
        {208, 163}, //ذ
        {209, 164}, //ر
        {210, 165},//ز
        {211, 167},//س
        {212, 169},//ش
        {213, 171}, //ص
        {214, 173}, //ض
        {216, 175}, //ط
        {217, 224}, //ظ
        {218, 225}, //ع
        {219, 229}, //غ
        {220, 139}, //-
        {221, 233},//ف
        {222, 235},//ق
        {223, 237},//ك
        {225, 241},//ل
        {227, 244},//م
        {228, 246},//ن
        {229, 249},//ه
        {230, 248},//و
        {236, 253},//ى
        {237, 253},//ی
        {129, 148},//پ
        {141, 156},//چ
        {142, 166},//ژ
        {152, 237},//ک
        {144, 239},//گ


           };
        public Dictionary<byte, byte> CharachtersMapper_Group2 = new Dictionary<byte, byte>
        {
       {48,128},//
       {49,129},//
       {50,130},
       {51,131},//
       {52,132},//
       {53,133},
       {54,134},//
       {55,135},//
       {56,136},
       {57,137},//
       {161,138},//،
       {191,140},//?
       {193,143},//ء
       {194,141},//آ
       {195,144},//أ
       {196,248},//ؤ
       {197,144},//إ
       {198,254},//ئ
       {199,144},//ا
       {200,147},//ب
       {201,251},//ة
       {202,151},//ت
       {203,153},//ث
       {204,155},//ج
       {205,159},//ح
       {206,161},//خ
       {207,162},//د
       {208,163},//ذ
       {209,164},//ر
       {210,165},//ز
       {211,168},//س
       {212,170},//ش
       {213,172},//ص
       {214,174},//ض
       {216,175},//ط
       {217,224},//ظ
       {218,228},//ع
       {219,232},//غ
       {220,139},//-
       {221,234},//ف
       {222,236},//ق
       {223,238},//ك
       {225,243},//ل
       {227,245},//م
       {228,247},//ن
       {229,251},//ه
       {230,248},//و
       {236,254},//ی
       {237,254},//ی
       {129,149},//پ
       {141 ,157},//چ
       {142,166},//ژ
       {152,238},//ک
       {144,240},//گ
       
       
        };

        public Dictionary<byte, byte> CharachtersMapper_Group3 = new Dictionary<byte, byte>
        {

        {48 , 128}, // 0
        {49 , 129}, // 1
        {50 , 130}, // 2
        {51 , 131}, // 3
        {52 , 132}, // 4
        {53 , 133}, // 5
        {54 , 134}, // 6
        {55 , 135}, // 7
        {56 , 136}, // 8
        {57 , 137}, // 9
        {161, 138}, // ،
        {191, 140}, // ؟
        {193, 143}, //
        {194, 141}, //
        {195, 145}, //
        {196, 248}, //
        {197, 145}, // 
        {198, 252}, //
        {199, 145}, // 
        {200, 146}, // 
        {201, 249}, //
        {202, 150}, //
        {203, 152}, // 
        {204, 154}, //
        {205, 158}, // 
        {206, 160}, //
        {207, 162}, //
        {208, 163}, // 
        {209, 164}, //
        {210, 165}, //
        {211, 167}, // 
        {212, 169}, // 
        {213, 171}, //
        {214, 173}, // 
        {216, 175}, // 
        {217, 224}, //
        {218, 226}, // 
        {219, 230}, // 
        {220, 139}, //
        {221, 233}, // 
        {222, 235}, //
        {223, 237}, //
        {225, 241}, // 
        {227, 244}, //
        {228, 246}, //
        {229, 249}, //   
        {230, 248}, // 
        {236, 252}, //
        {237, 252}, // 
        {129, 148}, // 
        {141, 156}, //
        {142, 166}, // 
        {152, 237}, // 
        {144, 239}//
};
        public Dictionary<byte, byte> CharachtersMapper_Group4 = new Dictionary<byte, byte>
        {
            {48 , 128}, // 0
            {49 , 129}, // 1
            {50 , 130}, // 2
            {51 , 131}, // 3
            {52 , 132}, // 4
            {53 , 133}, // 5
            {54 , 134}, // 6
            {55 , 135}, // 7
            {56 , 136}, // 8
            {57 , 137}, // 9
            {161, 138}, // ،
            {191, 140}, // ؟
            {193,143}, //
            {194,141}, //
            {195,145}, //
            {196,248}, // 
            {197,145}, // 
            {198,254}, //
            {199,145}, // 
            {200,147}, // 
            {201,250}, //
            {202,151}, //
            {203,153}, //
            {204,155}, //
            {205,159}, //
            {206,161}, //
            {207,162}, //
            {208,163}, //
            {209,164}, //
            {210,165}, //
            {211,168}, // 
            {212,170}, //
            {213,172}, //
            {214,174}, //
            {216,175}, // 
            {217,224}, //
            {218,227}, //
            {219,231}, //
            {220,139}, //
            {221,234}, //
            {222,236}, //
            {223,238}, //
            {225,243}, //
            {227,245}, // 
            {228,247}, //
            {229,250}, //
            {230,248}, //
            {236,254}, //
            {237,254}, // 
            {129,149}, //
            {141,157}, //
            {142,166}, // 
            {152,238}, // 
            {144,240}, //
};


        public bool is_Lattin_Letter(byte c)
        {
            if (c < 128)// && c > 31)
            {
                return true;
            }
            return false;

        }
        public byte get_Lattin_Letter(byte c)
        {
            //if ("0123456789".IndexOf((char)c) >= 0)
            //{
            //    //return (byte)c;
            //    return (byte)(c + 80);
            //}
            //return get_FarsiExceptions(c);
            return c;
        }

        private byte get_FarsiExceptions(byte c)
        {
            switch (c)
            {
                case (byte)'(': return (byte)')';
                case (byte)'{': return (byte)'}';
                case (byte)'[': return (byte)']';
                case (byte)')': return (byte)'(';
                case (byte)'}': return (byte)'{';
                case (byte)']': return (byte)'[';
                default: return (byte)c;

            }

        }

        public bool is_Final_Letter(byte c)
        {
            string s = "ءآأؤإادذرزژو";

            if (s.ToString().IndexOf((char)c) >= 0)
            {
                return true;

            }
            return false;
        }
        public bool IS_White_Letter(byte c)
        {
            if (c == 8 || c == 09 || c == 10 || c == 13 || c == 27 || c == 32 || c == 0)
            {
                return true;
            }
            return false;
        }
        public bool Char_Cond(byte c)
        {
            return IS_White_Letter(c)
                || is_Lattin_Letter(c)
                || c == 191;
        }


        public List<byte> get_Unicode_To_IranSystem(string Unicode_Text)
        {

            // " رشته ای که فارسی است را دو کاراکتر فاصله به ابتدا و انتهایآن اضافه می کنیم
            string unicodeString = " " + Unicode_Text + " ";
            //ایجاد دو انکدینگ متفاوت
            Encoding ascii = //Encoding.ASCII;
                Encoding.GetEncoding("windows-1256");

            Encoding unicode = Encoding.Unicode;

            // تبدیل رشته به بایت
            byte[] unicodeBytes = unicode.GetBytes(unicodeString);

            // تبدیل بایتها از یک انکدینگ به دیگری
            byte[] asciiBytes = Encoding.Convert(unicode, ascii, unicodeBytes);

            // Convert the new byte[] into a char[] and then into a string.
            char[] asciiChars = new char[ascii.GetCharCount(asciiBytes, 0, asciiBytes.Length)];
            ascii.GetChars(asciiBytes, 0, asciiBytes.Length, asciiChars, 0);
            string asciiString = new string(asciiChars);
            byte[] b22 = Encoding.GetEncoding("windows-1256").GetBytes(asciiChars);


            int limit = b22.Length;

            byte pre = 0, cur = 0;


            List<byte> IS_Result = new List<byte>();
            for (int i = 0; i < limit; i++)
            {
                if (is_Lattin_Letter(b22[i]))
                {
                    cur = get_Lattin_Letter(b22[i]);

                    IS_Result.Add(cur);


                    pre = cur;
                }
                else if (i != 0 && i != b22.Length - 1)
                {
                    cur = get_Unicode_To_IranSystem_Char(b22[i - 1], b22[i], b22[i + 1]);

                    if (cur == 145) // برای بررسی استثنای لا
                    {
                        if (pre == 243)
                        {
                            IS_Result.RemoveAt(IS_Result.Count - 1);
                            IS_Result.Add(242);
                        }
                        else
                        {
                            IS_Result.Add(cur);
                        }
                    }
                    else
                    {
                        IS_Result.Add(cur);
                    }
                    pre = cur;
                }

            }

            IS_Result.RemoveAt(IS_Result.Count - 1); // Remove last space
            IS_Result.RemoveAt(0);// Remove first space

            if (!IsNumber(Unicode_Text) && !IsLatin(Unicode_Text))
            {
                IS_Result.Reverse();
            }

            return IS_Result;
        }

        static bool IsLatin(string str)
        {
            Encoding unicode = Encoding.Unicode;
            Encoding windowsArabic = Encoding.GetEncoding("windows-1256");
            byte[] arabicBytes = Encoding.Convert(unicode,windowsArabic , unicode.GetBytes(str));

            foreach (byte b in arabicBytes)
            {
                if (b >= 128)
                {
                    return false;
                }
            }

            return true;
        }

        static bool IsNumber(string str)
        {
            return Regex.IsMatch(str, @"^[\d]+$");
        }

        public byte get_Unicode_To_IranSystem_Char(byte PreviousChar, byte CurrentChar, byte NextChar)
        {

            bool PFlag = Char_Cond(PreviousChar) || is_Final_Letter(PreviousChar);
            bool NFlag = Char_Cond(NextChar);
            if (PFlag && NFlag) return UCTOIS_Group_1(CurrentChar);
            else if (PFlag) return UCTOIS_Group_2(CurrentChar);
            else if (NFlag) return UCTOIS_Group_3(CurrentChar);

            return UCTOIS_Group_4(CurrentChar);

        }

        private byte UCTOIS_Group_1(byte CurrentChar)
        {
            if (CharachtersMapper_Group1.ContainsKey(CurrentChar))
            {
                return (byte)CharachtersMapper_Group1[CurrentChar];
            }
            return (byte)CurrentChar;
        }
        private byte UCTOIS_Group_2(byte CurrentChar)
        {
            if (CharachtersMapper_Group2.ContainsKey(CurrentChar))
            {
                return (byte)CharachtersMapper_Group2[CurrentChar];
            }
            return (byte)CurrentChar;
        }
        private byte UCTOIS_Group_3(byte CurrentChar)
        {
            if (CharachtersMapper_Group3.ContainsKey(CurrentChar))
            {
                return (byte)CharachtersMapper_Group3[CurrentChar];
            }
            return (byte)CurrentChar;
        }
        private byte UCTOIS_Group_4(byte CurrentChar)
        {

            if (CharachtersMapper_Group4.ContainsKey(CurrentChar))
            {
                return (byte)CharachtersMapper_Group4[CurrentChar];
            }
            return (byte)CurrentChar;
        }
    }
}