using System.Collections;
using System.Collections.Generic;
using UnityEngine;
using UnityEditor;
using System.IO;
using System.Text;
using System.Data;
using ExcelDataReader;

public class ExcelTools
{
    private const string INPUT_DIR = @"D:\Workspace\a0\document\excel";
    private const string OUTPUT_DIR = @"D:\Workspace\a0\project\Assets\Scripts\Excel";
    private static List<string> m_codeFileNames;

    [MenuItem("Tools/Excel To Code")]
    private static void Excel2Code()
    {
        DirectoryInfo _dirInfo = new DirectoryInfo(OUTPUT_DIR);
        // clear directory
        foreach (FileInfo _file in _dirInfo.GetFiles())
        {
            File.Delete(_file.FullName);
        }
        // read excel file
        _dirInfo = new DirectoryInfo(INPUT_DIR);
        m_codeFileNames = new List<string>();
        foreach (FileInfo _file in _dirInfo.GetFiles())
        {
            GenerateCSharpCode(_file);
        }
        GenerateManagerCode();
        EditorUtility.DisplayDialog(string.Empty, "Export Excel Finish", "OK");
        AssetDatabase.Refresh();
    }

    private static void GenerateManagerCode()
    {
        StringBuilder _codeBuilder = new StringBuilder();
        _codeBuilder.AppendLine("namespace ExcelData");
        _codeBuilder.AppendLine("{");
        _codeBuilder.AppendLine("public class ExcelManager");
        _codeBuilder.AppendLine("{");
        _codeBuilder.AppendLine("\tprivate static ExcelManager m_instance;");
        _codeBuilder.AppendLine("\tpublic static ExcelManager instance");
        _codeBuilder.AppendLine("\t{");
        _codeBuilder.AppendLine("\t\tget");
        _codeBuilder.AppendLine("\t\t{");
        _codeBuilder.AppendLine("\t\t\tif (m_instance == null)");
        _codeBuilder.AppendLine("\t\t\t{");
        _codeBuilder.AppendLine("\t\t\t\tm_instance = new ExcelManager();");
        _codeBuilder.AppendLine("\t\t\t}");
        _codeBuilder.AppendLine("\t\t\treturn m_instance;");
        _codeBuilder.AppendLine("\t\t}");
        _codeBuilder.AppendLine("\t}");
        foreach (string _className in m_codeFileNames)
        {
            _codeBuilder.AppendLine(string.Format("\tpublic {0} _{1};", _className, _className.ToLower()));
        }
        _codeBuilder.AppendLine("\tpublic ExcelManager()");
        _codeBuilder.AppendLine("\t{");
        foreach (string _className in m_codeFileNames)
        {
            _codeBuilder.AppendLine(string.Format("\t\t_{0} = new {1}();", _className.ToLower(), _className));
            _codeBuilder.AppendLine(string.Format("\t\t_{0}.Init();", _className.ToLower()));
        }
        _codeBuilder.AppendLine("\t}");
        _codeBuilder.AppendLine("}");
        _codeBuilder.AppendLine("}");
        FileStream _codeFile = new FileStream(Path.Combine(OUTPUT_DIR, "ExcelManager.cs"), FileMode.Create, FileAccess.Write);
        StreamWriter _sw = new StreamWriter(_codeFile);
        _sw.Write(_codeBuilder.ToString());
        _sw.Close();
        _codeFile.Close();
    }

    private static void GenerateCSharpCode(FileInfo fileInfo)
    {
        string _fileName = fileInfo.Name.Replace(".xlsx", "");
        m_codeFileNames.Add(_fileName);
        StringBuilder _codeBuilder = new StringBuilder();
        // write code
        _codeBuilder.AppendLine("using System.Collections;");
        _codeBuilder.AppendLine("using System.Collections.Generic;");
        _codeBuilder.AppendLine("using UnityEngine;");
        _codeBuilder.AppendLine();
        _codeBuilder.AppendLine("namespace ExcelData");
        _codeBuilder.AppendLine("{");
        _codeBuilder.AppendLine(string.Format("public class {0}", _fileName));
        _codeBuilder.AppendLine("{");

        FileStream _excelStream = File.Open(fileInfo.FullName, FileMode.Open, FileAccess.Read);
        IExcelDataReader _excelReader = ExcelReaderFactory.CreateOpenXmlReader(_excelStream);
        DataSet _result = _excelReader.AsDataSet();
        var _data = _result.Tables[0];

        int _rowCount = _data.Rows.Count;
        int _columnCount = _data.Columns.Count;
        // field name
        List<string> _fieldNames = new List<string>();
        for (int i = 0; i < _columnCount; i++)
        {
            _fieldNames.Add(_data.Rows[0][i].ToString());
        }
        // data type
        List<string> _dataTypes = new List<string>();
        for (int i = 0; i < _columnCount; i++)
        {
            _dataTypes.Add(_data.Rows[1][i].ToString());
        }
        // Entity
        _codeBuilder.AppendLine("\tpublic class Entity");
        _codeBuilder.AppendLine("\t{");
        for (int i = 0; i < _columnCount; i++)
        {
            _codeBuilder.AppendLine(string.Format("\t\tpublic {0} {1};", _dataTypes[i], _fieldNames[i]));
        }
        _codeBuilder.AppendLine("\t}");

        // class field
        _codeBuilder.AppendLine("\tpublic Dictionary<int, Entity> dataTable;");
        // init function
        _codeBuilder.AppendLine("\tpublic void Init()");
        _codeBuilder.AppendLine("\t{");
        _codeBuilder.AppendLine("\t\tdataTable = new Dictionary<int, Entity>();");
        // dataTable.Add(1, new Entity { id = 1, event_list = new List<int>() { 1, 2 }, next_group = 0 });

        for (int i = 3; i < _rowCount; i++)
        {
            _codeBuilder.Append("\t\tdataTable.Add(");
            for (int j = 0; j < _columnCount; j++)
            {
                string _value = _data.Rows[i][j].ToString();
                if (j == 0)
                {
                    _codeBuilder.AppendFormat("{0}, new Entity {{", _value);
                }

                string _codeStr = ParseCode(_dataTypes[j], _fieldNames[j], _value);
                if (j != _columnCount - 1)
                {
                    _codeBuilder.Append(string.Format("{0},", _codeStr));
                }
                else
                {
                    _codeBuilder.Append(_codeStr);
                }
            }
            _codeBuilder.AppendLine("});");
        }

        // append data
        _codeBuilder.AppendLine("\t}");

        // append function
        _codeBuilder.AppendLine("\tpublic Entity GetById(int id)");
        _codeBuilder.AppendLine("\t{");
        _codeBuilder.AppendLine("\t\tEntity _ret = null;");
        _codeBuilder.AppendLine("\t\tdataTable.TryGetValue(id, out _ret);");
        _codeBuilder.AppendLine("\t\treturn _ret;");
        _codeBuilder.AppendLine("\t}");
        _excelReader.Close();
        _codeBuilder.AppendLine("}");
        _codeBuilder.AppendLine("}");
        // save to file
        FileStream _codeFile = new FileStream(Path.Combine(OUTPUT_DIR, _fileName + ".cs"), FileMode.Create, FileAccess.Write);
        StreamWriter _sw = new StreamWriter(_codeFile);
        _sw.Write(_codeBuilder.ToString());
        _sw.Close();
        _codeFile.Close();
    }

    private static string ParseCode(string fieldType, string fieldName, string fieldValue)
    {
        string _ret = string.Empty;
        switch (fieldType)
        {
            case "int":
                if (string.IsNullOrEmpty(fieldValue))
                {
                    fieldValue = "0";
                }
                _ret = string.Format("{0} = {1}", fieldName, fieldValue);
                break;
            case "long":
                if (string.IsNullOrEmpty(fieldValue))
                {
                    fieldValue = "0";
                }
                _ret = string.Format("{0} = {1}", fieldName, fieldValue);
                break;
            case "float":
                if (string.IsNullOrEmpty(fieldValue))
                {
                    fieldValue = "0";
                }
                _ret = string.Format("{0} = {1}f", fieldName, fieldValue);
                break;
            case "string":
                _ret = string.Format("{0} = @\"{1}\"", fieldName, fieldValue);
                break;
            case "Vector3":
                _ret = ParseVector3Code(fieldName, fieldValue);
                break;
            default:
                if (fieldType.IndexOf("List<") >= 0)
                {
                    _ret = ParseListCode(fieldType, fieldName, fieldValue);
                }
                break;
        }
        return _ret;
    }

    private static string ParseVector3Code(string fieldName, string fieldValue)
    {
        string _ret = string.Empty;
        string[] _values = fieldValue.ToString().Split(new char[] { ',' });
        if (_values.Length > 2)
        {
            _ret = string.Format("{0} = new Vector3({1}f, {2}f, {3}f)", fieldName, _values[0], _values[1], _values[2]);
        }
        else
        {
            _ret = string.Format("{0} = Vector3.zero", fieldName);
        }
        return _ret;
    }

    private static string ParseListCode(string fieldType, string fieldName, string fieldValue)
    {
        string _ret = string.Empty;

        if (fieldType.IndexOf("<int>") >= 0) // int 
        {
            string[] _values = fieldValue.ToString().Split(new char[] { ',' });
            int _count = _values.Length;
            StringBuilder _sb = new StringBuilder();
            _sb.AppendFormat("{0} = new {1} {{", fieldName, fieldType);
            for (int i = 0; i < _count; i++)
            {
                if (i != _count - 1)
                {
                    _sb.AppendFormat("{0}, ", _values[i]);
                }
                else
                {
                    _sb.Append(_values[i]);
                }
            }
            _sb.Append("}");
            _ret = _sb.ToString();
        }
        else if (fieldType.IndexOf("<Vector3>") >= 0) // Vector3
        {
            if (fieldValue.Equals(string.Empty))
            {
                _ret = string.Format("{0} = null", fieldName);
            }
            else
            {
                string[] _values = fieldValue.ToString().Split(new char[] { '|' });
                int _count = _values.Length;
                StringBuilder _sb = new StringBuilder();
                _sb.AppendFormat("{0} = new {1}() {{ ", fieldName, fieldType);
                for (int i = 0; i < _count; i++)
                {
                    string[] _v3 = _values[i].Split(new char[] { ',' });
                    if (i != _count - 1)
                    {
                        _sb.AppendFormat("new Vector3({0}f, {1}f, {2}f), ", _v3[0], _v3[1], _v3[2]);
                    }
                    else
                    {
                        _sb.AppendFormat("new Vector3({0}f, {1}f, {2}f)", _v3[0], _v3[1], _v3[2]);
                    }
                }
                _sb.Append("}");
                _ret = _sb.ToString();
            }
        }


        return _ret;
    }
}
