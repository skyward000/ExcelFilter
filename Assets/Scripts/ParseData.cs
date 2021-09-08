using System.Collections;
using System.Collections.Generic;
using UnityEngine;
using UnityQuickSheet;
using UnityEngine.UI;
using NPOI.SS.UserModel;
using NPOI.HSSF.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.IO;
using NPOI.HSSF.Util;

public class ParseData : MonoBehaviour
{
    public DataGroupsCtrl dataGroupsCtrl;
    public Button parseData;
    public string inputPath;
    //public string outputPath;

    private ExcelQuery _excelQuery;
    private string[] _sheetNames;
    private List<string> _optionNames = new List<string>();
    private List<int> _optionNameIndexes = new List<int>();

    private void Awake()
    {
        parseData.onClick.AddListener(StartFiltering);
    }

    //private void Update()
    //{
    //    if (Input.GetKeyDown(KeyCode.Return))
    //    {
    //        LoadSheetName();
    //        LoadOption();
    //        dataGroupsCtrl.CreateGroups(_optionNames, _optionNameIndexes);


    //        //IWorkbook workbook = null;
    //        //FileStream fs = null;
    //        //IRow row = null;
    //        //ISheet sheet = null;
    //        //ICell cell = null;
    //        //try
    //        //{
    //        //    workbook = new XSSFWorkbook();
    //        //    sheet = workbook.CreateSheet("Sheet0");//����һ������ΪSheet0�ı�  
    //        //    int rowCount = 5;
    //        //    int columnCount = 5;

    //        //    //������ͷ  
    //        //    row = sheet.CreateRow(0);//excel��һ����Ϊ��ͷ  
    //        //    for (int c = 0; c < columnCount; c++)
    //        //    {
    //        //        cell = row.CreateCell(c);
    //        //        cell.SetCellValue(2);
    //        //    }

    //        //    //����ÿ��ÿ�еĵ�Ԫ��,  
    //        //    for (int i = 0; i < rowCount; i++)
    //        //    {
    //        //        row = sheet.CreateRow(i + 1);
    //        //        for (int j = 0; j < columnCount; j++)
    //        //        {
    //        //            cell = row.CreateCell(j);//excel�ڶ��п�ʼд������  
    //        //            cell.SetCellValue(3);
    //        //        }
    //        //    }
    //        //    using (fs = File.Create(@"C:\Users\Administrator\Desktop\output.xlsx"))
    //        //    {
    //        //        workbook.Write(fs);//��򿪵����xls�ļ���д������  
    //        //    }
    //        //}
    //        //catch (Exception ex)
    //        //{
    //        //    if (fs != null)
    //        //    {
    //        //        fs.Close();
    //        //    }
    //        //}
    //    }

    //}


    public void StartParseData()
    {
        LoadSheetName();
        LoadOption();
        dataGroupsCtrl.CreateGroups(_optionNames, _optionNameIndexes);
    }

    private void LoadSheetName()
    {
        _excelQuery = new ExcelQuery(inputPath);
        _sheetNames = _excelQuery.GetSheetNames();

        //foreach (var name in _sheetNames)
        //{
        //    Debug.Log(name);
        //}
    }

    private void LoadOption()
    {
        _optionNames.Clear();
        _optionNameIndexes.Clear();

        ISheet sheet = _excelQuery.Workbook.GetSheet(_sheetNames[0]);
        IRow row = sheet.GetRow(0);

        for (int i = 0; i < row.LastCellNum; i++)
        {
            if (row.GetCell(i) == null || row.GetCell(i).CellType == CellType.Blank)
            {
                continue;
            }

            _optionNameIndexes.Add(i);
            _optionNames.Add(row.GetCell(i).StringCellValue);
        }
    }

    public void StartFiltering()
    {
        EventCenter.Broadcast<CallBack<List<OptionData>>>(EventDefine.GetDataGroup, FilterData);
    }

    private void FilterData(List<OptionData> data)
    {
        List<ISheet> sheets = new List<ISheet>();
        List<IRow> resultRows = new List<IRow>();

        for (int i = 0; i < _sheetNames.Length; i++)
        {
            sheets.Add(_excelQuery.Workbook.GetSheet(_sheetNames[i]));
        }

        for (int i = 0; i < sheets.Count; i++)
        {
            foreach (IRow row in sheets[i])
            {
                if (row.RowNum == 0) continue;

                if (row.GetCell(0) != null)
                {
                    if (RowReachCondition(row, data))
                    {
                        resultRows.Add(row);
                    }
                }
            }
        }
        CreateSheet(resultRows, data);
    }

    private bool RowReachCondition(IRow row, List<OptionData> data)
    {
        for (int i = data[0].columnIndex, j = 0; i < data[data.Count - 1].columnIndex + 1; i++, j++)
        {
            if (!CellReachCondition(row.GetCell(i), data[j]))
            {
                return false;
            }
        }
        return true;
    }

    private bool CellReachCondition(ICell cell, OptionData data)
    {
        if (cell == null)
            return true;

        //Debug.Log(cell.CellType);
        switch (data.valueType)
        {
            case ValueType.Float:
                switch (data.compareType)
                {
                    case CompareType.LessOrEqual:
                        return cell.NumericCellValue <= System.Convert.ToDouble(data.value);
                    case CompareType.Equal:
                        return cell.NumericCellValue == System.Convert.ToDouble(data.value);
                    case CompareType.NotEqual:
                        return cell.NumericCellValue != System.Convert.ToDouble(data.value);
                    case CompareType.GreaterOrEqual:
                        return cell.NumericCellValue >= System.Convert.ToDouble(data.value);
                    case CompareType.None:
                        return true;
                    default:
                        throw new Exception("�Ƚ����Ͳ�Ӧ��Ϊ��" + data.compareType);
                }
            case ValueType.Date:
                int compareResult = 0;
                if (data.compareType != CompareType.None)
                    compareResult = DateTime.Compare(cell.DateCellValue, System.Convert.ToDateTime(data.value));

                switch (data.compareType)
                {
                    case CompareType.LessOrEqual:
                        return compareResult < 0 || compareResult == 0;
                    case CompareType.Equal:
                        return compareResult == 0;
                    case CompareType.NotEqual:
                        return compareResult != 0;
                    case CompareType.GreaterOrEqual:
                        return compareResult > 0 || compareResult == 0;
                    case CompareType.None:
                        return true;
                    default:
                        throw new Exception("�Ƚ����Ͳ�Ӧ��Ϊ��" + data.compareType);
                }
            case ValueType.String:
                switch (data.compareType)
                {
                    case CompareType.Equal:
                        return cell.StringCellValue.Equals(data.value);
                    case CompareType.NotEqual:
                        return !cell.StringCellValue.Equals(data.value);
                    case CompareType.Include:
                        return cell.StringCellValue.Contains(data.value);
                    case CompareType.None:
                        return true;
                    default:
                        throw new Exception("�Ƚ����Ͳ�Ӧ��Ϊ��" + data.compareType);
                }

            default:
                throw new Exception("ֵ���Ͳ�Ӧ��Ϊ��" + data.compareType);
        }
    }

    private void CreateSheet(List<IRow> results, List<OptionData> optionData)
    {
        //for (int i = 0; i < results.Count; i++)
        //{
        //    foreach (ICell cell in results[i])
        //    {
        //        //Debug.Log(string.Format("�����{0}��ֵ{1}", i, cell.StringCellValue));
        //    }
        //}
        IWorkbook workbook;
        if (_excelQuery.Workbook is HSSFWorkbook)
            workbook = new HSSFWorkbook();
        else if (_excelQuery.Workbook is XSSFWorkbook)
            workbook = new XSSFWorkbook();
        else
            throw new Exception("Wrong file.");


        //ISheet sheet = workbook.CreateSheet("Sheet0");
        ISheet sheet = workbook.CreateSheet("Sheet0");
        int[] columnWidth = new int[12] { 21, 6, 35, 12, 14, 11, 12, 12, 12, 12, 12, 12 };
        for (int i = 0; i < columnWidth.Length; i++)
        {
            sheet.SetColumnWidth(i, 256 * columnWidth[i]);
        }

        IRow rowToCreate;
        ICell cellToCreate;
        IDataFormat dataFormatCustom = workbook.CreateDataFormat();
        ICellStyle dateStyle = workbook.CreateCellStyle();
        dateStyle.Alignment = HorizontalAlignment.Left;                             //����ˮƽ���뷽ʽ
        dateStyle.VerticalAlignment = VerticalAlignment.Center;
        dateStyle.DataFormat = dataFormatCustom.GetFormat("yyyy-MM-dd");

        //ICellStyle style1 = _excelQuery.Workbook.CreateCellStyle();//��ʽ
        //style1.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Left;//����ˮƽ���뷽ʽ
        //style1.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center;//���ִ�ֱ���뷽ʽ
        //                                                                      //���ñ߿�
        //style1.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;
        //style1.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin;
        //style1.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;
        //style1.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;
        //style1.WrapText = true;//�Զ�����

        //ICellStyle style2 = _excelQuery.Workbook.CreateCellStyle();//��ʽ
        //IFont font1 = _excelQuery.Workbook.CreateFont();//����
        //font1.FontName = "����";
        //font1.Color = HSSFColor.Red.Index;//������ɫ
        //font1.Boldweight = (short)FontBoldWeight.Normal;//����Ӵ���ʽ
        //style2.SetFont(font1);//��ʽ����������þ����������ʽ
        //                      //���ñ���ɫ
        //style2.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.Yellow.Index;
        //style2.FillPattern = FillPattern.SolidForeground;
        //style2.FillBackgroundColor = NPOI.HSSF.Util.HSSFColor.Yellow.Index;
        //style2.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Left;//����ˮƽ���뷽ʽ
        //style2.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center;//���ִ�ֱ���뷽ʽ


        //��������
        rowToCreate = sheet.CreateRow(0);
        rowToCreate.Height = 30 * 20;
        for (int i = 0; i < optionData[optionData.Count - 1].columnIndex + 1; i++)
        {
            cellToCreate = rowToCreate.CreateCell(i);
            if (i < optionData[0].columnIndex)
                continue;

            cellToCreate.SetCellValue(optionData[i - optionData[0].columnIndex].dataName);
        }



        for (int i = 0; i < results.Count; i++)
        {
            rowToCreate = sheet.CreateRow(i + 1);
            rowToCreate.Height = 30 * 20;

            //����������
            for (int j = 0; j < results[i].LastCellNum; j++)
            {
                cellToCreate = rowToCreate.CreateCell(j);
                if (results[i].GetCell(j) == null || results[i].GetCell(j).CellType == CellType.Blank)
                    continue;

                if (j < optionData[0].columnIndex || j > optionData[optionData.Count - 1].columnIndex)
                {
                    switch (results[i].GetCell(j).CellType)
                    {
                        case CellType.Numeric:
                            cellToCreate.SetCellValue(results[i].GetCell(j).NumericCellValue);
                            break;
                        case CellType.String:
                            cellToCreate.SetCellValue(results[i].GetCell(j).StringCellValue);
                            break;
                        default:
                            throw new Exception("��֧��" + results[i].GetCell(j).CellType + "����");
                    }
                }
                else
                {
                    switch (optionData[j - optionData[0].columnIndex].valueType)
                    {
                        case ValueType.Float:
                            cellToCreate.SetCellValue(results[i].GetCell(j).NumericCellValue);
                            break;
                        case ValueType.Date:
                            cellToCreate.SetCellValue(results[i].GetCell(j).NumericCellValue);
                            cellToCreate.CellStyle = dateStyle;
                            break;
                        case ValueType.String:
                            cellToCreate.SetCellValue(results[i].GetCell(j).StringCellValue);
                            break;
                        default:
                            throw new Exception("��֧�ֵ�����");
                    }
                }
                //cellToCreate.SetCellValue(results[i].GetCell(j).ty)
            }
        }

        try
        {
            //    using (fs = File.Create(@"C:\Users\Administrator\Desktop\output.xlsx"))
            //    {
            //        workbook.Write(fs);//��򿪵����xls�ļ���д������  
            //    }
            FileStream fs;
            if (workbook is HSSFWorkbook)
                fs = File.Create(@"D:\output.xls");
            else if (workbook is XSSFWorkbook)
                fs = File.Create(@"D:\output.xlsx");
            else
                return;

            workbook.Write(fs);
            fs.Close();
        }
        catch (Exception e)
        {
            throw new Exception(e.Message);
        }
    }
}


