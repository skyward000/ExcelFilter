using System.Collections;
using System.Collections.Generic;
using UnityEngine;
using UnityQuickSheet;
using UnityEngine.UI;
using NPOI.SS.UserModel;

public class ParseData : MonoBehaviour
{
    public string inputPath;
    public string outputPath;

    private ExcelQuery _excelQuery;
    private string[] _sheetNames;
    private string[] _optionNames;
    private int[] _optionNameIndexes;

    private void Update()
    {
        if (Input.GetKeyDown(KeyCode.Return))
        {
            LoadSheetName();
            LoadOption();
        }
    }

    public void StartParseData()
    {
        
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
        ISheet sheet = _excelQuery.Workbook.GetSheet(_sheetNames[0]);
        IRow row = sheet.GetRow(0);

        _optionNames = new string[row.Cells.Count];
        _optionNameIndexes = new int[row.Cells.Count];
        for (int i = 0; i < row.Cells.Count; i++)
        {
            _optionNameIndexes[i] = row.FirstCellNum + i;
            _optionNames[i] = row.Cells[i].StringCellValue;
        }

        foreach (var item in _optionNameIndexes)
            Debug.Log(item);
        foreach (var item in _optionNames)
            Debug.Log(item);


    }

}
