using UnityEngine;
using System.Collections;
using System.Runtime.InteropServices;
using UnityEngine.UI;
using UnityEditor;
using System.IO;

public class UIOpenDialog : MonoBehaviour
{
    public enum InOutMode { Input, Output }

    public InOutMode mode;
    public Text pathText;
    public ParseData parseData;
    private Button _btn;

    private void Awake()
    {
        _btn = GetComponent<Button>();
        _btn.onClick.AddListener(OpenDialog);
    }

    private void OpenDialog()
    {
        OpenFileName openFileName = new OpenFileName();
        openFileName.structSize = Marshal.SizeOf(openFileName);
        //openFileName.filter = "Excel文件(*.xlsx)\0*.xlsx";
        openFileName.file = new string(new char[256]);
        openFileName.maxFile = openFileName.file.Length;
        openFileName.fileTitle = new string(new char[64]);
        openFileName.maxFileTitle = openFileName.fileTitle.Length;
        openFileName.initialDir = Application.streamingAssetsPath;//.Replace('/', '\\');//默认路径
        openFileName.title = "窗口标题";
        openFileName.flags = 0x00080000 | 0x00001000 | 0x00000800 | 0x00000008;

        if (LocalDialog.GetSaveFileName(openFileName))
        {
            //Debug.Log(openFileName.file);
            //Debug.Log(openFileName.fileTitle);
            pathText.text = openFileName.file;
        }

        if (mode == InOutMode.Input)
        {
            parseData.inputPath = openFileName.file;
        }
    }
    //void OnGUI()
    //{
    //    if (GUI.Button(new Rect(10, 10, 100, 50), "Open"))
    //    {

    //    }
    //    //引用Editor的代码如果在运行时仍要调用，打包时会报错
    //    //if (GUI.Button(new Rect(10, 10, 100, 50), "Open"))
    //    //{
    //    //    string path = Application.dataPath;
    //    //    string folder = Path.GetDirectoryName(path);
    //    //    path = EditorUtility.OpenFilePanel("Open Excel file", folder, "excel files;*.xls;*.xlsx");
    //    //    Debug.Log(path);
    //    //}
    //}
}