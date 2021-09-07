using System.Collections;
using System.Collections.Generic;
using UnityEngine;
using UnityEngine.UI;

public class UIDataGroup : MonoBehaviour
{
    public Text dataName;
    public Dropdown dropdown_ValueType;
    public Dropdown dropdown_CompareType;
    public InputField inputField;

    private void Awake()
    {
        dataName = transform.Find("ColumnName").GetComponent<Text>();
        dropdown_ValueType = transform.Find("ValueType").GetComponent<Dropdown>();
        dropdown_CompareType = transform.Find("CompareType").GetComponent<Dropdown>();
        inputField = transform.Find("InputField").GetComponent<InputField>();
    }

    public void Init(string name)
    {
        dataName.text = name;
    }

    public OptionData GetData()
    {
        OptionData data = new OptionData();
        data.dataName = dataName.text;
        data.valueType = (ValueType)dropdown_ValueType.value;
        data.compareType = (CompareType)dropdown_CompareType.value;
        data.value = inputField.text;
        return data;
    }
}
