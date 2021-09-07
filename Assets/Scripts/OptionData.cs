using System.Collections;
using System.Collections.Generic;
using UnityEngine;
using UnityEngine.UI;

public class OptionData
{
    public string dataName;
    public int columnIndex;
    public ValueType valueType;
    public CompareType compareType;
    public string value;

}

public enum CompareType
{
    LessOrEqual,
    Equal,
    NotEqual,
    GreaterOrEqual,
    None
}

public enum ValueType
{
    Float,
    Date,
    String
}