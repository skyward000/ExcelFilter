using System.Collections;
using System.Collections.Generic;
using UnityEngine;
using UnityEngine.UI;

public class OptionData : MonoBehaviour
{
    public string optionName;
    public OptionType optionType;
    public string inputValue;

    private void Awake()
    {
        //GetComponent<Dropdown>().itemText
    }

}

public enum OptionType
{ 
    LessOrEqual,
    Equal,
    GreaterOrEqual,
    None
}
