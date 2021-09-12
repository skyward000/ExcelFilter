using System.Collections;
using System.Collections.Generic;
using UnityEngine;
using UnityEngine.UI;

public class DataGroupsCtrl : MonoBehaviour
{
    public GameObject dataGroupPf;

    private List<GameObject> _groups = new List<GameObject>();
    private List<int> _indexes;

    private void OnEnable()
    {
        EventCenter.AddListener<CallBack<List<OptionData>>>(EventDefine.GetDataGroup, GetData);
    }

    private void OnDisable()
    {
        EventCenter.RemoveListener<CallBack<List<OptionData>>>(EventDefine.GetDataGroup, GetData);
    }

    public void CreateGroups(List<string> names, List<int> indexes)
    {
        for (int i = 0; i < _groups.Count; i++)
        {
            Destroy(_groups[i]);
        }

        _indexes = indexes;
        for (int i = 0; i < names.Count; i++)
        {
            GameObject dataObj = Instantiate(dataGroupPf, transform);
            dataObj.GetComponent<RectTransform>().anchoredPosition = new Vector2(50, -60 * i);
            dataObj.GetComponent<UIDataGroup>().Init(names[i]);
            _groups.Add(dataObj);
        }
    }

    public void GetData(CallBack<List<OptionData>> callBack)
    {
        List<OptionData> data = new List<OptionData>();
        for (int i = 0; i < _groups.Count; i++)
        {
            data.Add(_groups[i].GetComponent<UIDataGroup>().GetData());
            data[data.Count - 1].columnIndex = _indexes[i];
        }
        callBack(data);
    }
}
