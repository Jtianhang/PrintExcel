using System.Collections;
using System.Collections.Generic;
using UnityEngine;

public class UIBaseWindow : MonoBehaviour
{
    public GameObject[] OpenPanel;
    public GameObject[] ClosePanel;

    public void OnCreatButtonMsg()
    {
        for(int iter = 0;iter< ClosePanel.Length;iter++)
        {
            ClosePanel[iter].gameObject.SetActive(false);
        }
        for (int iter = 0; iter < OpenPanel.Length; iter++)
        {
            OpenPanel[iter].gameObject.SetActive(true);
        }
    }
}