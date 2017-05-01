using UnityEngine;
using System.Collections.Generic;
using UnityEngine.UI;

public class PrintExcel : MonoBehaviour
{
    public Text FileNameText;

    public void OnPrintExcel()
    {
        ExcelMakerManager.Instance.ExcelMaker(FileNameText.text);
    }
}