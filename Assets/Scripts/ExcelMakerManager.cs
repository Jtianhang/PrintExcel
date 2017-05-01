using System;
using System.Collections;
using System.Collections.Generic;
using UnityEngine;
using org.in2bits.MyXls;
using System.IO;

[System.Serializable]
public class SignDateData
{
    public static readonly int SignDateCount = 3;
    public string[] SignDate;

    public SignDateData()
    {
        SignDate = new string[SignDateCount];
    }
}
[System.Serializable]
public class EmployeeData
{
    public static readonly int DayCount = 31;
    public int ID;
    public string Name;
    public SignDateData[] SignDateDatas;

    public EmployeeData()
    {
        SignDateDatas = new SignDateData[DayCount];
        for (int iter = 0; iter < DayCount; iter++)
        {
            SignDateDatas[iter] = new SignDateData();
        }
    }
}
[System.Serializable]
public class AllEmployee
{
    public List<EmployeeData> Data = new List<EmployeeData>();
}

public class ExcelMakerManager
{
    public static ExcelMakerManager _instance;
    public static ExcelMakerManager Instance
    {
        get
        {
            if (_instance == null)
            {
                _instance = new ExcelMakerManager();
            }
            return _instance;
        }
    }
    //链表为物体信息
    public void ExcelMaker(string fileName)
    {
        string path = GetFilePath(fileName);
        FileInfo file = new FileInfo(path);
        if (!file.Directory.Exists)
        {
            file.Directory.Create();
        }

        EmployeeData en = new EmployeeData();
        en.ID = 0;
        en.Name = "11";
        for (int iter = 0; iter < EmployeeData.DayCount; iter++)
        {
            for (int jter = 0; jter < SignDateData.SignDateCount; jter++)
            {
                en.SignDateDatas[iter].SignDate[jter] = (iter * jter).ToString();
            }
        }
        AllEmployee obj = new AllEmployee();
        obj.Data.Add(en);
        obj.Data.Add(en);
        obj.Data.Add(en);
        string str = JsonUtility.ToJson(obj, true);
        File.WriteAllText(path + ".txt", str);
        PrintExcel(fileName);
    }

    public void PrintExcel(string fileName)
    {
        string path = GetFilePath(fileName);
        FileInfo fileinfo = new FileInfo(path + ".txt");
        if (!fileinfo.Exists)
        {
            return;
        }
        string str = File.ReadAllText(path + ".txt");
        AllEmployee data = JsonUtility.FromJson<AllEmployee>(str);
        XlsDocument xls = new XlsDocument();
        //新建一个xls文档 
        xls.FileName = path;// @"D:\tests.xls";//设定文件名  //Add some metadata (visible from Excel under File -> Properties)
        xls.SummaryInformation.Author = "xyy"; //填加xls文件作者信息
        xls.SummaryInformation.Subject = "test";//填加文件主题信息  
        string sheetName = "Sheet0";
        Worksheet sheet = xls.Workbook.Worksheets.Add(sheetName);//填加名为"chc 实例"的sheet页 
        Cells cells = sheet.Cells;//Cells实例是sheet页中单元格（cell）集合
        //cells
        //表头
        XF xf = xls.NewXF();
        xf.TextDirection = TextDirections.LeftToRight;
        cells.Add(1, 1, "姓名//日期");
        for (int iter = 2; iter < EmployeeData.DayCount + 2; iter++)
        {
            cells.Add(1, iter, (iter - 1).ToString());
        }
        //内容
        for (int iter = 0; iter < data.Data.Count; iter++)
        {
            WriteExcel(cells, data.Data[iter], iter * SignDateData.SignDateCount);
        }
        xls.Save();
    }
    void WriteExcel(Cells cells, EmployeeData en, int offset)
    {
        cells.Add(2 + offset, 1, en.Name);
        for (int iter = 2; iter < SignDateData.SignDateCount + 2; iter++)
        {
            for (int jter = 2; jter < EmployeeData.DayCount + 2; jter++)
            {
                cells.Add(iter + offset, jter, en.SignDateDatas[jter - 2].SignDate[iter - 2]);
            }
        }
    }
    string GetFilePath(string fileName)
    {
        string path;
#if UNITY_EDITOR
        path = Application.dataPath + "/Print" + "/" + fileName;
#elif UNITY_ANDROID
        path = Application.persistentDataPath + "/Print" + "/" + fileName;
#endif
        return path;
    }


}