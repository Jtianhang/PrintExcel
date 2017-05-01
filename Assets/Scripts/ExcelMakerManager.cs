using System;
using System.Collections;
using System.Collections.Generic;
using UnityEngine;
using Excel;
using System.Data;
using System.IO;
using OfficeOpenXml;

public class TestInfo
{
    public string name;
    public string id;
    public string num;
};

public enum DayState
{
    Morning,
    Noon,
    Night,
}


[System.Serializable]
public class EmployeeData
{
    public static readonly int DayCount = 31;

    public int ID;
    public string Name;
    public DayState State;
    public string[] SignDate;

    public EmployeeData()
    {
        SignDate = new string[DayCount];
    }
}




public class FishData
{
    public int id;
    public string fname;
    public int odds;
    public int typeIndex;
    public string hitTypeS;

    public float rotateAngleRnd;
    public float rotateInterval;
    public float rotateIntervalRnd;

    public float roundCircleRadius;
    public float speed;
    public float rotateSpeed;

}

public class ExcelMakerManager
{
    static ExcelMakerManager _instance = null;
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
    public static string ExcelName = "FishData.xlsx";

    //链表为物体信息
    public void ExcelMaker(string name)
    {
        string path = FilePath(name);

        WriteExcel(path);
        return;
        FileStream stream = File.Open(path, FileMode.Open, FileAccess.Read, FileShare.Read);
        IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(stream);
        DataSet result = excelReader.AsDataSet();
        DataRowCollection collect = result.Tables[0].Rows;

        for (int i = 1; i < collect.Count; i++)
        {
            DataRow row = collect[i];
            if (row[1].ToString() == "")
            {
                continue;
            }
            for (int iter = 0; iter < row.ItemArray.Length; iter++)
            {
                Debug.LogError(row.ItemArray[iter].ToString());
            }
        }
        stream.Close();
        stream.Dispose();
        stream = null;
        excelReader.Close();
        excelReader.Dispose();
        excelReader = null;
    }

    public static List<FishData> SelectFishTable()
    {
        FileStream stream = File.Open(FilePath(ExcelName), FileMode.Open, FileAccess.Read, FileShare.Read);
        IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(stream);
        DataSet result = excelReader.AsDataSet();
        DataRowCollection collect = result.Tables[0].Rows;

        List<FishData> fishArray = new List<FishData>();

        for (int i = 1; i < collect.Count; i++)
        {
            if (collect[i][1].ToString() == "") continue;

            FishData fish = new FishData();
            fish.id = int.Parse(collect[i][0].ToString());
            fish.fname = collect[i][1].ToString();
            fish.odds = int.Parse(collect[i][2].ToString());
            fish.typeIndex = int.Parse(collect[i][3].ToString());
            fish.hitTypeS = collect[i][4].ToString();

            fish.rotateAngleRnd = float.Parse(collect[i][5].ToString());
            fish.rotateInterval = float.Parse(collect[i][6].ToString());
            fish.rotateIntervalRnd = float.Parse(collect[i][7].ToString());

            fish.roundCircleRadius = float.Parse(collect[i][8].ToString());
            fish.speed = float.Parse(collect[i][9].ToString());
            fish.rotateSpeed = float.Parse(collect[i][10].ToString());
            fishArray.Add(fish);
        }
        return fishArray;
    }

    /// <summary>
    /// 读取 Excel 需要添加 Excel; System.Data;
    /// </summary>
    /// <param name="sheet"></param>
    /// <returns></returns>
    static DataRowCollection ReadExcel(string sheet)
    {
        FileStream stream = File.Open(FilePath(ExcelName), FileMode.Open, FileAccess.Read, FileShare.Read);
        IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(stream);

        DataSet result = excelReader.AsDataSet();
        //int columns = result.Tables[0].Columns.Count;
        //int rows = result.Tables[0].Rows.Count;
        return result.Tables[sheet].Rows;
    }

    public static string FilePath(string name)
    {
        if (!Directory.Exists(Application.dataPath + "/Prints"))
        {
            Directory.CreateDirectory(Application.dataPath + "/Prints");
        }


        string path = Application.dataPath + "/Prints/" + name + ".xlsx";
        return path;
    }
    public static void WriteExcel(string outputDir)
    {
        FileInfo newFile = new FileInfo(outputDir);
        using (ExcelPackage package = new ExcelPackage(newFile))
        {
            if (package.Workbook.Worksheets.Count == 0)
            {
                package.Workbook.Worksheets.Add("Sheet1");
            }
            ExcelWorksheet worksheet = package.Workbook.Worksheets["Sheet1"];
            //Add the headers
            worksheet.Cells[1, 1].Value = "ID";
            worksheet.Cells[1, 2].Value = "Product";
            worksheet.Cells[1, 3].Value = "Quantity";
            worksheet.Cells[1, 4].Value = "Price";
            worksheet.Cells[1, 5].Value = "Value";
            worksheet.Cells[1, 6].Value = "th";
            //Add some items...
            worksheet.Cells["A2"].Value = 12001;
            worksheet.Cells["B2"].Value = "Nails";
            worksheet.Cells["C2"].Value = 37;
            worksheet.Cells["D2"].Value = 3.99;

            worksheet.Cells["A3"].Value = 12002;
            worksheet.Cells["B3"].Value = "Hammer";
            worksheet.Cells["C3"].Value = 5;
            worksheet.Cells["D3"].Value = 12.10;

            worksheet.Cells["A4"].Value = 12003;
            worksheet.Cells["B4"].Value = "Saw";
            worksheet.Cells["C4"].Value = 12;
            worksheet.Cells["D4"].Value = 15.37;

            //save our new workbook and we are done!
            for (int iter = 0; iter < EmployeeData.DayCount; iter++)
            {
                worksheet.Cells["A2"].Value = iter;
                worksheet.Cells["B2"].Value = "Nails";
                worksheet.Cells["C2"].Value = 37;
                worksheet.Cells["D2"].Value = 3.99;
            }



            package.Save();
        }
    }



}