using org.in2bits.MyXls;
using System.Collections.Generic;
public class TestInfo
{
    public string name;
    public string id;
    public string num;
};
public class ExcelMakerManager
{
    public static ExcelMakerManager eInstance;
    public static ExcelMakerManager CreateExcelMakerManager()
    {
        if (eInstance == null)
        {
            eInstance = new ExcelMakerManager();
        }
        return eInstance;
    }
    //链表为物体信息
    public void ExcelMaker(string name, List<TestInfo> listInfo)
    {
        XlsDocument xls = new XlsDocument();
        //新建一个xls文档 
        xls.FileName = name;// @"D:\tests.xls";//设定文件名  //Add some metadata (visible from Excel under File -> Properties)
        xls.SummaryInformation.Author = "xyy"; //填加xls文件作者信息
        xls.SummaryInformation.Subject = "test";//填加文件主题信息  
        string sheetName = "Sheet0";
        Worksheet sheet = xls.Workbook.Worksheets.Add(sheetName);//填加名为"chc 实例"的sheet页 
        Cells cells = sheet.Cells;//Cells实例是sheet页中单元格（cell）集合  
        int rowNum = listInfo.Count;
        int rowMin = 1;
        int row = 0;
        for (int x = 0; x < rowNum + 1; x++)
        {
            if (x == 0)
            {
                //根据具体的物体信息 .需要重新写 
                cells.Add(1, 1, "名字");
                cells.Add(1, 2, "ID");
                cells.Add(1, 3, "数量");
            }
            else
            {
                cells.Add(rowMin + x, 1, listInfo[row].id);
                cells.Add(rowMin + x, 2, listInfo[row].name);
                cells.Add(rowMin + x, 3, listInfo[row].num); row++;
            }
        }
        xls.Save();
    }
}