using UnityEngine;
using System.Collections;
using System.IO;
using org.in2bits.MyXls;
using System;
using System.Collections.Generic;
public class test : MonoBehaviour
{
    TestInfo test1;
    TestInfo test2;
    TestInfo test3;
    List<TestInfo> listInfos;
    // Use this for initialization 
    void Start()
    {
        // --测试数据  
        test1 = new TestInfo();
        test1.id = "one";
        test1.name = "test1";
        test1.num = "x";
        test2 = new TestInfo();
        test2.id = "two";
        test2.name = "test2";
        test2.num = "22";
        test3 = new TestInfo();
        test3.id = "tree";
        test3.name = "test3";
        test3.num = "333";
        listInfos = new List<TestInfo>();
        listInfos.Add(test1);
        listInfos.Add(test2);
        listInfos.Add(test3);
        // --测试数据 
    }

    void OnGUI()
    {
        if (GUI.Button(new Rect(100, 0, 100, 100), "aa"))
        {
            PrintExcel();
            Debug.Log("aaaa");
        }
    }
    public void PrintExcel()
    {
        if (!Directory.Exists(Application.dataPath + "/Prints"))
        {
            Directory.CreateDirectory(Application.dataPath + "/Prints");
        }
        string path = Application.dataPath + "/Prints/qiandao" + ".xls";
        ExcelMakerManager.Instance.ExcelMaker(path, listInfos);
    }
}