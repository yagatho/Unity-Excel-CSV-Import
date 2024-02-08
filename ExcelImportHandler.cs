using System.Collections.Generic;
using System.Text.RegularExpressions;
using UnityEditor;
using UnityEngine;


public class ExcelImportHandler : MonoBehaviour
{
    #region Variables
    [Header("Excel Column Names")]
    [SerializeField] private List<ExcelColumn> excelColumns = new List<ExcelColumn>();

    [Header("Options")]
    [SerializeField] private string csvFileName;
    [SerializeField] private List<GameObject> prefabs = new List<GameObject>();

    [Space(20f)]

    [Header("Debug")]
    [SerializeField] private List<Item> itemDatabase = new List<Item>();




    private Item blankItem;
    private List<GameObject> itemList;

    #endregion

    #region Functions

    public void SpawnObjects()
    {
        //destroy old items
        foreach (var item in itemList)
        {
            DestroyImmediate(item);
        }

        //clear item list
        itemList.Clear();



        LoadItemData();

        //Spawn Objects
        foreach (var item in itemDatabase)
        {
            //create temp pf bar
            GameObject pfObj = null;

            //chack for prefab
            foreach (var prefab in prefabs)
            {
                if(item.prefabName == prefab.name)
                    pfObj = prefab;
            }

            //instantiate prefab
            GameObject go = Instantiate(pfObj);

            //get transform
            Transform t = go.GetComponent<Transform>();

            //set parent and postion
            t.SetPositionAndRotation(item.position, Quaternion.Euler(item.rotation));
            t.parent = gameObject.transform;

            itemList.Add(go);
        }
    }

    private void LoadItemData()
    {
        //Clear db
        itemDatabase.Clear();

      

        //Read CSV
        List<Dictionary<string, object>> data = CSVReader.Read(csvFileName);
        for (int i = 0; i < data.Count; i++)
        {
            //Create temporary item
            Item tempItem = new Item(blankItem);

            //Check for each column
            foreach (var item in excelColumns)
            {
                //Check 
                switch (item.itemVariable)
                {
                    case itemVariable.prefabName:
                        tempItem.prefabName = data[i][item.columnName].ToString();

                        break;
                    case itemVariable.objectName:
                        tempItem.objectName = data[i][item.columnName].ToString();

                        break;
                    case itemVariable.posX:
                        tempItem.position.x = float.Parse(data[i][item.columnName].ToString(), System.Globalization.NumberStyles.Float);

                        break;
                    case itemVariable.posZ:
                        tempItem.position.z = float.Parse(data[i][item.columnName].ToString(), System.Globalization.NumberStyles.Float);

                        break;
                    case itemVariable.posY:
                        tempItem.position.y = float.Parse(data[i][item.columnName].ToString(), System.Globalization.NumberStyles.Float);

                        break;
                    case itemVariable.rotY:
                        tempItem.rotation.y = float.Parse(data[i][item.columnName].ToString(), System.Globalization.NumberStyles.Float);

                        break;
                }
            }

            itemDatabase.Add(tempItem);
        }
    }

    #endregion
}

[System.Serializable]
public class Item { 
    public string prefabName; //name of the prefab in the prefab folder
    public string objectName; //name of the object in scene
    public Vector3 position = Vector3.zero; //position in world scale
    public Vector3 rotation = Vector3.zero; //rotation in world scale

    public Item(Item i)
    {
        prefabName = i.prefabName;
        objectName = i.objectName;  
        position = i.position;
        rotation = i.rotation;
    }
}

[System.Serializable]
public class ExcelColumn
{
    public itemVariable itemVariable;
    public string columnName;

}

[System.Serializable]
public enum itemVariable
{
    prefabName,
    objectName,
    posX,
    posY,
    posZ,
    rotY,

}

public class CustomEditorWindow : EditorWindow
{
    [MenuItem("Tools/Get Objects From Excel Spreadsheet")]
    public static void ShowWindows()
    {
        GetWindow<CustomEditorWindow>("SpawnObjects");
    }

    private void OnGUI()
    {
        GUILayout.Label("Get Objects From Excel Spreadsheet", EditorStyles.boldLabel);

        GUILayout.TextField("Make sure to attach Excel Import Handler script to some empty object in scene!");

        if (GUILayout.Button("Spawn Objects"))
        {
            if (FindObjectOfType<ExcelImportHandler>() != null)
                FindObjectOfType<ExcelImportHandler>().SpawnObjects();
            else
                Debug.LogError("Attach Excel Import Handler to some empty object in scene!");
        }
    }
}

public class CSVReader
{
    static string SPLIT_RE = @",(?=(?:[^""]*""[^""]*"")*(?![^""]*""))";
    static string LINE_SPLIT_RE = @"\r\n|\n\r|\n|\r";
    static char[] TRIM_CHARS = { '\"' };

    public static List<Dictionary<string, object>> Read(string file)
    {
        var list = new List<Dictionary<string, object>>();
        TextAsset data = Resources.Load(file) as TextAsset;

        var lines = Regex.Split(data.text, LINE_SPLIT_RE);

        if (lines.Length <= 1) return list;

        var header = Regex.Split(lines[0], SPLIT_RE);
        for (var i = 1; i < lines.Length; i++)
        {

            var values = Regex.Split(lines[i], SPLIT_RE);
            if (values.Length == 0 || values[0] == "") continue;

            var entry = new Dictionary<string, object>();
            for (var j = 0; j < header.Length && j < values.Length; j++)
            {
                string value = values[j];
                value = value.TrimStart(TRIM_CHARS).TrimEnd(TRIM_CHARS).Replace("\\", "");
                object finalvalue = value;
                int n;
                float f;
                if (int.TryParse(value, out n))
                {
                    finalvalue = n;
                }
                else if (float.TryParse(value, out f))
                {
                    finalvalue = f;
                }
                entry[header[j]] = finalvalue;
            }
            list.Add(entry);
        }
        return list;
    }
}