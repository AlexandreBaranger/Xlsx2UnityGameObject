using UnityEditor;
using UnityEngine;
using System.Collections.Generic; // Pour Dictionary
using System.Data;
using System.IO;
using ExcelDataReader;

public class DuplicateGameObjectsEditor : EditorWindow
{
    private string excelFilePath = "Assets/Data/ObjectsToDuplicate.xlsx";
    private GameObject baseObject;
    private GameObject newParent;

    [MenuItem("Tools/Duplicate GameObjects from Excel")]
    public static void ShowWindow()
    {
        GetWindow<DuplicateGameObjectsEditor>("Duplicate GameObjects from Excel");
    }

    private void OnGUI()
    {
        GUILayout.Label("Duplicate GameObjects from Excel", EditorStyles.boldLabel);

        baseObject = (GameObject)EditorGUILayout.ObjectField("Base GameObject", baseObject, typeof(GameObject), true);
        newParent = (GameObject)EditorGUILayout.ObjectField("New Parent GameObject", newParent, typeof(GameObject), true);
        excelFilePath = EditorGUILayout.TextField("Excel File Path", excelFilePath);

        if (GUILayout.Button("Duplicate GameObjects"))
        {
            if (baseObject != null && newParent != null)
            {
                DuplicateGameObjects();
            }
            else
            {
                Debug.LogError("Please assign both the Base GameObject and the New Parent GameObject.");
            }
        }
    }

    private void DuplicateGameObjects()
    {
        if (!File.Exists(excelFilePath))
        {
            Debug.LogError("The Excel file does not exist at the specified path!");
            return;
        }

        FileStream fileStream = new FileStream(excelFilePath, FileMode.Open, FileAccess.Read);
        IExcelDataReader reader = ExcelReaderFactory.CreateOpenXmlReader(fileStream);
        DataSet result = reader.AsDataSet();
        reader.Close();

        DataTable table = result.Tables[0];

        Transform[] allChildren = baseObject.GetComponentsInChildren<Transform>(true);

        Dictionary<string, GameObject> gameObjectsInHierarchy = new Dictionary<string, GameObject>();

        foreach (Transform child in allChildren)
        {
            if (child.gameObject != baseObject)
            {
                if (!gameObjectsInHierarchy.ContainsKey(child.gameObject.name))
                {
                    gameObjectsInHierarchy.Add(child.gameObject.name, child.gameObject);
                }
            }
        }

        for (int i = 1; i < table.Rows.Count; i++)
        {
            DataRow row = table.Rows[i];
            if (row == null) continue;

            string objectNameWithExtension = row[7].ToString();
            string objectName = objectNameWithExtension.Replace(".csv", "");

            if (gameObjectsInHierarchy.ContainsKey(objectName))
            {
                GameObject originalObject = gameObjectsInHierarchy[objectName];
                GameObject duplicatedObject = Instantiate(originalObject);

                duplicatedObject.transform.SetParent(newParent.transform);
                duplicatedObject.transform.localPosition = Vector3.zero;

                // Mark the object as dirty so it is saved in the scene
                EditorUtility.SetDirty(duplicatedObject);

                Debug.Log($"Duplicated object '{objectName}' under '{newParent.name}'.");
            }
            else
            {
                Debug.LogWarning($"Object '{objectName}' not found in hierarchy!");
            }
        }

        fileStream.Close();
        Debug.Log("Duplication process completed.");
    }
}
