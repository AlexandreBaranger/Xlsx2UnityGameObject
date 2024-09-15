using UnityEditor;
using UnityEngine;
using System.Collections.Generic; // Pour Dictionary
using System.Data;
using System.IO;
using ExcelDataReader;

public class CreateGameObjectsFromExcelEditor : EditorWindow
{
    private string excelFilePath = "Assets/Data/ObjectsToCreate.xlsx";
    private GameObject newParent;

    [MenuItem("Tools/Create GameObjects from Excel")]
    public static void ShowWindow()
    {
        GetWindow<CreateGameObjectsFromExcelEditor>("Create GameObjects from Excel");
    }

    private void OnGUI()
    {
        GUILayout.Label("Create GameObjects from Excel", EditorStyles.boldLabel);

        newParent = (GameObject)EditorGUILayout.ObjectField("New Parent GameObject", newParent, typeof(GameObject), true);
        excelFilePath = EditorGUILayout.TextField("Excel File Path", excelFilePath);

        if (GUILayout.Button("Create GameObjects"))
        {
            if (newParent != null)
            {
                CreateGameObjects();
            }
            else
            {
                Debug.LogError("Please assign a New Parent GameObject.");
            }
        }
    }

    private void CreateGameObjects()
    {
        if (!File.Exists(excelFilePath))
        {
            Debug.LogError("The Excel file does not exist at the specified path!");
            return;
        }

        // Lire le fichier Excel
        FileStream fileStream = new FileStream(excelFilePath, FileMode.Open, FileAccess.Read);
        IExcelDataReader reader = ExcelReaderFactory.CreateOpenXmlReader(fileStream);
        DataSet result = reader.AsDataSet();
        reader.Close();

        // Récupérer la première table (feuille) du fichier Excel
        DataTable table = result.Tables[0];

        // Parcourir les lignes du fichier Excel
        for (int i = 1; i < table.Rows.Count; i++) // Ignorer la première ligne (titres)
        {
            DataRow row = table.Rows[i];
            if (row == null) continue;

            // Lire le nom de l'objet à partir de la colonne du fichier Excel
            string objectName = row[0].ToString(); // On suppose que la première colonne contient les noms

            if (!string.IsNullOrEmpty(objectName))
            {
                // Créer un nouveau GameObject avec le nom lu dans le fichier Excel
                GameObject newObject = new GameObject(objectName);

                // Le mettre sous l'objet parent
                newObject.transform.SetParent(newParent.transform);
                newObject.transform.localPosition = Vector3.zero; // Positionner à zéro relatif au parent

                // Marquer l'objet comme "Dirty" pour qu'il soit enregistré dans la scène
                EditorUtility.SetDirty(newObject);

                Debug.Log($"Created new object '{objectName}' under '{newParent.name}'.");
            }
            else
            {
                Debug.LogWarning("Found an empty row in the Excel file!");
            }
        }

        fileStream.Close();
        Debug.Log("GameObject creation process completed.");
    }
}
