using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB.ExtensibleStorage;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using Nice3point.Revit.Toolkit.External;
using OperationCanceledException = Autodesk.Revit.Exceptions.OperationCanceledException;

namespace ExtensibleStorage.Commands;


/// <summary>
///     Command to read extensible storage data from a user-selected Revit element.
///     Displays stored schema data using a TaskDialog.
/// </summary>
[UsedImplicitly]
[Transaction(TransactionMode.Manual)]
internal class ReadExtensibleStorage : ExternalCommand
{
    /// <summary>
    ///     Entry point for the external command.
    ///     Prompts the user to select an element, then reads and displays extensible storage data.
    /// </summary>
    public override void Execute()
    {
        try
        {
            // Prompt user to select an element in the active document
            Reference reference = UiDocument.Selection.PickObject(ObjectType.Element, "Please select an element");
            Element element = Document.GetElement(reference);

            
            // Retrieve the schema by name and vendor ID
            Schema schema = Schema.ListSchemas()
                .FirstOrDefault(s => s.SchemaName == "MySchema" && s.VendorId == "KHALID"); // Replace with your actual VendorId

            if (schema == null)
            {
                TaskDialog.Show("Schema Not Found", "The schema 'MySchema' was not found.");
                return;
            }

            // Validate selected element
            if (!element.IsValidObject)
            {
                TaskDialog.Show("Invalid Element", "The selected element cannot store external data.");
                return;
            }

            // Try to get the associated entity from the element
            Entity entity = element.GetEntity(schema);

            if (!entity.IsValid())
            {
                TaskDialog.Show("No Data", "This element does not contain extensible storage data.");
                return;
            }

            // Retrieve and format the stored values
#if R21_OR_GREATER
            XYZ location = entity.Get<XYZ>(schema.GetField("Location"), UnitTypeId.Meters);
            double length = entity.Get<double>(schema.GetField("WireLength"), UnitTypeId.Meters);
#else
var location = entity.Get<XYZ>(schema.GetField("Location"), DisplayUnitType.DUT_METERS);
var length = entity.Get<double>(schema.GetField("WireLength"), DisplayUnitType.DUT_METERS);
#endif

            string description = entity.Get<string>(schema.GetField("Description"));
            string timeString = entity.Get<string>(schema.GetField("SpliceTime"));

            // Build a message to show in the TaskDialog
            string msg = $"📌 Extensible Storage Data:\n\n" +
                         $"- Location: ({location.X:F2}, {location.Y:F2}, {location.Z:F2})\n" +
                         $"- Wire Length: {length:F2} m\n" +
                         $"- Description: {description}\n" +
                         $"- Splice Time: {timeString}";

            // Show the data in a Revit dialog
            TaskDialog.Show("Wire Splice Info", msg);
        }
        catch (OperationCanceledException)
        {
            // User canceled the selection
            TaskDialog.Show("Canceled", "Operation canceled by user.");
        }
        catch (Exception ex)
        {
            // Unexpected error occurred
            TaskDialog.Show("Error", ex.Message);
            throw;
        }
    }
}