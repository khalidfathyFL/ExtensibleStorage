using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB.ExtensibleStorage;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using Nice3point.Revit.Toolkit.External;
using OperationCanceledException = Autodesk.Revit.Exceptions.OperationCanceledException;

namespace ExtensibleStorage.Commands;

/// <summary>
///     Writes extensible storage data to a user-selected element in a Revit document.
///     The data includes spatial and descriptive metadata such as location, length, a note, and a timestamp.
/// </summary>
[UsedImplicitly]
[Transaction(TransactionMode.Manual)]
internal class WriteExtensibleStorage : ExternalCommand
{
    /// <summary>
    ///     The entry point for the external command.
    ///     Prompts the user to select an element, creates a schema if necessary,
    ///     and writes a set of extensible storage fields to that element.
    /// </summary>
    public override void Execute()
    {
        // Step 1: Create a new schema builder with a unique GUID
        SchemaBuilder schemaBuilder = new SchemaBuilder(Guid.NewGuid());

        // Define schema metadata
        schemaBuilder.SetVendorId("Khalid"); // MUST match the VendorId in your .addin file
        schemaBuilder.SetSchemaName("MySchema"); // Unique schema name
        schemaBuilder.SetReadAccessLevel(AccessLevel.Public);
        schemaBuilder.SetWriteAccessLevel(AccessLevel.Vendor);

        // Step 2: Add fields to the schema

        // Field: Location (XYZ)
        FieldBuilder locationFieldBuilder = schemaBuilder.AddSimpleField("Location", typeof(XYZ));
#if R21_OR_GREATER
        locationFieldBuilder.SetSpec(SpecTypeId.Length);
#else
        locationFieldBuilder.SetUnitType(UnitType.UT_Length);
#endif
        locationFieldBuilder.SetDocumentation("3D point representing the splice location.");

        // Field: WireLength (double)
        FieldBuilder lengthFieldBuilder = schemaBuilder.AddSimpleField("WireLength", typeof(double));
#if R21_OR_GREATER
        lengthFieldBuilder.SetSpec(SpecTypeId.Length);
#else
        lengthFieldBuilder.SetUnitType(UnitType.UT_Length);
#endif
        lengthFieldBuilder.SetDocumentation("Total wire length from the splice to the endpoint.");

        // Field: Description (string)
        FieldBuilder descriptionFieldBuilder = schemaBuilder.AddSimpleField("Description", typeof(string));

        descriptionFieldBuilder.SetDocumentation("Human-readable description of the wire splice.");

        // Field: SpliceTime (string — ISO 8601 format of DateTime)
        FieldBuilder timeFieldBuilder = schemaBuilder.AddSimpleField("SpliceTime", typeof(string));

        timeFieldBuilder.SetDocumentation("The date and time of the splice in ISO-8601 format.");

        // Step 3: Finalize the schema
        Schema schema = schemaBuilder.Finish();

        // Step 4: Create an entity based on the schema and assign field values
        Entity entity = new Entity(schema);

#if R21_OR_GREATER
        entity.Set(schema.GetField("Location"), new XYZ(1, 2, 3), UnitTypeId.Meters);
        entity.Set(schema.GetField("WireLength"), 12.5, UnitTypeId.Meters);
#else
        entity.Set(schema.GetField("Location"), new XYZ(1, 2, 3), DisplayUnitType.DUT_METERS);
        entity.Set(schema.GetField("WireLength"), 12.5, DisplayUnitType.DUT_METERS);
#endif

        entity.Set(schema.GetField("Description"), "Main splice near junction box");
        entity.Set(schema.GetField("SpliceTime"), DateTime.Now.ToString("o")); // ISO format

        // Step 5: Prompt user to select an element to apply the data
        try
        {
            Reference reference = UiDocument.Selection.PickObject(ObjectType.Element, "Please select an element");
            Element element = Document.GetElement(reference);

            // Write the entity data to the selected element inside a transaction
            using Transaction transaction = new Transaction(Document);
            transaction.Start("Write Extensible Storage");
            element.SetEntity(entity);
            transaction.Commit();
        }
        catch (OperationCanceledException)
        {
            // User canceled the element selection
            TaskDialog.Show("Canceled", "Operation was canceled by user.");
        }
        catch (Exception e)
        {
            // Handle any unexpected errors
            TaskDialog.Show("Error", e.Message);
            throw;
        }
    }
}