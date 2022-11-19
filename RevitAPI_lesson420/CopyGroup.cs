using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;

namespace RevitAPI_lesson420
{
    [TransactionAttribute(TransactionMode.Manual)]

    public class CopyGroup : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)

        {
            try
            {
                UIDocument uIDoc = commandData.Application.ActiveUIDocument;
                Document doc = uIDoc.Document;

                GroupPickFilter groupPickFilter = new GroupPickFilter();
                
                Reference reference = uIDoc.Selection.PickObject(ObjectType.Element, groupPickFilter, "Выберите группу объектов");
                Element element = doc.GetElement(reference);
                Group group = element as Group;

                XYZ groupCenter = GetElementCenter(group);
                Room roomBase = GetRoomByPoint(doc, groupCenter);
                XYZ roomBaseCenter = GetElementCenter(roomBase);
                XYZ offset = groupCenter - roomBaseCenter;

                XYZ pointUserSelected = uIDoc.Selection.PickPoint("Выберите комнату вставки группы");
                Room roomInsert = GetRoomByPoint(doc, pointUserSelected);
                XYZ pointInsert = GetElementCenter(roomInsert) + offset;

                using (Transaction transaction = new Transaction(doc))

                {
                    transaction.Start("Копирование группы объектов");
                    doc.Create.PlaceGroup(pointInsert, group.GroupType);
                    transaction.Commit();
                }
            }

            catch (Autodesk.Revit.Exceptions.OperationCanceledException)
            {
                return Result.Cancelled;
            }

            catch (Exception ex)
            {
                message = ex.Message;                
                return Result.Failed;
            }

            return Result.Succeeded;
        }

        public XYZ GetElementCenter (Element element)
        {
            BoundingBoxXYZ boundingBoxXYZ = element.get_BoundingBox(null);
            return ((boundingBoxXYZ.Max + boundingBoxXYZ.Min) / 2);
        }

        public Room GetRoomByPoint (Document document, XYZ point)
        {
            FilteredElementCollector collector = new FilteredElementCollector(document);
            collector.OfCategory(BuiltInCategory.OST_Rooms);
            foreach (Element element in collector)
            {
                Room room = element as Room;

                if (room == null)
                    continue;

                if (room.IsPointInRoom(point))
                    return room;
            }

            return null;
        }

    }
}
