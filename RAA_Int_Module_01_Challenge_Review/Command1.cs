#region Namespaces
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Reflection;
using Autodesk.Revit.DB.Architecture;

#endregion

namespace RAA_Int_Module_01_Challenge_Review
{
    [Transaction(TransactionMode.Manual)]
    public class Command1 : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            // this is a variable for the Revit application
            UIApplication uiapp = commandData.Application;

            // this is a variable for the current Revit model
            Document doc = uiapp.ActiveUIDocument.Document;

            // 1. Get all rooms
            List<Element> roomList = GetAllRooms(doc);

            // 2. Loop through rooms and get dept names
            int counter = 0;
            List<string> allDeptList = new List<string>();
            foreach(Room room in roomList)
            {
                string curDept = room.LookupParameter("Department").AsString();
                allDeptList.Add(curDept);
            }

            List<string> deptList = allDeptList.Distinct().ToList();

            // 3. Loop through dept list 
            using (Transaction t = new Transaction(doc))
            {
                t.Start("Create schedule");
                foreach (string deptName in deptList)
                {
                    // 4. Create department schedules
                    ViewSchedule curSchedule = CreateRoomScheduleByDepartment(doc, deptName);
                    counter++;

                }

                // 5. Create all department schedule
                ViewSchedule curSchedule2 = CreateDepartmentSchedule(doc);

                t.Commit();
            }
            TaskDialog.Show("Complete", $"Created {counter} schedules.");
            return Result.Succeeded;
        }

        private ViewSchedule CreateDepartmentSchedule(Document doc)
        {
            // 5a. Create schedule
            ViewSchedule curSched = CreateSchedule(doc, BuiltInCategory.OST_Rooms, "All Departments");

            // 5b. Get instance of room
            Element room = GetAllRooms(doc).First();

            // 5c. Add fields to schedule
            ScheduleField deptField = AddFieldToSchedule(curSched, ScheduleFieldType.Instance, GetParameterByName(room, "Department"), false);
            ScheduleField areaField = AddFieldToSchedule(curSched, ScheduleFieldType.ViewBased, GetParameterByName(room, BuiltInParameter.ROOM_AREA), false);

            areaField.DisplayType = ScheduleFieldDisplayType.Totals;

            // 5d. Sort schedule by department
            ScheduleSortGroupField deptSort = new ScheduleSortGroupField(deptField.FieldId, ScheduleSortOrder.Ascending);
            curSched.Definition.AddSortGroupField(deptSort);

            curSched.Definition.IsItemized = false;
            curSched.Definition.ShowGrandTotal = true;
            curSched.Definition.ShowGrandTotalTitle = true;

            return curSched;
        }

        private ViewSchedule CreateRoomScheduleByDepartment(Document doc, string deptName)
        {
            // 4a. Create schedule
            ViewSchedule curSched = CreateSchedule(doc, BuiltInCategory.OST_Rooms, $"Department - {deptName}");

            // 4b. Get instance of room
            Element room = GetAllRooms(doc).First();

            // 4c. Add fields to schedule
            ScheduleField roomNumField = AddFieldToSchedule(curSched, ScheduleFieldType.Instance, GetParameterByName(room, "Number"), false);
            ScheduleField roomNameField = AddFieldToSchedule(curSched, ScheduleFieldType.Instance, GetParameterByName(room, "Name"), false);
            ScheduleField deptField = AddFieldToSchedule(curSched, ScheduleFieldType.Instance, GetParameterByName(room, "Department"), false);
            ScheduleField commentsField = AddFieldToSchedule(curSched, ScheduleFieldType.Instance, GetParameterByName(room, "Comments"), false);
            ScheduleField areaField = AddFieldToSchedule(curSched, ScheduleFieldType.ViewBased, GetParameterByName(room, BuiltInParameter.ROOM_AREA), false);
            ScheduleField levelField = AddFieldToSchedule(curSched, ScheduleFieldType.Instance, GetParameterByName(room, "Level"), true);

            areaField.DisplayType = ScheduleFieldDisplayType.Totals;

            // 4d. filter schedule by department
            ScheduleFilter deptFilter = new ScheduleFilter(deptField.FieldId, ScheduleFilterType.Equal, deptName);
            curSched.Definition.AddFilter(deptFilter);

            // 4e. sort and group schedule
            ScheduleSortGroupField levelSort = new ScheduleSortGroupField(levelField.FieldId, ScheduleSortOrder.Ascending);
            levelSort.ShowHeader = true;
            levelSort.ShowFooter = true;
            levelSort.ShowBlankLine = true;
            curSched.Definition.AddSortGroupField(levelSort);

            ScheduleSortGroupField roomNameSort = new ScheduleSortGroupField(roomNameField.FieldId, ScheduleSortOrder.Ascending);
            curSched.Definition.AddSortGroupField(roomNameSort);

            // 4f. set grand total properties
            curSched.Definition.IsItemized = true;
            curSched.Definition.ShowGrandTotal = true;
            curSched.Definition.ShowGrandTotalTitle = true;
            curSched.Definition.ShowGrandTotalCount = true;

            return curSched;
        }

        private ScheduleField AddFieldToSchedule(ViewSchedule curSched, ScheduleFieldType fieldType, Parameter parameter, bool isHidden)
        {
            ScheduleField curField = curSched.Definition.AddField(fieldType, parameter.Id);
            curField.IsHidden = isHidden;

            return curField;
        }

        private Parameter GetParameterByName(Element curElem, string name)
        {
            Parameter curParam = curElem.LookupParameter(name);

            return curParam;
        }

        private Parameter GetParameterByName(Element curElem, BuiltInParameter bip)
        {
            Parameter curParam = curElem.get_Parameter(bip);

            return curParam;
        }

        private ViewSchedule CreateSchedule(Document doc, BuiltInCategory bic, string name)
        {
            ElementId catId = new ElementId(bic);
            ViewSchedule newSchedule = ViewSchedule.CreateSchedule(doc, catId);
            newSchedule.Name = name;

            return newSchedule;
        }

        private List<Element> GetAllRooms(Document doc)
        {
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            collector.OfCategory(BuiltInCategory.OST_Rooms);
            
            return collector.ToList();
        }

        internal static PushButtonData GetButtonData()
        {
            // use this method to define the properties for this command in the Revit ribbon
            string buttonInternalName = "btnCommand1";
            string buttonTitle = "Button 1";

            ButtonDataClass myButtonData1 = new ButtonDataClass(
                buttonInternalName,
                buttonTitle,
                MethodBase.GetCurrentMethod().DeclaringType?.FullName,
                Properties.Resources.Blue_32,
                Properties.Resources.Blue_16,
                "This is a tooltip for Button 1");

            return myButtonData1.Data;
        }
    }
}
