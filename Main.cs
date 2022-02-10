using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Z5_ModelCreation
{
    [Transaction(TransactionMode.Manual)]
    public class Main : IExternalCommand
    {

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            // Получаем доступ к Revit, активному документу, базе данных документа
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;


            // запускаем транзакцию
            Transaction t = new Transaction(doc);
            t.Start("Create model");

            // получаем уровни
            Level levelBase = GetLevel(doc, "Уровень 1");
            Level levelTop = GetLevel(doc, "Уровень 2");

            // создаем стены  с заданными габаритами на любом уровне и до любого уровня
            var walls = CreateWalls(doc, levelBase, levelTop, 18000, 6000);

            // создаем дверь в первой стене
            AddDoors(doc, levelBase, walls[0]);
            // создаем окна в оставшихся стенах
            for (int i = 1; i <= 3; i++)
            {
                AddWindows(doc, levelBase, walls[i]);
            }

            t.Commit();

            return Result.Succeeded;
        }

        private static double mmToFeet(double lenght)
        {
            return UnitUtils.ConvertToInternalUnits(lenght, UnitTypeId.Millimeters);
        }

        private static List<Wall> CreateWalls(Document doc, Level levelBase, Level levelTop, double lenght, double width)
        {
            double dx = mmToFeet(lenght) * 0.5;
            double dy = mmToFeet(width) * 0.5;
            List<XYZ> wallPoints = new List<XYZ>();
            wallPoints.Add(new XYZ(-dx, -dy, 0));
            wallPoints.Add(new XYZ(-dx, dy, 0));
            wallPoints.Add(new XYZ(dx, dy, 0));
            wallPoints.Add(new XYZ(dx, -dy, 0));
            wallPoints.Add(wallPoints[0]);
            // список стен
            List<Wall> walls = new List<Wall>();
            // создаем стены на заданном уровне, высотой до требуемого уровня
            for (int i = 0; i < 4; i++)
            {
                Line line = Line.CreateBound(wallPoints[i], wallPoints[i + 1]);
                Wall wall = Wall.Create(doc, line, levelBase.Id, false);
                wall.get_Parameter(BuiltInParameter.WALL_HEIGHT_TYPE).Set(levelTop.Id);
                walls.Add(wall);
            }
            return walls;
        }

        private static void AddWindows(Document doc, Level levelBase, Wall wall)
        {
            FamilySymbol windowType = new FilteredElementCollector(doc)
                .OfClass(typeof(FamilySymbol))
                .OfCategory(BuiltInCategory.OST_Windows)
                .OfType<FamilySymbol>()
                .Where(d => d.Name.Equals("0915 x 1830 мм"))
                .Where(f => f.FamilyName.Equals("Фиксированные"))
                .FirstOrDefault();
            if (!windowType.IsActive)
            {
                windowType.Activate();
            }
            LocationCurve hostCurve = wall.Location as LocationCurve;
            XYZ insertPoint = (hostCurve.Curve.GetEndPoint(0) + hostCurve.Curve.GetEndPoint(1)) * 0.5;
            FamilyInstance window = doc.Create.NewFamilyInstance(insertPoint, windowType, wall, levelBase, Autodesk.Revit.DB.Structure.StructuralType.NonStructural);
            window.get_Parameter(BuiltInParameter.INSTANCE_SILL_HEIGHT_PARAM).Set(mmToFeet(900));
        }

        private static void AddDoors(Document doc, Level levelBase, Wall wall)
        {
            FamilySymbol doorType = new FilteredElementCollector(doc)
                .OfClass(typeof(FamilySymbol))
                .OfCategory(BuiltInCategory.OST_Doors)
                .OfType<FamilySymbol>()
                .Where(d => d.Name.Equals("0915 x 2134 мм"))
                .Where(f => f.FamilyName.Equals("Одиночные-Щитовые"))
                .FirstOrDefault();
            if (!doorType.IsActive)
            {
                doorType.Activate();
            }
            LocationCurve hostCurve =  wall.Location as LocationCurve;
            XYZ insertPoint = (hostCurve.Curve.GetEndPoint(0) + hostCurve.Curve.GetEndPoint(1)) * 0.5;
            doc.Create.NewFamilyInstance(insertPoint, doorType, wall, levelBase, Autodesk.Revit.DB.Structure.StructuralType.NonStructural);
        }

        private static Level GetLevel (Document doc, string levelName)
        {
            //получаем все уровни
            return new FilteredElementCollector(doc)
                .OfClass(typeof(Level))
                .OfType<Level>()
                .Where(l => l.Name.Equals(levelName))
                .FirstOrDefault();
        }
    }
}
