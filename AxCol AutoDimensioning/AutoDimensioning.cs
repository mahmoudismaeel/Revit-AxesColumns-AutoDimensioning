using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;

namespace AxCol_AutoDimensioning
{
    [Transaction(TransactionMode.Manual)]

    public class AutoDimensioning : IExternalCommand  
    {

        [Autodesk.Revit.Attributes.Regeneration(Autodesk.Revit.Attributes.RegenerationOption.Manual)]
        [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]

        public class helper
        {
            public ElementId id { get; set; }
            public double value { get; set; }

            public XYZ pt { get; set; }

            public helper(ElementId id, double value, XYZ pt)
            {
                this.id = id;
                this.value = value;
                this.pt = pt;
            }
        }

        // GETTING GRIDS' NAMES
        public void GridNamesToString(string locationMarkName, out string g1Name, out string g2Name)   
        {
            g1Name = "";
            g2Name = "";
            IList<char> part1 = new List<char>();
            IList<char> part2 = new List<char>();

            for (int i = 0; i < locationMarkName.Length; i++)                       // IT CAN BE LIKE  9(-300)-AA(600)
            {
                if (locationMarkName[i] == '(' || locationMarkName[i] == '-')
                {
                    break;
                }

                part1.Add(locationMarkName[i]);
            }

            for (int i = 0; i < locationMarkName.Length; i++)
            {
                if (locationMarkName[i] == '-' && locationMarkName[i - 1] != '(')
                {
                    i++;
                    string x = locationMarkName.Substring(i);
                    for (int j = 0; j < x.Length; j++)
                    {
                        if (x[j] == '(')
                        {
                            break;
                        }

                        part2.Add(x[j]);
                    }
                }
            }

            g1Name = new string(part1.ToArray());
            g2Name = new string(part2.ToArray());
        }



        // CHECKING PARALLELISM
        public bool IsParallel(XYZ dir1, XYZ dir2)
        {
            if (dir1.AngleTo(dir2) < .0000001 || Math.Abs(dir1.AngleTo(dir2)) - Math.PI < .0000001)
            {
                return true;
            }

            return false;
        }


        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            Options op = new Options();
            op.ComputeReferences = true;
            op.IncludeNonVisibleObjects = true;
            op.View = doc.ActiveView;

            // SETTING ILISTS OF DIMENSIONS FOR EVERY DIRCETION
            IList<helper> lID1 = new List<helper>();
            IList<helper> lID2 = new List<helper>();


            // GETTING COLUMNS FROM THE ACTIVE VIEW
            FilteredElementCollector columns = new FilteredElementCollector(doc, doc.ActiveView.Id).
                                                              OfCategory(BuiltInCategory.OST_StructuralColumns).
                                                              WhereElementIsNotElementType();

            // GETTING GRIDS FROM THE ACTIVE VIEW
            FilteredElementCollector grids = new FilteredElementCollector(doc, doc.ActiveView.Id).
                                                              OfCategory(BuiltInCategory.OST_Grids).
                                                              WhereElementIsNotElementType();


            string locationMark;
            string g1Name;
            string g2Name;
            XYZ pt1 = new XYZ();
            XYZ pt2 = new XYZ();

            using (Transaction t = new Transaction(doc, "Dims"))
            {
                t.Start();
                foreach (FamilyInstance column in columns)
                {

                    // GETTING THE LOCATION MARK FOR EVERY COLUMN
                    locationMark = column.get_Parameter(BuiltInParameter.COLUMN_LOCATION_MARK).AsString();
                    GridNamesToString(locationMark, out g1Name, out g2Name);

                    // SETTING GRIDS' REFERENCES  
                    Grid g1 = grids.Cast<Grid>().FirstOrDefault<Grid>(x => x.Name == g1Name);
                    Reference g1Ref = new Reference(g1);
                    Grid g2 = grids.Cast<Grid>().FirstOrDefault<Grid>(x => x.Name == g2Name);
                    Reference g2Ref = new Reference(g2);

                    GeometryElement geoEle = column.get_Geometry(op);
                    foreach (GeometryObject geoObj in geoEle)
                    {
                        if (geoObj.GetType() == typeof(GeometryInstance))
                        {
                            GeometryInstance geoinst = geoObj as GeometryInstance;
                            GeometryElement geometryElement = geoinst.GetSymbolGeometry();
                            XYZ ptt = geoinst.Transform.Origin;

                            foreach (GeometryObject geometryObject in geometryElement)
                            {
                                if (geometryObject.GetType() == typeof(Solid))                // IT HAS A POINT, LINE, AND TWO SOLIDS
                                {
                                    Solid solid = geometryObject as Solid;

                                    if (solid.SurfaceArea != 0)                               // ONE HAS ZEROS AND THE OTHER HAS VALUES (FACES AND EDGES)
                                    {
                                        foreach (Face face in solid.Faces)
                                        {
                                            PlanarFace planarFace = face as PlanarFace;
                                            if (planarFace.FaceNormal.IsAlmostEqualTo(new XYZ(0, 0, -1)))      // A RECT. COLUMS HAS 6 FACES, I SELECT THE TOP ONE
                                            {
                                                EdgeArrayArray earrarr = face.EdgeLoops;
                                                EdgeArray earr = earrarr.get_Item(0);                          // EACH RECT. FACE HAS 4 EDGES 

                                                foreach (Edge edge in earr)                                    
                                                {
                                                    Curve c1 = edge.AsCurve();                                 // CONVERTING EDGE TO CURVE TO GET THE END POINT
                                                    XYZ edgeP = c1.GetEndPoint(0);

                                                    Line lc1 = c1 as Line;
                                                    XYZ edgeLineDir = lc1.Direction;                           // CONVERTING CURVE TO LINE TO GET THE DIRECTION

                                                    Curve g1Curv = g1.Curve;                                   // CONVERTING GRIDS TO CURVE THEN CURVES TO LINES TO GET DIRECTIONS
                                                    Curve g2Curv = g2.Curve;
                                                    Line lg1 = g1Curv as Line;
                                                    Line lg2 = g2Curv as Line;
                                                    XYZ g1LineDir = lg1.Direction;
                                                    XYZ g2LineDir = lg2.Direction;

                                                    IntersectionResult intersec1 = g1Curv.Project(edgeP);      // GETTING THE EDGE PROJECTED POINTS ON GRIDS
                                                    IntersectionResult intersec2 = g2Curv.Project(edgeP);
                                                    XYZ g1P = intersec1.XYZPoint;                              // SETTING POINTS ON THOSE PROJECTED POINTS
                                                    XYZ g2P = intersec2.XYZPoint;

                                                    Line l1 = Line.CreateBound(edgeP, g1P);                    // CREATING LINES TO DRAW THE DIMENSIONS ON BETWEEN EDGE AND GRIDS
                                                    Line l2 = Line.CreateBound(edgeP, g2P);
                                                    Reference edgeReff = edge.Reference;

                                                    if (IsParallel(edgeLineDir, g1LineDir))                    // CHECK WHICH GRID IS PARALLEL TO THE EDGE AND CREATE A DIMENSION BETWEEN
                                                    {
                                                        ReferenceArray refArr = new ReferenceArray();
                                                        refArr.Append(edgeReff);
                                                        refArr.Append(g1Ref);

                                                        Dimension dim1 = doc.Create.NewDimension(doc.ActiveView, l1, refArr);

                                                        pt1 = ptt - g1P;
                                                        if (dim1.Value != -1)
                                                        { 
                                                            lID1.Add(new helper(dim1.Id, (double)dim1.Value, pt1));         // ADDING THE DIMENSIONS OF EVERY DIRECTION ON A LIST
                                                        }
                                                    }


                                                    if (IsParallel(edgeLineDir, g2LineDir))                                  // SAME AS ABOVE FOR THE ANOTHER DIRECTION
                                                    {
                                                        ReferenceArray refArr = new ReferenceArray();
                                                        refArr.Append(edgeReff);
                                                        refArr.Append(g2Ref);

                                                        Dimension dim2 = doc.Create.NewDimension(doc.ActiveView, l2, refArr);

                                                        pt2 = ptt - g2P;
                                                        if (dim2.Value != -1)
                                                        { 
                                                            lID2.Add(new helper(dim2.Id, (double)dim2.Value, pt2));
                                                        }
                                                    }
                                                }

                                                // SORTING THE LIST OF DIMENSIONS OF EVERY DIRECTION AND GETTING THE SHORTEST ONE
                                                helper del1 = lID1.OrderBy(s => s.value).First();
                                                ElementTransformUtils.MoveElement(doc, del1.id, del1.pt + new XYZ(3, 3, 0));
                                                helper del2 = lID2.OrderBy(s => s.value).First();
                                                ElementTransformUtils.MoveElement(doc, del2.id, del2.pt + new XYZ(3, 3, 0));

                                                // GETTING THE LONGEST DIMENSION FROM THE LIST OF EVERY DIRECTION AND DELETE IT
                                                del1 = lID1.OrderBy(s => s.value).Last();
                                                del2 = lID2.OrderBy(s => s.value).Last();
                                                doc.Delete(del1.id);
                                                doc.Delete(del2.id);

                                                // CHECKING IF THERE IS AN ALIGNED DIMENSION TO A GRID TO DELETE IT
                                                foreach (var item in lID1)
                                                {
                                                    if (item.value < 0.00001)
                                                    {
                                                        doc.Delete(item.id);
                                                    }
                                                }
                                                foreach (var item in lID2)
                                                {
                                                    if (item.value < 0.000001)
                                                    {
                                                        doc.Delete(item.id);
                                                    }
                                                }

                                                lID1.Clear();
                                                lID2.Clear();
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }

                t.Commit();
            }

            return Result.Succeeded;
        }
    }
}
