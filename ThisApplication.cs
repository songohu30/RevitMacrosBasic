using System;
using Autodesk.Revit.UI;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI.Selection;
using Autodesk.Revit.DB.Plumbing;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using System.Threading.Tasks;

namespace RevitMacrosBasic
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    [Autodesk.Revit.DB.Macros.AddInId("12CB6F4B-7135-403D-9BC1-8D11A7B07CE9")]
	public partial class ThisApplication
	{
		private void Module_Startup(object sender, EventArgs e)
		{

		}

		private void Module_Shutdown(object sender, EventArgs e)
		{

		}

		#region Revit Macros generated code
		private void InternalStartup()
		{
			this.Startup += new System.EventHandler(Module_Startup);
			this.Shutdown += new System.EventHandler(Module_Shutdown);
		}
		#endregion
		
		#region FilteredElementCollector examples
		/// <summary>
		/// Selects walls of specified thickness range in active view
		/// </summary>
		public void SelectWallsInActiveView()
		{			
			//input in millimeters
			double minThickness = 100;
			double maxThickness = 200;
			
			Document doc = ActiveUIDocument.Document;
			FilteredElementCollector ficol = new FilteredElementCollector(doc, doc.ActiveView.Id); //elements visible in active view
			//FilteredElementCollector ficol = new FilteredElementCollector(doc); //elements in entire document
			
			List<ElementId> wallIds = ficol.OfClass(typeof(Wall)).Select(e => e as Wall).Where(e => IsWithinThicknessRange(e, minThickness, maxThickness)).Select(e => e.Id).ToList();
			ActiveUIDocument.Selection.SetElementIds(wallIds);			
		}
		
		/// <summary>
		/// Checks if wall thickness is not less than min or more than max
		/// </summary>
		/// <param name="wall"></param>
		/// <param name="minThickness"></param>
		/// <param name="maxThickness"></param>
		/// <returns>bool</returns>
		private bool IsWithinThicknessRange(Wall wall, double minThickness, double maxThickness)
		{
			bool result = true;
			WallType wallType = wall.WallType;
			Parameter widthParameter = wallType.get_Parameter(BuiltInParameter.WALL_ATTR_WIDTH_PARAM);
			double thicknessInternalUnit = widthParameter.AsDouble();
			double thicknessMetric = UnitUtils.ConvertFromInternalUnits(thicknessInternalUnit, UnitTypeId.Millimeters);
			if(thicknessMetric < minThickness || thicknessMetric > maxThickness) result = false;			
			return result;
		}
		
		/// <summary>
		/// Selects pipes of specified diameter range in active view
		/// </summary>
		public void SelectPipesInActiveView()
		{
			//input in millimeters
			double minDiameter = 100;
			double maxDiameter = 200;
			
			Document doc = ActiveUIDocument.Document;
			FilteredElementCollector ficol = new FilteredElementCollector(doc, doc.ActiveView.Id); //elements visible in active view
			//FilteredElementCollector ficol = new FilteredElementCollector(doc); //elements in entire document
			
			List<ElementId> pipeIds = ficol.OfClass(typeof(Pipe)).Select(e => e as Pipe).Where(e => IsWithinDiameterRange(e, minDiameter, maxDiameter)).Select(e => e.Id).ToList();
			ActiveUIDocument.Selection.SetElementIds(pipeIds);	
		}
		
		/// <summary>
		/// Checks if pipe diameter is not less than min or more than max
		/// </summary>
		/// <param name="pipe"></param>
		/// <param name="minDiameter"></param>
		/// <param name="maxDiameter"></param>
		/// <returns></returns>
		private bool IsWithinDiameterRange(Pipe pipe, double minDiameter, double maxDiameter)
		{
			bool result = true;
			Parameter diameterPar = pipe.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM);
			double diameterInternalUnit = diameterPar.AsDouble();
			double diameterMetric = UnitUtils.ConvertFromInternalUnits(diameterInternalUnit, UnitTypeId.Millimeters);
			if(diameterMetric < minDiameter || diameterMetric > maxDiameter) result = false;			
			return result;
		}
		#endregion
		
		#region Parameters - Create, Add, Read, Write
		/// <summary>
		/// Adds new shared parameters to an existing shared parameter file in Revit '22 '23
		/// </summary>
		public void AddNewSharedParameters()
		{
			//input:
			string sharedParameterGroupName = "TestGroup";
			
			//in revit version '22, '23
			Dictionary<string, ForgeTypeId> sharedParameters = new Dictionary<string, ForgeTypeId>();
			sharedParameters.Add("TestParameter1", SpecTypeId.String.Text);
			sharedParameters.Add("TestParameter2", SpecTypeId.Number);
			sharedParameters.Add("TestParameter3", SpecTypeId.Area);
			
			//in revit version '19, '20, '21
//			Dictionary<string, ParameterType> sharedParameters = new Dictionary<string, ParameterType>();
//			sharedParameters.Add("TestParameter1", ParameterType.Text);
//			sharedParameters.Add("TestParameter2", ParameterType.Number);
//			sharedParameters.Add("TestParameter3", ParameterType.Area);

			DefinitionFile definitionFile = Application.OpenSharedParameterFile();
			
			if(definitionFile == null)
			{
				TaskDialog.Show("Error", "Shared parameters file is not set!");
				return;
			}
			
				if(definitionFile.Groups.Select(e => e.Name).Contains(sharedParameterGroupName))
				{
					DefinitionGroup group = definitionFile.Groups.Where(e => e.Name == sharedParameterGroupName).FirstOrDefault();
					
					foreach (var sharedParameter in sharedParameters) 
					{
						if(group.Definitions.Select(e => e.Name).Contains(sharedParameter.Key)) continue;
						ExternalDefinitionCreationOptions parameterCreation = new ExternalDefinitionCreationOptions(sharedParameter.Key, sharedParameter.Value);
						group.Definitions.Create(parameterCreation);
					}
				}
				else
				{
					TaskDialog.Show("Error", string.Format("Group: {0}, does  not exist in current shared paramter file", sharedParameterGroupName));
					return;
				}		
		}
		/// <summary>
		/// Adds existing shared parameters to specified categories in current project
		/// </summary>
		public void AddSharedParametersToProject()
		{
			Document doc = ActiveUIDocument.Document;
			//input:
			string[] sharedParameterNames = {"TestParameter1", "TestParameter2", "TestParameter3"};
			
			BuiltInParameterGroup parameterGroup = BuiltInParameterGroup.INVALID; //Enum group to add parameters Invalid = Other
			
			//example categories to add project parameter
			Category category1_toBindParameter = doc.Settings.Categories.get_Item(BuiltInCategory.OST_PipeFitting);
			Category category2_toBindParameter = doc.Settings.Categories.get_Item(BuiltInCategory.OST_PipeCurves);
			Category category3_toBindParameter = doc.Settings.Categories.get_Item(BuiltInCategory.OST_Walls);
			Category category4_toBindParameter = doc.Settings.Categories.get_Item(BuiltInCategory.OST_Views);
			
			CategorySet catset = Application.Create.NewCategorySet();
			catset.Insert(category1_toBindParameter);
			catset.Insert(category2_toBindParameter);
			catset.Insert(category3_toBindParameter);
			catset.Insert(category4_toBindParameter);
			Binding binding = Application.Create.NewInstanceBinding(catset);
			
			DefinitionFile definitionFile = this.Application.OpenSharedParameterFile();
			List<Definition> foundDefinitions = new List<Definition>();
			
			foreach (DefinitionGroup group in definitionFile.Groups) 
			{
				List<Definition> definitions = group.Definitions.Where(e => sharedParameterNames.Contains(e.Name)).ToList();
				if(definitions != null && definitions.Count > 0)
				{
					foundDefinitions.AddRange(definitions);
				}
			}
			
			using(Transaction tx = new Transaction(doc, "New project parameters"))
			{
				tx.Start();
				foreach (string pName in sharedParameterNames) 
				{
					if(foundDefinitions.Select(e => e.Name).Contains(pName))
					{
						Definition definition = foundDefinitions.Where(e => e.Name == pName).FirstOrDefault();
						doc.ParameterBindings.Insert(definition, binding, parameterGroup);
					}
					else
					{
						TaskDialog.Show("Info", "Shared parameters file does not contain: " + pName);
					}
				}
				tx.Commit();
			}
		}
		
		public void ReadParameterValuesFromSelectedElement()
		{			
			ICollection<ElementId> selectedIds = ActiveUIDocument.Selection.GetElementIds();
			if(selectedIds.Count == 0)
			{
				TaskDialog.Show("Info", "No elements selected!");
			}

			string parameterName = "Comments"; //we can search by name
			BuiltInParameter builtInParamter = BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS; //we can search by builtinparamter enum: revit language independent
			
			SearchBy searchBy = SearchBy.BuiltInParameter;
			//SearchBy searchBy = SearchBy.ParameterName;
			
			string report = "Found parameter values:" + Environment.NewLine;
			
			foreach (ElementId id in selectedIds) 
			{
				Element element = ActiveUIDocument.Document.GetElement(id);
				switch (searchBy) 
				{
					case SearchBy.ParameterName:
						Parameter p1 = element.LookupParameter(parameterName); //parameters may have the same name: first will be returned
						report += "Element Id: " + id.IntegerValue.ToString() + ", Element name: " + element.Name + ", Parameter value: " + ReadParameterValue(p1) + Environment.NewLine + Environment.NewLine;
						
						break;
					case SearchBy.BuiltInParameter:
						Parameter p2 = element.get_Parameter(builtInParamter);
						report += "Element Id: " + id.IntegerValue.ToString() + ", Element name: " + element.Name + ", Parameter value: " + ReadParameterValue(p2) + Environment.NewLine + Environment.NewLine;
						break;
					default:
						
						break;
				}
			}
			TaskDialog.Show("ReportInfo", report);
		}
			
		private string ReadParameterValue(Parameter parameter)
		{
			string result = "parameter not found";
			if(parameter == null) return result;
			switch (parameter.StorageType) 
			{
				case StorageType.String:
					result = parameter.AsString();
					break;
				case StorageType.Double:	
					result = parameter.AsDouble().ToString();
					break;
				case StorageType.ElementId:
					result = parameter.AsElementId().IntegerValue.ToString();
					break;					
				case StorageType.Integer:
					result = parameter.AsInteger().ToString();
					break;
				default:
					result = "Unknown storage type";
					break;
			}
			if(string.IsNullOrEmpty(result)) result = "empty";
			return result;
		}
		/// <summary>
		/// Set shared parameters values in selected elements
		/// </summary>
		public void SetParameterValuesForSelectedElements()
		{
			Document doc = ActiveUIDocument.Document;
			
			string[] sharedParameterNames = {"TestParameter1", "TestParameter2", "TestParameter3"};
			ICollection<ElementId> selectedIds = ActiveUIDocument.Selection.GetElementIds();
			if(selectedIds.Count == 0)
			{
				TaskDialog.Show("Info", "No elements selected!");
			}
			DefinitionFile definitionFile = this.Application.OpenSharedParameterFile();
			List<Definition> foundDefinitions = new List<Definition>();
			
			foreach (DefinitionGroup group in definitionFile.Groups) 
			{
				List<Definition> definitions = group.Definitions.Where(e => sharedParameterNames.Contains(e.Name)).ToList();
				if(definitions != null && definitions.Count > 0)
				{
					foundDefinitions.AddRange(definitions);
				}
			}
			
			if(foundDefinitions.Count == 0) TaskDialog.Show("Info", "None of shared parameter was found!");
			
			using (Transaction tx = new Transaction(doc, "Set parameter values"))
			{
				tx.Start();
				foreach (ElementId id in selectedIds) 
				{
					Element element = doc.GetElement(id);
					foreach (Definition pDefinition in foundDefinitions) 
					{
						Parameter p = element.get_Parameter(pDefinition);
						if(p != null)
						{
							switch(p.StorageType)
							{
								case StorageType.String:
									p.Set("Hello World!");
									break;
								case StorageType.Double:
									ForgeTypeId forgeTypeId = pDefinition.GetDataType();
									if(forgeTypeId == SpecTypeId.Number)
									{
										double x = UnitUtils.ConvertToInternalUnits(100, UnitTypeId.Meters);
										p.Set(x);
									}
									if(forgeTypeId == SpecTypeId.Area)
									{
										double x = UnitUtils.ConvertToInternalUnits(100, UnitTypeId.SquareMeters);
										p.Set(x);
									}
//In revit '19, '20, '21
//									ParameterType parameterType = definition.ParameterType;
//									if(parameterType == ParameterType.Number)
//									{
//										double x = UnitUtils.ConvertToInternalUnits(100, DisplayUnitType.DUT_MILLIMETERS);
//										p.Set(x);
//									}
//									if(parameterType == ParameterType.Area)
//									{
//										double x = UnitUtils.ConvertToInternalUnits(100, DisplayUnitType.DUT_SQUARE_METERS);
//										p.Set(x);
//									}
									break;
								default:
									break;
							}
						}
					}
				}
				tx.Commit();
			}			
		}
		
		enum SearchBy
		{
			ParameterName,
			BuiltInParameter
		}
		
		#endregion
	}
}