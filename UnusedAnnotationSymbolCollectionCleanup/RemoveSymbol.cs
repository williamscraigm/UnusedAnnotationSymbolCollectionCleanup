using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using System.Runtime.InteropServices;
using ESRI.ArcGIS.Carto;
using ESRI.ArcGIS.Geodatabase;
using ESRI.ArcGIS.Display;
using ESRI.ArcGIS.esriSystem;

namespace UnusedAnnotationSymbolCollectionCleanup
{
  public class RemoveSymbol : ESRI.ArcGIS.Desktop.AddIns.Button
  {
    public RemoveSymbol()
    {
    }

    protected override void OnClick()
    {
      ILayer layer = ArcMap.Document.SelectedLayer;
      IAnnotationLayer annoLayer = layer as IAnnotationLayer;
      if (annoLayer != null)
      {
        ProcessAnnotation(annoLayer);
      }
    }

    protected override void OnUpdate()
    {
      ILayer layer = ArcMap.Document.SelectedLayer;
      IAnnotationLayer annoLayer = layer as IAnnotationLayer;
      if (annoLayer != null)
      {
        Enabled = true;
      }
      else
      {
        Enabled = false;
      }
    }
    private void ProcessAnnotation(IAnnotationLayer annoLayer)
    {
      
      //get unique values
      IFeatureLayer featureLayer = annoLayer as IFeatureLayer;
      IDataStatistics dataStats = new DataStatisticsClass();
      dataStats.Field = "SymbolID";
      dataStats.SampleRate = -1; //all records
      IQueryFilter queryFilter = new QueryFilterClass();
      queryFilter.SubFields = "SymbolID";
      dataStats.Cursor = featureLayer.Search(queryFilter, true) as ICursor;

      object value = null;
      var enumVar = dataStats.UniqueValues;

      //now remove items
      IFeatureClass featureClass = featureLayer.FeatureClass;
      IAnnotationClassExtension annoClassExt = featureClass.Extension as IAnnotationClassExtension;
      ISymbolCollection2 symbolCollection = annoClassExt.SymbolCollection as ISymbolCollection2;

      ISymbolCollection2 newSymColl = new SymbolCollectionClass();
      
      while (enumVar.MoveNext())
      {
        value = enumVar.Current;
        ISymbolIdentifier2 symbolIdent = null;
        symbolCollection.GetSymbolIdentifier(Convert.ToInt32(value), out symbolIdent);
        newSymColl.set_Symbol(symbolIdent.ID, symbolIdent.Symbol);
        newSymColl.RenameSymbol(symbolIdent.ID, symbolIdent.Name);
      }

      //update class extension with the new collection
      IAnnoClassAdmin3 annoClassAdmin = annoClassExt as IAnnoClassAdmin3;
      annoClassAdmin.SymbolCollection = (ISymbolCollection)newSymColl;
      annoClassAdmin.UpdateProperties();


    }

  }
}
