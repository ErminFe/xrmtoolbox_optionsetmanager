using System;
using System.Collections.Generic;
using System.ComponentModel.Composition;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using XrmToolBox.Extensibility;
using XrmToolBox.Extensibility.Interfaces;

namespace smartpoint.XrmToolBoxPlugins.OptionSetManager
{
  [Export(typeof(IXrmToolBoxPlugin)),
  ExportMetadata("BackgroundColor", "MediumBlue"),
  ExportMetadata("PrimaryFontColor", "White"),
  ExportMetadata("SecondaryFontColor", "LightGray"),
  ExportMetadata("SmallImageBase64", null),
  ExportMetadata("BigImageBase64", null),
  ExportMetadata("Name", "OptionSetManager"),
  ExportMetadata("Description", "Enables the manipulation of option sets in a bulk fashion.")]
  public class Plugin : PluginBase
  {
    public override IXrmToolBoxPluginControl GetControl()
    {
      return new PluginControl();
    }
  }
}
