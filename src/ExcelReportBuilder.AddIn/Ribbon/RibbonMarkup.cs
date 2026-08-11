namespace ExcelReportBuilder.AddIn.Ribbon
{
    internal static class RibbonMarkup
    {
        internal const string CustomUi = @"<?xml version=""1.0"" encoding=""utf-8""?>
<customUI xmlns=""http://schemas.microsoft.com/office/2009/07/customui"" onLoad=""OnRibbonLoad"">
  <ribbon>
    <tabs>
      <tab idMso=""TabData"">
        <group id=""ExcelReportBuilder.Group"" label=""Report Builder"">
          <toggleButton id=""ExcelReportBuilder.OpenPane""
                        label=""Report Builder""
                        size=""large""
                        imageMso=""PivotTableInsert""
                        screentip=""Open Excel Report Builder""
                        supertip=""Choose data, build a dense management report, work with the assistant, and review checks.""
                        onAction=""OnToggleTaskPane""
                        getPressed=""GetTaskPanePressed""
                        getEnabled=""GetTaskPaneEnabled"" />
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>";
    }
}
