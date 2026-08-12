namespace ExcelReportBuilder.AddIn.Ribbon
{
    internal static class RibbonMarkup
    {
        internal const string CustomUi = @"<?xml version=""1.0"" encoding=""utf-8""?>
<customUI xmlns=""http://schemas.microsoft.com/office/2009/07/customui"" onLoad=""OnRibbonLoad"">
  <ribbon>
    <tabs>
      <tab idMso=""TabData"">
        <group id=""ExcelReportBuilder.Group"" label=""PivotTable+"">
          <toggleButton id=""ExcelReportBuilder.OpenPane""
                        label=""PivotTable+""
                        size=""large""
                        imageMso=""PivotTableInsert""
                        screentip=""Open PivotTable+""
                        supertip=""Enhance the selected native PivotTable with familiar layout controls and validated extras.""
                        onAction=""OnToggleTaskPane""
                        getPressed=""GetTaskPanePressed""
                        getEnabled=""GetTaskPaneEnabled"" />
          <button id=""ExcelReportBuilder.OpenFieldList""
                  label=""Excel Field List""
                  imageMso=""PivotFieldListShowHide""
                  screentip=""Show Excel PivotTable Fields""
                  supertip=""Toggle Excel's built-in PivotTable Fields pane for familiar drag-and-drop editing.""
                  onAction=""OnOpenExcelFieldList""
                  getEnabled=""GetPivotActionEnabled"" />
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>";
    }
}
