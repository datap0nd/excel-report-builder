using System;

namespace ExcelReportBuilder.Excel.PivotPlus.NamedSets
{
    /// <summary>
    /// Exact-object discovery entry point. Callers pass the workbook and
    /// PivotTable selected during the trusted host workflow; no active-cell or
    /// active-workbook lookup exists in this service.
    /// </summary>
    internal sealed class PivotNamedSetSchemaDiscoveryService
    {
        private readonly IPivotNamedSetGateway gateway;

        public PivotNamedSetSchemaDiscoveryService()
            : this(new LateBoundPivotNamedSetGateway())
        {
        }

        internal PivotNamedSetSchemaDiscoveryService(IPivotNamedSetGateway gateway)
        {
            this.gateway = gateway ?? throw new ArgumentNullException(nameof(gateway));
        }

        public PivotNamedSetSchemaDiscoveryResult Discover(
            object workbook,
            object pivotTable,
            PivotTableContext context)
        {
            BoundPivotNamedSetTarget target = gateway.Bind(
                workbook,
                pivotTable,
                context);
            return gateway.DiscoverSchema(target);
        }
    }
}
