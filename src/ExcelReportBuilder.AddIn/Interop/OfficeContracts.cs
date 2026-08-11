using System;
using System.Runtime.InteropServices;

namespace ExcelReportBuilder.AddIn.Interop
{
    /// <summary>
    /// Hand-written definition of the classic Office COM add-in contract.
    /// Keeping this boundary local avoids a deployment dependency on Office PIAs.
    /// </summary>
    [ComImport]
    [Guid("B65AD801-ABAF-11D0-BB8B-00A0C90F2744")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IDTExtensibility2
    {
        [DispId(1)]
        void OnConnection(
            [MarshalAs(UnmanagedType.IDispatch)] object application,
            ExtConnectMode connectMode,
            [MarshalAs(UnmanagedType.IDispatch)] object addInInstance,
            ref Array custom);

        [DispId(2)]
        void OnDisconnection(ExtDisconnectMode removeMode, ref Array custom);

        [DispId(3)]
        void OnAddInsUpdate(ref Array custom);

        [DispId(4)]
        void OnStartupComplete(ref Array custom);

        [DispId(5)]
        void OnBeginShutdown(ref Array custom);
    }

    public enum ExtConnectMode
    {
        AfterStartup = 0,
        Startup = 1,
        External = 2,
        CommandLine = 3,
        Solution = 4
    }

    public enum ExtDisconnectMode
    {
        HostShutdown = 0,
        UserClosed = 1,
        SetupChanged = 2,
        SolutionClosed = 3,
        Unloaded = 4
    }

    /// <summary>
    /// Hand-written Office ribbon interface used by Excel to request Ribbon XML.
    /// </summary>
    [ComImport]
    [Guid("000C0396-0000-0000-C000-000000000046")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IRibbonExtensibility
    {
        [DispId(1)]
        [return: MarshalAs(UnmanagedType.BStr)]
        string GetCustomUI([MarshalAs(UnmanagedType.BStr)] string ribbonId);
    }

    /// <summary>
    /// Office calls this interface after it has created the custom task-pane factory.
    /// The factory stays late bound so the add-in has no Office assembly reference.
    /// </summary>
    [ComImport]
    [Guid("000C033E-0000-0000-C000-000000000046")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface ICustomTaskPaneConsumer
    {
        void CTPFactoryAvailable([In, MarshalAs(UnmanagedType.Interface)] object taskPaneFactory);
    }

    [ComVisible(true)]
    [Guid("84D63B19-E890-42B8-A7D2-21736B282896")]
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface IExcelReportBuilderRibbonCallbacks
    {
        [DispId(1)]
        void OnRibbonLoad([MarshalAs(UnmanagedType.IDispatch)] object ribbonUi);

        [DispId(2)]
        void OnToggleTaskPane([MarshalAs(UnmanagedType.IDispatch)] object control, bool isPressed);

        [DispId(3)]
        bool GetTaskPanePressed([MarshalAs(UnmanagedType.IDispatch)] object control);

        [DispId(4)]
        bool GetTaskPaneEnabled([MarshalAs(UnmanagedType.IDispatch)] object control);
    }
}
