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
            [In, MarshalAs(UnmanagedType.IDispatch)] object application,
            [In] ExtConnectMode connectMode,
            [In, MarshalAs(UnmanagedType.IDispatch)] object addInInstance,
            [In, MarshalAs(
                UnmanagedType.SafeArray,
                SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom);

        [DispId(2)]
        void OnDisconnection(
            [In] ExtDisconnectMode removeMode,
            [In, MarshalAs(
                UnmanagedType.SafeArray,
                SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom);

        [DispId(3)]
        void OnAddInsUpdate(
            [In, MarshalAs(
                UnmanagedType.SafeArray,
                SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom);

        [DispId(4)]
        void OnStartupComplete(
            [In, MarshalAs(
                UnmanagedType.SafeArray,
                SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom);

        [DispId(5)]
        void OnBeginShutdown(
            [In, MarshalAs(
                UnmanagedType.SafeArray,
                SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom);
    }

    public enum ExtConnectMode
    {
        AfterStartup = 0,
        Startup = 1,
        External = 2,
        CommandLine = 3,
        Solution = 4,
        UISetup = 5
    }

    public enum ExtDisconnectMode
    {
        HostShutdown = 0,
        UserClosed = 1,
        UISetupComplete = 2,
        SolutionClosed = 3
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
        string GetCustomUI([In, MarshalAs(UnmanagedType.BStr)] string ribbonId);
    }

    /// <summary>
    /// Office custom task-pane factory supplied to the add-in at startup.
    /// </summary>
    [ComImport]
    [Guid("000C033D-0000-0000-C000-000000000046")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface ICTPFactory
    {
        [DispId(1)]
        [return: MarshalAs(UnmanagedType.Interface)]
        object CreateCTP(
            [In, MarshalAs(UnmanagedType.BStr)] string controlProgramId,
            [In, MarshalAs(UnmanagedType.BStr)] string title,
            [In, Optional, MarshalAs(UnmanagedType.Struct)] object parentWindow);
    }

    /// <summary>
    /// Office calls this interface after it has created the custom task-pane factory.
    /// The local COM contract avoids a deployment dependency on the Office PIA.
    /// </summary>
    [ComImport]
    [Guid("000C033E-0000-0000-C000-000000000046")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface ICustomTaskPaneConsumer
    {
        [DispId(1)]
        void CTPFactoryAvailable(
            [In, MarshalAs(UnmanagedType.Interface)] ICTPFactory taskPaneFactory);
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

        [DispId(5)]
        void OnOpenExcelFieldList([MarshalAs(UnmanagedType.IDispatch)] object control);

        [DispId(6)]
        bool GetPivotActionEnabled([MarshalAs(UnmanagedType.IDispatch)] object control);
    }
}
