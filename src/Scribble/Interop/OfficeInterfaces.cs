using System;
using System.Runtime.InteropServices;

namespace Scribble.Interop
{
    public enum ExtConnectMode
    {
        AfterStartup = 0,
        Startup = 1,
        External = 2,
        CommandLine = 3
    }

    public enum ExtDisconnectMode
    {
        HostShutdown = 0,
        UserClosed = 1
    }

    [ComImport]
    [Guid("B65AD801-ABAF-11D0-BB8B-00A0C90F2744")]
    [TypeLibType(TypeLibTypeFlags.FDispatchable | TypeLibTypeFlags.FDual)]
    public interface IDTExtensibility2
    {
        [DispId(1)]
        void OnConnection(
            [In, MarshalAs(UnmanagedType.IDispatch)] object application,
            [In] ExtConnectMode connectMode,
            [In, MarshalAs(UnmanagedType.IDispatch)] object addInInstance,
            [In, Out, MarshalAs(
                UnmanagedType.SafeArray,
                SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom);

        [DispId(2)]
        void OnDisconnection(
            [In] ExtDisconnectMode removeMode,
            [In, Out, MarshalAs(
                UnmanagedType.SafeArray,
                SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom);

        [DispId(3)]
        void OnAddInsUpdate(
            [In, Out, MarshalAs(
                UnmanagedType.SafeArray,
                SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom);

        [DispId(4)]
        void OnStartupComplete(
            [In, Out, MarshalAs(
                UnmanagedType.SafeArray,
                SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom);

        [DispId(5)]
        void OnBeginShutdown(
            [In, Out, MarshalAs(
                UnmanagedType.SafeArray,
                SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom);
    }

    [ComImport]
    [Guid("000C0396-0000-0000-C000-000000000046")]
    [TypeLibType(TypeLibTypeFlags.FDispatchable | TypeLibTypeFlags.FDual)]
    public interface IRibbonExtensibility
    {
        [DispId(1)]
        [return: MarshalAs(UnmanagedType.BStr)]
        string GetCustomUI(
            [In, MarshalAs(UnmanagedType.BStr)] string ribbonId);
    }

    [ComImport]
    [Guid("000C033E-0000-0000-C000-000000000046")]
    [TypeLibType(
        TypeLibTypeFlags.FDispatchable |
        TypeLibTypeFlags.FNonExtensible |
        TypeLibTypeFlags.FDual)]
    public interface ICustomTaskPaneConsumer
    {
        [DispId(1)]
        void CTPFactoryAvailable(
            [In, MarshalAs(UnmanagedType.Interface)] object ctpFactory);
    }
}
