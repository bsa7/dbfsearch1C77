using System;
using System.EnterpriseServices;
using System.Reflection;
using System.Runtime.InteropServices;


namespace AddIn {
 internal static class HOST  {
  internal static Object app;
  internal static IAsyncEvent evt;
 }

 [Guid(@"AB634001-F13D-11D0-A459-004095E1DAEA"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
 public interface IInitDone  {
  void Init([MarshalAs(UnmanagedType.IDispatch)] Object pConnection);
  void Done();
  void GetInfo([MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref Object[] pInfo);
 }

 [Guid("AB634003-F13D-11d0-A459-004095E1DAEA"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
 public interface ILanguageExtender {
  void RegisterExtensionAs( [MarshalAs(UnmanagedType.BStr)] ref String extensionName);
  void GetNProps(ref Int32 props);
  void FindProp( [MarshalAs(UnmanagedType.BStr)] String propName, ref Int32 propNum);
  void GetPropName( Int32 propNum, Int32 propAlias, [MarshalAs(UnmanagedType.BStr)] ref String propName);
  void GetPropVal( Int32 propNum, [MarshalAs(UnmanagedType.Struct)] ref object propVal);
  void SetPropVal( Int32 propNum, [MarshalAs(UnmanagedType.Struct)] ref object propVal);
  void IsPropReadable(Int32 propNum, ref bool propRead);
  void IsPropWritable(Int32 propNum, ref Boolean propWrite);
  void GetNMethods(ref Int32 pMethods);
  void FindMethod( [MarshalAs(UnmanagedType.BStr)] String methodName, ref Int32 methodNum);
  void GetMethodName(Int32 methodNum, Int32 methodAlias, [MarshalAs(UnmanagedType.BStr)] ref String methodName);
  void GetNParams(Int32 methodNum, ref Int32 pParams);
  void GetParamDefValue( Int32 methodNum, Int32 paramNum, [MarshalAs(UnmanagedType.Struct)] ref object paramDefValue);
  void HasRetVal(Int32 methodNum, ref Boolean retValue);
  void CallAsProc( Int32 methodNum, [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType=VarEnum.VT_VARIANT)] ref object[] pParams);
  void CallAsFunc( Int32 methodNum, [MarshalAs(UnmanagedType.Struct)] ref object retValue, [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType=VarEnum.VT_VARIANT)] ref object[] pParams);
 }

 [Guid("AB634004-F13D-11D0-A459-004095E1DAEA"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
 public interface IAsyncEvent {
  void SetEventBufferDepth(Int32 depth);
  void GetEventBufferDepth(ref Int32 depth);
  void ExternalEvent([MarshalAs(UnmanagedType.BStr)] String source, [MarshalAs(UnmanagedType.BStr)] String message, [MarshalAs(UnmanagedType.BStr)] String data);
  void CleanBuffer();
 }

 [ComVisible(true), ProgId(@"AddIn.DBFSearch"), Guid(@"94ad95e2-a420-4212-8e38-4ebdc8728963"), ClassInterface(ClassInterfaceType.AutoDispatch)]
// public class DBFSearch : IInitDone {
 public class DBFSearch : IInitDone, ILanguageExtender {
  public DBFSearch() { }  

  public void FindMethod( [MarshalAs(UnmanagedType.BStr)] String methodName, ref Int32 methodNum){
   methodNum = (Int32)nameToNumber[methodName];
  }
  public void GetNParams(Int32 methodNum, ref Int32 pParams) {
   pParams = (Int32)numberToParams[methodNum];
  }
  public void HasRetVal(Int32 methodNum, ref Boolean retValue) {
   retValue = (Boolean)numberToRetVal[methodNum];
  }
  public void CallAsFunc( Int32 methodNum, [MarshalAs(UnmanagedType.Struct)] ref object retValue, [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType=VarEnum.VT_VARIANT)] ref object[] pParams) {
   retValue = allMethodInfo[(int)numberToMethodInfoIdx[methodNum]].Invoke(this, pParams);
  }
  public void Init([MarshalAs(UnmanagedType.IDispatch)] Object pConnection) {
   try {
    HOST.app = pConnection;
    HOST.evt = (IAsyncEvent)pConnection;
   }
   catch {
    throw new COMException(@"Unknown object context");
   }   
  }
 
  public void Done() {
   HOST.app = null;
   HOST.evt = null;
  }
 
  public void GetInfo([MarshalAs(UnmanagedType.SafeArray,  SafeArraySubType=VarEnum.VT_VARIANT)] ref Object[] pInfo) {
   pInfo[0] = @"2000"; /* просто так надо :) */
  }
 
  public void SayHello([MarshalAs(UnmanagedType.BStr)] String name) {
   try {
    /* Сгенерируем внешнее событие для 1С */
    HOST.evt.ExternalEvent(@"DBFSearch", @"SayHello", @"Hello, " + name + @" !");
   }
   catch {
    throw new COMException(@"Failed to execute SayHello() method");
   }   
  }
 }
}
