using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Reflection;
/*[assembly:AssemblyKeyFile("addin1c.snk")]
[assembly:AssemblyVersion("1.0.0.0")]      */
namespace AddIn {
 /// <summary>
 /// Summary description for Class1.
 /// </summary>
 [Guid("AB634001-F13D-11d0-A459-004095E1DAEA"),
 InterfaceType(ComInterfaceType.InterfaceIsDual)]
 public interface IInitDone {
  void  Init([MarshalAs(UnmanagedType.IDispatch)]object pConnection);
  void Done();
  void GetInfo(object []pInfo);
 }
 // IStatusLine Interface
 [Guid("AB634005-F13D-11D0-A459-004095E1DAEA"),
 InterfaceType(ComInterfaceType.InterfaceIsDual)]
 public interface IStatusLine {
  void  SetStatusLine(string  bstrSource);
  void  ResetStatusLine();
 }
  
 public interface ITestComponent {
  String About { get; }
  string SayHello();
 }
 [ComVisible(true), ProgId("AddIn.dbfregsearch"),Guid("8B3630A9-6B78-4d8c-B418-CC10F974F412")]
 public class dbfregsearch :IInitDone, ITestComponent {
  public object V7Object;
  IStatusLine pStatusLine;
  public void Init([MarshalAs(UnmanagedType.IDispatch)] object pConnection) {
   try {
    V7Object = pConnection;
    System.Windows.Forms.MessageBox.Show("hkhjkhjkh");
   }
   catch {
    throw new COMException("Unknown V7 object context", 0);
   }
   pStatusLine = null;
   //AB634005-F13D-11D0-A459-004095E1DAEA
   pStatusLine = (IStatusLine) pConnection;
   pStatusLine.SetStatusLine("fjdfgdfgfdkgjkl");
  }
 
  public void Done() {
   V7Object = null;
  }
 
  public void GetInfo(object []pInfo) {
   pInfo.SetValue("1000", 0);
  }
  public string SayHello() {
   return "Hello world!";
  }
  public String About {
   get {
    try {
     pStatusLine = (IStatusLine) V7Object;
     pStatusLine.SetStatusLine("About");
    }
    catch {
     throw new COMException("Unknown V7 object context", 0);
    }
    return " test";
   }
  }		
 }
}