//TYPELIB "RegExpSearchInDBF.tlb" 
using System;
using System.Reflection;
using System.Runtime.InteropServices;

namespace AddIn {
  internal static class ExternalAddIn {
    internal static Object App; //Объект внешней компоненты
    internal static IAsyncEvent ExtEvent; //Внешнее событие
  }

  [Guid(@"AB634001-F13D-11D0-A459-004095E1DAEA"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
  //Жизненно важный интерфейс 
  public interface IInitDone { 
    void Init([MarshalAs(UnmanagedType.IDispatch)] Object pConnection);
    void Done(); //А здесь указана не реализованная процедура.
    void GetInfo([MarshalAs(UnmanagedType.SafeArray,
    SafeArraySubType = VarEnum.VT_VARIANT)] ref Object[] pInfo);
  }

  [Guid("AB634004-F13D-11D0-A459-004095E1DAEA"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
  public interface IAsyncEvent {
    void SetEventBufferDepth(Int32 depth);
    void GetEventBufferDepth(ref Int32 depth);
    void ExternalEvent([MarshalAs(UnmanagedType.BStr)] String source, [MarshalAs(UnmanagedType.BStr)] String message, [MarshalAs(UnmanagedType.BStr)] String data);
    void CleanBuffer();
  }

  [ComVisible(true), ProgId(@"AddIn.RegExpSearchInDBF"), Guid(@"A0ED3412-327B-40E5-82CC-C55E44C1DF28"), ClassInterface(ClassInterfaceType.AutoDispatch)]
  // [ComVisible(true), ProgId(@"AddIn.DotNetAddIn"), Guid(@"9F0CF3B4-B799-4852-8293-9BB9500A3098"), ClassInterface(ClassInterfaceType.AutoDispatch)]
  public class RegExpSearchInDBF : IInitDone {
    public void RegExpSearchInDBFSample() { } 
    //Инициализация внешней компоненты
    public void Init([MarshalAs(UnmanagedType.IDispatch)] Object pConnection) { 
      try {
        ExternalAddIn.App = pConnection;
        ExternalAddIn.ExtEvent = (IAsyncEvent)pConnection;
      }
      catch {
        throw new COMException(@"Unknown object context");
      } 
    }

    //Освобождение ресурсов при уничтожении экземпляра компоненты в 1С
    public void Done() { 
      ExternalAddIn.App = null;
      ExternalAddIn.ExtEvent = null;
    }

    public void GetInfo([MarshalAs(UnmanagedType.SafeArray, SafeArraySubType=VarEnum.VT_VARIANT)] ref Object[] pInfo) {
      pInfo[0] = @"2000";
    }
    //Метод компоненты, вызываемый в 1С
    public void TestExtEvent([MarshalAs(UnmanagedType.BStr)] String msg) { 
      try {
        //Генерация внешнего события
        ExternalAddIn.ExtEvent.ExternalEvent(@"RegExpSearchInDBF", @"TestExtEvent", @"Test message for, " + msg + @" !");
      }
      catch {
        throw new COMException(@"Failed to execute TestExtEvent() method");
      } 
    }
  }
}
