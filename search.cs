using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.IO;
using System.Numerics;
using System.Collections;
using System.EnterpriseServices;

//Эта строка для проверки работы веток на github

/*
[assembly: Guid("b50eb73d-7977-4c4c-8216-686fd4d454fc")]
[assembly: ApplicationID("b50eb73d-7977-4c4c-8216-686fd4d454fc")]
 
[assembly: ApplicationAccessControl(AccessChecksLevel=AccessChecksLevelOption.Application)]
[assembly: ApplicationName("Fulltext regexp search in 1C77 dbf")]
[assembly: ApplicationActivation(ActivationOption.Library)]
[assembly: AssemblyKeyFile("key.snk")]
[Guid("99df1c52-85f3-4ed5-8d31-d8d74c3789cc")]
[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
  */
namespace SlonProgs {
/*
 public interface IInitDone {
  void Init(
    [MarshalAs(UnmanagedType.IDispatch)]
    object connection);

  void Done();

 void GetInfo(
    [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType=VarEnum.VT_VARIANT)]
    ref object[] info);
}

Guid("99df1c53-85f3-4ed5-8d31-d8d74c3789cc")]
[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
public interface ILanguageExtender
{

  void RegisterExtensionAs(
    [MarshalAs(UnmanagedType.BStr)]
    ref String extensionName);


  void GetNProps(ref Int32 props);


  void FindProp(
    [MarshalAs(UnmanagedType.BStr)]
    String propName,
    ref Int32 propNum);


  void GetPropName(
    Int32 propNum,
    Int32 propAlias,
    [MarshalAs(UnmanagedType.BStr)]
    ref String propName);


  void GetPropVal(
    Int32 propNum,
    [MarshalAs(UnmanagedType.Struct)]
    ref object propVal);


  void SetPropVal(
    Int32 propNum,
    [MarshalAs(UnmanagedType.Struct)]
    ref object propVal);


  void IsPropReadable(Int32 propNum, ref bool propRead);


  void IsPropWritable(Int32 propNum, ref Boolean propWrite);


  void GetNMethods(ref Int32 pMethods);


  void FindMethod(
    [MarshalAs(UnmanagedType.BStr)]
    String methodName,
    ref Int32 methodNum);


  void GetMethodName(Int32 methodNum,
    Int32 methodAlias,
    [MarshalAs(UnmanagedType.BStr)]
    ref String methodName);


  void GetNParams(Int32 methodNum, ref Int32 pParams);

  void GetParamDefValue(
    Int32 methodNum,
    Int32 paramNum,
    [MarshalAs(UnmanagedType.Struct)]
    ref object paramDefValue);


  void HasRetVal(Int32 methodNum, ref Boolean retValue);


  void CallAsProc(
    Int32 methodNum,
    [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType=VarEnum.VT_VARIANT)]
    ref object[] pParams);


  void CallAsFunc(
    Int32 methodNum,
    [MarshalAs(UnmanagedType.Struct)]
    ref object retValue,
    [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType=VarEnum.VT_VARIANT)]
    ref object[] pParams);
}




[ClassInterface(ClassInterfaceType.None)]
[ProgId("AddIn.SimpleExternalComponent")]
//[Description("SimpleExternalComponent.Component class")]

public class Component : ServicedComponent,IInitDone
{

   //private IAsyncEvent asyncEvent    = null;
   //private IStatusLine statusLine    = null;
   Hashtable nameToNumber;
   Hashtable numberToName;
   Hashtable numberToParams;
   Hashtable numberToRetVal;
   Hashtable propertyNameToNumber;
   Hashtable propertyNumberToName;
   Hashtable numberToMethodInfoIdx;
   Hashtable propertyNumberToPropertyInfoIdx;
   MethodInfo[] allMethodInfo;
   PropertyInfo[] allPropertyInfo;
   

   public Component(){
    
   }

   public void RegisterExtensionAs([MarshalAs(UnmanagedType.BStr)]
               ref String extensionName)
   {
        
        try{
    // initialize data members
            nameToNumber      = new Hashtable();
            numberToName      = new Hashtable();
            numberToParams    = new Hashtable();
            numberToRetVal    = new Hashtable();

            propertyNameToNumber  = new Hashtable();
            propertyNumberToName  = new Hashtable();

            numberToMethodInfoIdx  = new Hashtable();
            propertyNumberToPropertyInfoIdx  = new Hashtable();

            // Заполнение хэш-таблиц
            Type type = this.GetType();
            Type[] allInterfaceTypes = type.GetInterfaces();

            
            // Определение идентификатора
            int Identifier = 0;

            foreach(Type interfaceType in allInterfaceTypes)
            {
                if (
                    !interfaceType.Name.Equals("IDisposable")
                    && !interfaceType.Name.Equals("IManagedObject")
                    && !interfaceType.Name.Equals("IRemoteDispatch")
                    && !interfaceType.Name.Equals("IServicedComponentInfo")
                    && !interfaceType.Name.Equals("IInitDone")
                    && !interfaceType.Name.Equals("ILanguageExtender")
                   )
                {
                    // Обработка методов
                    MethodInfo[] interfaceMethods = interfaceType.GetMethods();
                    foreach (MethodInfo interfaceMethodInfo in interfaceMethods)
                    {
                      nameToNumber.Add(interfaceMethodInfo.Name, Identifier);
                      numberToParams.Add(Identifier,
                      interfaceMethodInfo.GetParameters().Length);
                      if (typeof(void) != interfaceMethodInfo.ReturnType)
                        numberToRetVal.Add(Identifier, true);

                      Identifier++;
                    }
                    

                    // Обработка свойств
                    PropertyInfo[] interfaceProperties = interfaceType.GetProperties();
                    foreach (PropertyInfo interfacePropertyInfo in interfaceProperties)
                    {
                      propertyNameToNumber.Add(interfacePropertyInfo.Name, Identifier);

                      Identifier++;
                    }
                    
                }
            }

    foreach (DictionaryEntry entry in nameToNumber)
      numberToName.Add(entry.Value, entry.Key);
    foreach (DictionaryEntry entry in propertyNameToNumber)
      propertyNumberToName.Add(entry.Value, entry.Key);

    // Сохранение информации о методах класса
    allMethodInfo = type.GetMethods();

    // Сохранение информации о свойствах класса
    allPropertyInfo = type.GetProperties();

    
    // Отображение номера метода на индекс в массиве
    foreach (DictionaryEntry entry in numberToName)
    {
      bool found = false;
      for(int ii = 0; ii < allMethodInfo.Length; ii++)
      {
        if (allMethodInfo[ii].Name.Equals(entry.Value.ToString()))
        {
          numberToMethodInfoIdx.Add(entry.Key, ii);
          found = true;
        }
      }
      if (false == found)
        throw new COMException("Метод не реализован ");
    }
    
    // Отображение номера свойства на индекс в массиве
    foreach (DictionaryEntry entry in propertyNumberToName)
    {
      bool found = false;
      for(int ii = 0; ii < allPropertyInfo.Length; ii++)
      {
        if (allPropertyInfo[ii].Name.Equals(entry.Value.ToString()))
        {
          propertyNumberToPropertyInfoIdx.Add(entry.Key, ii);
          found = true;
        }
      }
      if (false == found)
        throw new COMException("The property is not implemented");
    }
    
    // Компонент инициализирован успешно
    // Возвращаем имя компонента
    
    extensionName  = "print";//componentName;
  }
  catch (Exception)
  {
    
    return;
  }
}


   public void Init([MarshalAs(UnmanagedType.IDispatch)]
     object connection)
   {
   

   }


   public void GetInfo([MarshalAs(UnmanagedType.SafeArray, SafeArraySubType=VarEnum.VT_VARIANT)]
     ref object[] info)
   {
  
   info[0] = 1000;

   }

   public void Done(){
 
   }

   public void FindMethod([MarshalAs(UnmanagedType.BStr)]
      String methodName,
    ref Int32 methodNum)
   {
  
   methodNum = (Int32)nameToNumber[methodName];
   }
    

   public void GetNParams(Int32 methodNum, ref Int32 pParams)
   {
   
      pParams = (Int32)numberToParams[methodNum];
   }


   public void HasRetVal(Int32 methodNum, ref Boolean retValue)
   {
   
      retValue = (Boolean)numberToRetVal[methodNum];
   }


   public void CallAsFunc(Int32 methodNum,
       [MarshalAs(UnmanagedType.Struct)]
       ref object retValue,
       [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType=VarEnum.VT_VARIANT)]
       ref object[] pParams)
   {
     
      retValue = allMethodInfo[(int)numberToMethodInfoIdx[methodNum]].Invoke(
       this, pParams);


   }

   public void GetNProps(ref Int32 props){
   
   }

   public void FindProp([MarshalAs(UnmanagedType.BStr)]
    String propName,
    ref Int32 propNum){
    
    }


   public void GetPropName(
    Int32 propNum,
    Int32 propAlias,
    [MarshalAs(UnmanagedType.BStr)]
    ref String propName){
    
   }


  public void GetPropVal(
    Int32 propNum,
    [MarshalAs(UnmanagedType.Struct)]
    ref object propVal){
    
  }

  public void SetPropVal(
    Int32 propNum,
    [MarshalAs(UnmanagedType.Struct)]
    ref object propVal){
    
  }

  public void IsPropReadable(Int32 propNum, ref bool propRead){
 
  }

  public void IsPropWritable(Int32 propNum, ref Boolean propWrite){
 
  }

  public void GetNMethods(ref Int32 pMethods){
  
  }

  public void GetMethodName(Int32 methodNum,
    Int32 methodAlias,
    [MarshalAs(UnmanagedType.BStr)]
    ref String methodName){
    
  }

  public void GetParamDefValue(
    Int32 methodNum,
    Int32 paramNum,
    [MarshalAs(UnmanagedType.Struct)]
    ref object paramDefValue){
    
  }

  public void CallAsProc(
    Int32 methodNum,
    [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType=VarEnum.VT_VARIANT)]
    ref object[] pParams){
    
  }
}
// Наш класс работающий в 1с
public class MyComponent:Component{
    public MyComponent(){}
    public void Print(){
        MessageBox.Show("Компонента работает нормально");
    }
}



  */
//namespace SlonProgs {

 public class SlonInDBFSearch {

  public struct DBFFieldDescription {
   public string NameOfField;
   public byte TypeOfField;
   public long FieldLength;
  }

  public struct DBFProperties {
   public long FileSize;
   public long RecordsCount;
   public long HeaderSize;
   public int RecordSize;
   public int FieldCount;
   public ArrayList FieldsProperties;
  }

  static void Main(string[] args) {
   string filepath=@"SC156.DBF";
   //Пример "быстрого" выражения (в начале выражения - никаких неопределённостей по количеству символов!)
   string pattern=@"(Н|н|H)\s*(А|а|A|a)\s*(Р|р|P|p)\s*(K|k|К|к)\s*(О|о|O|o)[^((М|м|M)(Ы|ы|O|o|О|о|У|у|Y|y|C|c|С|с))]\S*";
   //Пример медленных выражений. Всё дело в неопределённом заранее количестве неопределённых заранее символов (здесь 0 и более непробельных символов до начала шаблона)
   //    pattern=@"\S*?(Н|н|H)\s*(А|а|A|a)\s*(Р|р|P|p)\s*(K|k|К|к)\s*(О|о|O|o)\S*";
   //    pattern=@"\S*?(Н|н|H)\s*(А|а|A|a)\s*(Р|р|P|p)\s*(K|k|К|к)\s*(О|о|O|o)\S*";
   //Пример Недопустимых или бессмысленных выражений (Поиск приведёт только в начало или конец файла с вероятностью 0 целых хуй десятых %)
   //    pattern=@"^(Н|н|H)\s*(А|а|A|a)\s*(Р|р|P|p)\s*(K|k|К|к)\s*(О|о|O|o)\S*";
   //    pattern=@"(Н|н|H)\s*(А|а|A|a)\s*(Р|р|P|p)\s*(K|k|К|к)\s*(О|о|O|o)\S*$";
   if (args.Length==0) {
    Console.WriteLine("Поиск в DBF.");
    Console.WriteLine("Неверный вызов. Используй следующий формат параметров:");
    Console.WriteLine("search.exe <[Путь]ИмяФайла> <РегВыражениеПоиска>");
    Console.WriteLine("");
    Console.WriteLine("... Тем не менее - делаем тестовый прогон (наличие в текущей папке файла SC156.DBF обязательно!)");
    Console.WriteLine("Ищем наркотики ({0}) в файле {1}", pattern, filepath);
   }
   if (args.Length>=1) {
    filepath=args[0];
   }
   if (File.Exists(filepath)) {
    if (args.Length==2) {
     pattern=args[1];
    }
    ArrayList Matches=searchInDBF(filepath, pattern);
   }
   else { 
    Console.WriteLine("Файл {0} не найден.", filepath);
   }
  }

  // ************************************************************************************************************************************************
  // Функция поиска
  //
  static ArrayList searchInDBF(string filepath, string pattern) {
   string text="";
   DBFProperties DBFheader;
   try {
    text = System.IO.File.ReadAllText(filepath, Encoding.GetEncoding(1251)); // Тут могут возникнуть проблемы, если у тебя не Windows8 Но пока некогда разбираться с кодировками
   }
   catch(IOException e) {
    //Console.WriteLine("Обработано исключение {0}", e); Блядь, тут тоже руки не дошли. Нужен подробный разбор вариантов исключения.
    System.IO.File.Copy(filepath, "temp.dbf", true);
    filepath="temp.dbf";
    text = System.IO.File.ReadAllText(filepath, Encoding.GetEncoding(1251)); //В общем, надо ставить 1251, если windows 8 или что-то подходящее. Ещё надо смотреть, как БД кодирована.
   }
   DBFheader=GetDBFRecordsCount(filepath);
   long RecordLen=0;
   long CodeFieldStartPosition=0;
   long CodeFieldLength=0;
   foreach (DBFFieldDescription field in DBFheader.FieldsProperties) {
    //Console.WriteLine("Field: {0}, type: {1}, len: {2}", field.NameOfField, field.TypeOfField, field.FieldLength);
    if (field.NameOfField=="CODE") {
     CodeFieldStartPosition=RecordLen;
     CodeFieldLength=field.FieldLength;
     break;
    }
    else {
     RecordLen+=field.FieldLength;
    }
   }
   Match m = Regex.Match(text, pattern);
   ArrayList myMatches=new ArrayList(); //Только коды элементов возвращаются

   while (m.Success) {
    int inFilePosition=(int)(m.Index-DBFheader.HeaderSize)/DBFheader.RecordSize;
    //Console.WriteLine("{0} -> {1}", CodeFieldLength, (int)CodeFieldLength);
    string Code=text.Substring((int)(CodeFieldStartPosition+DBFheader.HeaderSize+(inFilePosition)*DBFheader.RecordSize)+1, (int)CodeFieldLength);
    myMatches.Add(Code);
    Console.WriteLine("'{0}' найдено на позиции {1} raw, Код: '{2}'", m.Value, inFilePosition, Code);
    m = m.NextMatch();
   }   
   Console.WriteLine("Всего в БД {0} записей, каждая по {3} полей, Размер заголовка: {2}, Размер записи: {1} Нажми любую кнопку блеать!", DBFheader.RecordsCount, DBFheader.RecordSize, DBFheader.HeaderSize, DBFheader.FieldCount);
   System.Console.ReadKey();
   if (filepath=="temp.dbf") {
    System.IO.File.Delete(filepath);    
   }
   return (myMatches);
  }


  /************************************************************************************************************************************/
  // Функция возвращает количество записей в таблице. Чтобы избежать проблем с переполнением - возвращает long - 64 битное число
  // Так же возвращает структуру таблицы в массиве {NameOfField, TypeOfField, FieldLength}
  // 
  static DBFProperties GetDBFRecordsCount(string filepath) {
   const int HEADERSIZE = 32;
   DBFProperties thisDBFprops;
   byte[] header = new byte[HEADERSIZE];
   FileInfo f = new FileInfo(filepath);
   thisDBFprops.FileSize=f.Length;
   using (StreamReader reader = new StreamReader(filepath)) {
    using (BinaryReader binaryReader = new BinaryReader(reader.BaseStream)) {
     binaryReader.Read(header, 0, HEADERSIZE);
     thisDBFprops.RecordsCount=header[7]*(long)Math.Pow(256, 3)+header[6]*(long)Math.Pow(256, 2)+header[5]*256+header[4];
     thisDBFprops.HeaderSize=header[9]*256+header[8];
     thisDBFprops.RecordSize=header[11]*256+header[10];
    }
   }
   thisDBFprops.FieldCount=(int)(thisDBFprops.FileSize-thisDBFprops.RecordSize*thisDBFprops.RecordsCount-32)/32;
   thisDBFprops.FieldsProperties = new ArrayList();
   byte[] allheader = new byte[HEADERSIZE*(1+thisDBFprops.FieldCount)];
   using (StreamReader reader = new StreamReader(filepath)) {
    using (BinaryReader binaryReader = new BinaryReader(reader.BaseStream)) {
     binaryReader.Read(allheader, 0, HEADERSIZE*(1+thisDBFprops.FieldCount));
 
     // Немножко говнокода, к сожалению for(;;) не знаю пока как заменить на итератор
     for (int i=1; i<=thisDBFprops.FieldCount; i++) {
      string strFieldName="";
      for (long j=0; j<=10; j++) {
       if (allheader[j+i*HEADERSIZE]==0) break;
       strFieldName+=(char)allheader[j+i*HEADERSIZE];
      }
      DBFFieldDescription fieldprops;
      fieldprops.NameOfField=strFieldName;
      fieldprops.TypeOfField=allheader[11+i*HEADERSIZE];
      fieldprops.FieldLength=allheader[16+i*HEADERSIZE];
      thisDBFprops.FieldsProperties.Add(fieldprops);
     }
    }
   }
  return (thisDBFprops);
  }
 }
}

