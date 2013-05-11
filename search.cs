using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.IO;
using System.Numerics;
using System.Collections;
using System.EnterpriseServices;
using System.Reflection;
using System.Runtime.InteropServices;

namespace AddIn { //
//namespace SlonProgs {
/*[assembly: ComVisible(true)]

[assembly: Guid("d28c6322-2c2a-4763-856e-5b379cd18202")]*/
 public class InDBFSearch {

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
   //string pattern=@"((I|i|И|и)(N|n|Н|н)(T|t|Т|т)(E|e|Е|е)(L|l|Л|л)|(A|a|А|а)(M|m|М|м)(D|d|Д|д))";
   string pattern=@"(К|к|C|c)(О|о|O|o)(М|м|M|m)(П|п|P|p)(Ь|ь|U|u)(Ю|ю|T|t)(Т|т|E|e)";
   string resultfilepath="result.txt";
   //string pattern=@"(Н|н|H)\s*(А|а|A|a)\s*(Р|р|P|p)\s*(K|k|К|к)\s*(О|о|O|o)[^((М|м|M)(Ы|ы|O|o|О|о|У|у|Y|y|C|c|С|с))]\S*";
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
    //Console.WriteLine("Ищем наркотики ({0}) в файле {1}", pattern, filepath);
    
   }
   if (args.Length==3) {
    filepath=args[0];
    //Console.WriteLine("На входе 3 параметра \n\rfilepath:|{0}|\n\rregexp:|{1}|\n\rOutput to:|{2}|", args[0], args[1], args[2]);
    pattern=args[1];
    resultfilepath=args[2];
   }
   if (File.Exists(filepath)) {
    //ArrayList Matches=searchInDBF(filepath, pattern);
    searchInDBF(filepath, pattern, resultfilepath);
   }
   else { 
    //Console.WriteLine("Файл {0} не найден.", filepath);
   }
   //System.Console.ReadKey();
  }

  // ************************************************************************************************************************************************
  // Функция поиска
  //
  static void searchInDBF(string filepath, string pattern, string resultfilepath) {
   string text="";
   DBFProperties DBFheader;
   FileInfo f = new FileInfo(filepath);
   int thisFileSize=(int)f.Length;
   byte[] fileInBytes = new byte[thisFileSize];
   try {
    //Console.WriteLine("Попытка создания FileStream.");
    using (FileStream fs = new FileStream(filepath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)) {
     //Console.WriteLine("FileStream создан, попытка чтения");
     Encoding temp = Encoding.GetEncoding("windows-1251");; //1Cv77 не работает с юникодом, 
     //Encoding temp = new Encoding(0);
     if (fs.Read(fileInBytes, 0, thisFileSize) == thisFileSize) {
      //Console.WriteLine("попытка чтения успешна");
      text=temp.GetString(fileInBytes);
      //Console.WriteLine("{0}", text.Substring(1000,200));
     }
     else {
      using (StreamWriter writer = File.CreateText("result.txt")) {
       //Console.WriteLine("Из потока считано меньше байт, чем размер файла {0}", filepath);
       writer.WriteLineAsync("<filename>"+filepath+"</filename>");
       writer.WriteLineAsync("<exception>File read uncomplete.</exception>");
      }
      return;
     }
    }
   }
   catch {//(IOException e) {
    //Console.WriteLine("Обработано исключение {0}", e);
    try {
     System.IO.File.Copy(filepath, "temp.dbf", true);
    }
    catch {//(IOException e1) {
     using (StreamWriter writer = File.CreateText("result.txt")) {
      //Console.WriteLine("Обработано исключение {0}", e1);
      writer.WriteLineAsync("<filename>"+filepath+"</filename>");
      writer.WriteLineAsync("<exception>FileBlocked</exception>");
     }
     return;
    }
    filepath="temp.dbf";
    text = System.IO.File.ReadAllText(filepath, Encoding.GetEncoding(1251)); //В общем, надо ставить 1251, если windows 8 или что-то подходящее. Ещё надо смотреть, как БД кодирована.
   }
   try {
    //Console.WriteLine("Попытка чтения заголовка");
    DBFheader=GetDBFRecordsCount(fileInBytes);
   }
   catch {//(IOException e2){
    using (StreamWriter writer = File.CreateText("result.txt")) {
     //Console.WriteLine("Обработано исключение {0}", e2);
     writer.WriteLineAsync("<filename>"+filepath+"</filename>");
     writer.WriteLineAsync("<exception>Can't read file header</exception>");
    }
    return;
   }
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
   using (StreamWriter writer = File.CreateText(resultfilepath)) {
    writer.WriteLineAsync("<filename>"+filepath+"</filename>");
    writer.WriteLineAsync("<regexp>"+pattern+"</regexp>");
    while (m.Success) {
     int inFilePosition=(int)(m.Index-DBFheader.HeaderSize)/DBFheader.RecordSize;
     //Console.WriteLine("{0} -> {1}", CodeFieldLength, (int)CodeFieldLength);
     string Code=text.Substring((int)(CodeFieldStartPosition+DBFheader.HeaderSize+(inFilePosition)*DBFheader.RecordSize)+1, (int)CodeFieldLength);
     myMatches.Add(Code);
     //Console.WriteLine("'{0}' найдено на позиции {1} raw, Код: '{2}'", m.Value, inFilePosition, Code);
     writer.WriteLine("<code>"+Code+"</code>");
     m = m.NextMatch();
    }   
   }
   //Console.WriteLine("Всего в БД {0} записей, каждая по {3} полей, Размер заголовка: {2}, Размер записи: {1} Нажми любую кнопку блеать!", DBFheader.RecordsCount, DBFheader.RecordSize, DBFheader.HeaderSize, DBFheader.FieldCount);
   //System.Console.ReadKey();
/*   if (filepath=="temp.dbf") {
    System.IO.File.Delete(filepath);    
   }*/
   //return (myMatches);
  }


  /************************************************************************************************************************************/
  // Функция возвращает количество записей в таблице. Чтобы избежать проблем с переполнением - возвращает long - 64 битное число
  // Так же возвращает структуру таблицы в массиве {NameOfField, TypeOfField, FieldLength}
  // 
  static DBFProperties GetDBFRecordsCount(byte[] header) {
   const int HEADERSIZE = 32;
   DBFProperties thisDBFprops;
   thisDBFprops.FileSize=header.Length;
   thisDBFprops.RecordsCount=header[7]*(long)Math.Pow(256, 3)+header[6]*(long)Math.Pow(256, 2)+header[5]*256+header[4];
   thisDBFprops.HeaderSize=header[9]*256+header[8];
   thisDBFprops.RecordSize=header[11]*256+header[10];
   thisDBFprops.FieldCount=(int)(thisDBFprops.FileSize-thisDBFprops.RecordSize*thisDBFprops.RecordsCount-32)/32;
   thisDBFprops.FieldsProperties = new ArrayList();
   for (int i=1; i<=thisDBFprops.FieldCount; i++) {
    string strFieldName="";
    for (long j=0; j<=10; j++) {
     if (header[j+i*HEADERSIZE]==0) break;
     strFieldName+=(char)header[j+i*HEADERSIZE];
    }
    DBFFieldDescription fieldprops;
    fieldprops.NameOfField=strFieldName;
    fieldprops.TypeOfField=header[11+i*HEADERSIZE];
    fieldprops.FieldLength=header[16+i*HEADERSIZE];
    thisDBFprops.FieldsProperties.Add(fieldprops);
   }
  return (thisDBFprops);
  }
 }
}

