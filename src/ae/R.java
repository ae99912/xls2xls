package ae;

//import java.io.*;
//import java.util.Properties;

/*
 Ресурсный класс

 Modify:
 13.07.23 определил карту переноса в виде текстового файла
 14.07.23 запись Excel через временный файл
 15.07.23 уникальность ячеек в наборе при чтении карты (yach.hashCode)
 15.07.23 задание диапазона ячеек (например, C7:F11)
 16.07.23 разбор строки с ячейками перенес в класс karta
 16.07.23 убрал класс work, работаем в Main
 17.07.23 при задании диапазона правая граница д.б. больше левой
 24.10.23 ввел свойство only01 - заносить только 0 или 1 в ячейку, свойство в карте - строка "@only01"
 09.11.23 можно указать номер листа
 04.12.23 обычные сообщения только в режиме отладки
 27.07.24 при копировании ячейки с формулой вставляется значение формулы
 29.07.24 ячейка с типом BLANK не изменяет выходную ячейку
 22.08.24 в карте переноса в каждой ячейке хранится свое свойство,в карте может быть несколько
          свойств, заданных для последующих ячеек.
          для обработки свойства задействовал регулярное выражение, теперь есть свойства:
            only01  - в ячейке только 0 или 1
            only-01 - в ячейке только - или 0 или 1
            regex   - указывается конкретное регулярное выражение проверки соответствия значения
 23.08.24 модификация регулярных выражений, добавил новые свойства:
            onlyint - в ячейке целое число
            onlynum - в ячейке любое число
 24.08.24 если в Excel нет нужной строки, то создаем ее
 04.09.24 функция getCellStrValue если формула, возвращает значение
 16.09.25 свойство ячейки @blank, когда требуется очистка выходной ячейки
          добавил функцию setCellBlank

*/

public class R {
    public final static String Ver = "Ver. 5.6"; // номер версии
    
    //final static String sep = System.getProperty("file.separator"); // разделитель имени каталогов

    // свойство очистки выходной ячейки
    public final static String PAT_BLANK = "blank";
    static boolean debug = false;   // отладочный вывод

    /**
     * Вывести сообщение, если задан режим отладки
     * @param message - сообщение
     */
    static public void out(String message)
    {
        if(debug) {
            System.out.println(message);
        }
    }

//    /**
//     * загрузка значений параметров по-умолчанию из файла res/default.properties
//     * Порядок определения каталогов:
//     */
//    void loadDefault()
//    {
//        // http://stackoverflow.com/questions/2815404/load-properties-file-in-jar
//        // Отобразим версию
//        System.out.println(Ver);
//        Properties props = new Properties();
//        try {
//            props.load(R.class.getResourceAsStream("res/default.properties"));
//            // прочитаем параметры из конфигурационного файла default.properties
//            workDir = r2s(props, "workDir", workDir);
//            if(workDir == null) {
//                workDir = System.getProperty("java.io.tmpdir", ".");
//            }
//            // колонки с числами
////            intIndex = r2s(props, "intIndex", intIndex);
////            dblIndex = r2s(props, "dblIndex", dblIndex);
//            //
//        } catch (IOException e) {
//            e.printStackTrace();
//        }
//    }

//    /**
//     * Выдать строковое значение из файла свойств, либо, если там
//     * нет такого свойства, вернуть значение по-умолчанию
//     * @param p                     свойства
//     * @param NameProp              имя свойства
//     * @param strResourceDefault    значение по-умолчанию
//     * @return  значение свойства, а если его нет, то значение по-умолчанию
//     */
//    private String r2s(Properties p, String NameProp, String strResourceDefault)
//    {
//        String str = p.getProperty(NameProp);
//        if(str == null) {
//            str = strResourceDefault;
//        }
//        return str;
//    }

//    /**
//     * Выдать числовое (long) значение из файла свойств, либо, если там
//     * нет такого свойства, вернуть значение по-умолчанию
//     * @param p                     свойства
//     * @param NameProp              имя свойства
//     * @param lngResourceDefault    значение по-умолчанию
//     * @return  значение свойства, а если его нет, то значение по-умолчанию
//     */
//    private long r2s(Properties p, String NameProp, long lngResourceDefault)
//    {
//        String str = p.getProperty(NameProp);
//        if(str == null) {
//            str = String.valueOf(lngResourceDefault);
//        }
//        return Long.parseLong(str);
//    }

//    /**
//     * Выдать числовое (int) значение из файла свойств, либо, если там
//     * нет такого свойства, вернуть значение по-умолчанию
//     * @param p                     свойства
//     * @param NameProp              имя свойства
//     * @param intResourceDefault    значение по-умолчанию
//     * @return  значение свойства, а если его нет, то значение по-умолчанию
//     */
//    private int r2s(Properties p, String NameProp, int intResourceDefault)
//    {
//        String str = p.getProperty(NameProp);
//        if(str == null) {
//            str = String.valueOf(intResourceDefault);
//        }
//        return Integer.parseInt(str);
//    }

//    /**
//     * прочитать ресурсный файл
//     * by novel  http://skipy-ru.livejournal.com/5343.html
//     * https://docs.oracle.com/javase/tutorial/deployment/webstart/retrievingResources.html
//     * @param nameRes - имя ресурсного файла
//     * @return - содержимое ресурсного файла
//     */
//    public String readRes(String nameRes)
//    {
//        String str = null;
//        ByteArrayOutputStream buf = readResB(nameRes);
//        if(buf != null) {
//            str = buf.toString();
//        }
//        return str;
//    }

//    /**
//     * Поместить ресурс в байтовый массив
//     * @param nameRes - название ресурса (относительно каталога пакета)
//     * @return - байтовый массив
//     */
//    private ByteArrayOutputStream readResB(String nameRes)
//    {
//        try {
//            // Get current classloader
//            InputStream is = getClass().getResourceAsStream(nameRes);
//            if(is == null) {
//                System.out.println("Not found resource: " + nameRes);
//                return null;
//            }
//            // https://habrahabr.ru/company/luxoft/blog/278233/ п.8
//            BufferedInputStream bin = new BufferedInputStream(is);
//            ByteArrayOutputStream bout = new ByteArrayOutputStream();
//            int len;
//            byte[] buf = new byte[512];
//            while((len=bin.read(buf)) != -1) {
//                bout.write(buf,0,len);
//            }
//            return bout;
//        } catch (IOException ex) {
//            ex.printStackTrace();
//        }
//        return null;
//    }

//    /**
//     * Записать в файл текст из строки
//     * @param strTxt - строка текста
//     * @param fileName - имя файла
//     * @return      true - записано, false - ошибка
//     */
//    public boolean writeStr2File(String strTxt, String fileName)
//    {
//        File f = new File(fileName);
//        try {
//            // сформируем командный файл BAT
//            PrintWriter out = new PrintWriter(f);
//            out.write(strTxt);
//            out.close();
//        } catch(IOException ex) {
//            ex.printStackTrace();
//            return false;
//        }
//        return true;
//    }

//    /**
//     *  Записать в файл ресурсный файл
//     * @param nameRes   имя ресурса (от корня src)
//     * @param fileName  имя файла, куда записывается ресурс
//     * @return  true - запись выполнена, false - ошибка
//     */
//    boolean writeRes2File(String nameRes, String fileName)
//    {
//        boolean b = false;
//        ByteArrayOutputStream buf = readResB(nameRes);
//        if(buf != null) {
//            try {
//                FileOutputStream fout = new FileOutputStream(fileName);
//                buf.writeTo(fout);
//                fout.close();
//                b = true;
//            } catch (IOException e) {
//                e.printStackTrace();
//            }
//        }
//        return b;
//    }
    
//    /**
//     * Загружает текстовый ресурс в заданной кодировке
//     * @param name      имя ресурса
//     * @param code_page кодировка, например "Cp1251"
//     * @return          строка ресурса
//     */
//    public String getText(String name, String code_page)
//    {
//        StringBuilder sb = new StringBuilder();
//        try {
//            InputStream is = this.getClass().getResourceAsStream(name);  // Имя ресурса
//            BufferedReader br = new BufferedReader(new InputStreamReader(is, code_page));
//            String line;
//            while ((line = br.readLine()) !=null) {
//                sb.append(line);  sb.append("\n");
//            }
//        } catch (IOException ex) {
//            ex.printStackTrace();
//        }
//        return sb.toString();
//    }
    
//    /**
//     * Пауза выполнения программы (потока)
//     * @param msec - задержка, мсек
//     */
//    public static void Sleep(long msec)
//    {
//        try {
//            Thread.sleep(msec);
//        } catch (InterruptedException e) {
//            e.printStackTrace();
//        }
//    }

} // end of class
