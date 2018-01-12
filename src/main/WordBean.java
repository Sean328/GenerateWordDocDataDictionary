package main;

import java.util.List;

import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.Dispatch;
import com.jacob.com.Variant;    
public class WordBean {    
   // 代表一个word 程序    
   private ActiveXComponent MsWordApp = null;    
   // 代表进行处理的word 文档    
   private Dispatch document = null;    
   public WordBean() {    
       // Open Word if we\'ve not done it already    
       if (MsWordApp == null) {    
           MsWordApp = new ActiveXComponent("Word.Application");    
       }    
   }    
   // 设置是否在前台打开 word程序 ，    
   public void setVisible(boolean visible) {    
       MsWordApp.setProperty("Visible", new Variant(visible));    
       // 这一句作用相同    
       // Dispatch.put(MsWordApp, "Visible", newVariant(visible));    
   }    
   // 创建一个新文档    
   public void createNewDocument() {    
       // Find the Documents collection object maintained by Word    
       // documents表示word的所有文档窗口，（word是多文档应用程序）    
       Dispatch documents = Dispatch.get(MsWordApp,"Documents").toDispatch();    
       // Call the Add method of the Documents collection to create    
       // a new document to edit    
       document = Dispatch.call(documents, "Add").toDispatch();    
   }    
   // 打开一个存在的word文档,并用document 引用 引用它    
   public void openFile(String wordFilePath) {    
       // Find the Documents collection object maintained by Word    
       // documents表示word的所有文档窗口，（word是多文档应用程序）    
       Dispatch documents = Dispatch.get(MsWordApp,"Documents").toDispatch();    
       document = Dispatch.call(documents, "Open", wordFilePath,    
                new Variant(true)/* 是否进行转换ConfirmConversions */,    
                new Variant(false)/* 是否只读 */).toDispatch();    
       // document = Dispatch.invoke(documents, "Open",Dispatch.Method,    
       // new Object[] { wordFilePath, new Variant(true),    
       // new Variant(false)    
       // }, new int[1]).toDispatch();    
    }    
   // 向 document 中插入文本内容    
   public void insertText(String textToInsert) {    
       // Get the current selection within Word at the moment.    
       // a new document has just been created then this will be at    
       // the top of the new doc 获得选 中的内容，如果是一个新创建的文件，因里面无内容，则光标应处于文件开头处    
       Dispatch selection = Dispatch.get(MsWordApp,"Selection").toDispatch();    
       // 取消选中,应该就是移动光标 ，否则 新添加的内容会覆盖选中的内容    
       Dispatch.call(selection, "MoveRight", new Variant(1), new Variant(1));    
       // Put the specified text at the insertion point    
       Dispatch.put(selection, "Text", textToInsert);    
       // 取消选中,应该就是移动光标    
       Dispatch.call(selection, "MoveRight", new Variant(1), new Variant(1));    
   }    
   // 向文档中添加 一个图片，    
   public void insertJpeg(String jpegFilePath) {    
       Dispatch selection = Dispatch.get(MsWordApp,"Selection").toDispatch();    
       Dispatch image = Dispatch.get(selection,"InLineShapes").toDispatch();    
       Dispatch.call(image, "AddPicture", jpegFilePath);    
   }    
   // 段落的处理,插入格式化的文本    
   public void insertFormatStr(String text) {    
       Dispatch wordContent = Dispatch.get(document,"Content").toDispatch(); // 取得word文件的内容    
       Dispatch.call(wordContent, "InsertAfter", text);// 插入一个段落到最后    
       Dispatch paragraphs = Dispatch.get(wordContent,"Paragraphs")    
                .toDispatch(); // 所有段落    
       int paragraphCount = Dispatch.get(paragraphs,"Count").changeType(    
                Variant.VariantInt).getInt();//一共的段落数    
       // 找到刚输入的段落，设置格式    
       Dispatch lastParagraph = Dispatch.call(paragraphs,"Item",    
                new Variant(paragraphCount)).toDispatch(); // 最后一段（也就是刚插入的）    
       // Range 对象表示文档中的一个连续范围，由一个起始字符位置和一个终止字符位置定义    
       Dispatch lastParagraphRange = Dispatch.get(lastParagraph,"Range")    
                .toDispatch();    
       Dispatch font = Dispatch.get(lastParagraphRange,"Font").toDispatch();    
       Dispatch.put(font, "Bold", new Variant(true)); // 设置为黑体    
       Dispatch.put(font, "Italic", new Variant(true)); // 设置为斜体    
       Dispatch.put(font, "Name", new Variant("宋体")); //    
       Dispatch.put(font, "Size", new Variant(12)); // 小四    
       Dispatch selection = Dispatch.get(MsWordApp, "Selection").toDispatch();    
       Dispatch.call(selection, "TypeParagraph");// 插入一个空行    
       Dispatch alignment = Dispatch.get(selection,"ParagraphFormat")    
                .toDispatch();// 段落格式    
       Dispatch.put(alignment, "Alignment", "2"); // (1:置中 2:靠右 3:靠左)    
   }    
   // word 中在对表格进行遍历的时候 ，是先列后行 先column 后cell    
   // 另外下标从1开始    
   public void insertTable(String tableTitle, int row, int column, List<List<String>> detail) {    
       Dispatch selection = Dispatch.get(MsWordApp,"Selection").toDispatch(); // 输入内容需要的对象    
       Dispatch.call(selection, "TypeText", tableTitle); // 写入标题内容 // 标题格行    
       
       Dispatch.call(selection, "TypeParagraph"); // 空一行段落    
       Dispatch.call(selection, "TypeParagraph"); // 空一行段落    
       Dispatch.call(selection, "MoveDown"); // 游标往下一行    
       // 建立表格    
       Dispatch tables = Dispatch.get(document,"Tables").toDispatch();    
       // int count = Dispatch.get(tables,    
       // "Count").changeType(Variant.VariantInt).getInt(); //document中的表格数量    
       // Dispatch table = Dispatch.call(tables, "Item", newVariant(    
       // 1)).toDispatch();//文档中第一个表格    
       Dispatch range = Dispatch.get(selection,"Range").toDispatch();// /当前光标位置或者选中的区域    
       Dispatch newTable = Dispatch.call(tables, "Add", range,    
                new Variant(row), new Variant(column), new Variant(1))    
                .toDispatch(); // 设置row,column,表格外框宽度    
       Dispatch cols = Dispatch.get(newTable,"Columns").toDispatch(); // 此表的所有列，    
       int colCount = Dispatch.get(cols, "Count").changeType(    
                Variant.VariantInt).getInt();//一共有多少列 实际上这个数==column    
       System.out.println(colCount + "列"+" "+row+"行");    
       for (int i = 1; i <= colCount; i++) { // 循环取出每一列    
           Dispatch col = Dispatch.call(cols, "Item", new Variant(i))    
                    .toDispatch();    
           Dispatch cells = Dispatch.get(col, "Cells").toDispatch();// 当前列中单元格    
           int cellCount = Dispatch.get(cells, "Count").changeType(    
                    Variant.VariantInt).getInt();//当前列中单元格数 实际上这个数等于row  
           for (int j = 1; j <= cellCount; j++) {// 每一列中的单元格数    
                // Dispatch cell =Dispatch.call(cells, "Item", new    
                // Variant(j)).toDispatch(); //当前单元格    
                // Dispatch cell =Dispatch.call(newTable, "Cell", new    
                // Variant(j) , new Variant(i)).toDispatch(); //取单元格的另一种方法    
                // Dispatch.call(cell,"Select");//选中当前单元格    
                // Dispatch.put(selection,"Text",    
               // "第"+j+"行，第"+i+"列");//往选中的区域中填值，也就是往当前单元格填值 
               if(!(null==detail.get(j-1).get(i-1))){
                   putTxtToCell(newTable, j, i,detail.get(j-1).get(i-1));// 与上面四句的作用相同 
               }
           }    
       }
       Dispatch.call(selection, "MoveDown"); // 游标往下一行    
       Dispatch.call(selection, "MoveDown"); // 游标往下一行    
       Dispatch.call(selection, "TypeParagraph"); // 空一行段落  
       Dispatch.call(selection, "MoveDown"); // 游标往下一行    
   }    
   
   /**  
    * 在指定的单元格里填写数据  
    *  
    * @param tableIndex  
    * @param cellRowIdx  
    * @param cellColIdx  
    * @param txt  
    */   
   public void putTxtToCell(Dispatch table, int cellRowIdx, int cellColIdx,    
           String txt) {    
       Dispatch cell = Dispatch.call(table, "Cell", new Variant(cellRowIdx),    
                new Variant(cellColIdx)).toDispatch();    
       Dispatch.call(cell, "Select");    
       Dispatch selection = Dispatch.get(MsWordApp,"Selection").toDispatch(); // 输入内容需要的对象    
       Dispatch.put(selection, "Text", txt);    
   }    
   /** */   
   /**  
    * 在指定的单元格里填写数据  
    *  
    * @param tableIndex  
    * @param cellRowIdx  
    * @param cellColIdx  
    * @param txt  
    */   
   public void putTxtToCell(int tableIndex, int cellRowIdx, int cellColIdx,    
           String txt) {    
       // 所有表格    
       Dispatch tables = Dispatch.get(document,"Tables").toDispatch();    
       // 要填充的表格    
       Dispatch table = Dispatch.call(tables, "Item", new Variant(tableIndex))    
                .toDispatch();    
       Dispatch cell = Dispatch.call(table, "Cell", new Variant(cellRowIdx),    
                new Variant(cellColIdx)).toDispatch();    
       Dispatch.call(cell, "Select");    
       Dispatch selection = Dispatch.get(MsWordApp,"Selection").toDispatch(); // 输入内容需要的对象    
       Dispatch.put(selection, "Text", txt);    
   }    
   // 合并两个单元格    
   public void mergeCell(Dispatch cell1, Dispatch cell2) {    
       Dispatch.call(cell1, "Merge", cell2);    
   }    
   public void mergeCell(Dispatch table, int row1, int col1, int row2, int col2) {    
       Dispatch cell1 = Dispatch.call(table, "Cell", new Variant(row1),    
                new Variant(col1)).toDispatch();    
       Dispatch cell2 = Dispatch.call(table, "Cell", new Variant(row2),    
               new Variant(col2)).toDispatch();    
       mergeCell(cell1, cell2);    
   }    
   public void mergeCellTest() {    
       Dispatch tables = Dispatch.get(document,"Tables").toDispatch();    
       int tableCount = Dispatch.get(tables, "Count").changeType(    
                Variant.VariantInt).getInt();// document中的表格数量    
       Dispatch table = Dispatch.call(tables, "Item", new Variant(tableCount))    
                .toDispatch();// 文档中最后一个table    
       mergeCell(table, 1, 1, 1, 2);// 将table 中x=1,y=1 与x=1,y=2的两个单元格合并    
   }    
   
   /**  
    * 把选定的内容或光标插入点向上移动  
    *  
    * @param pos  
    *            移动的距离  
    */   
   public void moveUp(int pos) {     
       Dispatch selection = Dispatch.get(MsWordApp,"Selection").toDispatch(); // 输入内容需要的对象    
       for (int i = 0; i < pos; i++) {    
           // MoveDown MoveLeft moveRight    
           // moveStart ( Dispatch.call(selection, "HomeKey", new Variant(6));    
           // )    
           // moveEnd Dispatch.call(selection, "EndKey", newVariant(6));    
           Dispatch.call(selection, "MoveUp");    
       }    
   }    
   /** */   
   /**  
    * 从选定内容或插入点开始查找文本  
    *  
    * @param toFindText  
    *            要查找的文本  
    * @return boolean true-查找到并选中该文本，false-未查找到文本  
    */   
   public boolean find(String toFindText) {    
       if (toFindText == null || toFindText.equals(""))    
           return false;    
       Dispatch selection = Dispatch.get(MsWordApp,"Selection").toDispatch(); // 输入内容需要的对象    
       // 从selection所在位置开始查询    
       Dispatch find = Dispatch.call(selection,"Find").toDispatch();    
       // 设置要查找的内容    
       Dispatch.put(find, "Text", toFindText);    
       // 向前查找    
       Dispatch.put(find, "Forward", "True");    
       // 设置格式    
       Dispatch.put(find, "Format", "True");    
       // 大小写匹配    
       Dispatch.put(find, "MatchCase", "True");    
       // 全字匹配    
       Dispatch.put(find, "MatchWholeWord", "True");    
       // 查找并选中    
       return Dispatch.call(find, "Execute").getBoolean();    
   }    
   /** */   
   /**  
    * 把选定选定内容设定为替换文本  
    *  
    * @param toFindText  
    *            查找字符串  
    * @param newText  
    *            要替换的内容  
    * @return  
    */   
   public boolean replaceText(String toFindText, String newText) {    
       if (!find(toFindText))    
           return false;    
       Dispatch selection = Dispatch.get(MsWordApp, "Selection").toDispatch();// 输入内容需要的对象    
       Dispatch.put(selection, "Text", newText);    
       return true;    
   }    
   public void printFile() {    
       // Just print the current document to the default printer    
       Dispatch.call(document, "PrintOut");    
   }    
   // 保存文档的更改    
   public void save() {    
       Dispatch.call(document, "Save");    
   }    
   public void saveFileAs(String filename) {    
       Dispatch.call(document, "SaveAs", filename);    
   }    
   public void closeDocument() {    
       // Close the document without saving changes    
       // 0 = wdDoNotSaveChanges    
       // -1 = wdSaveChanges    
       // -2 = wdPromptToSaveChanges    
       Dispatch.call(document, "Close", new Variant(0));    
       document = null;    
   }    
   public void closeWord() {    
       Dispatch.call(MsWordApp, "Quit");    
       MsWordApp = null;    
       document = null;    
   }    
   // 设置wordApp打开后窗口的位置    
   public void setLocation() {    
       Dispatch activeWindow = Dispatch.get(MsWordApp,"Application")    
                .toDispatch();    
       Dispatch.put(activeWindow, "WindowState", new Variant(1)); //0=default    
       // 1=maximize    
       // 2=minimize    
       Dispatch.put(activeWindow, "Top", new Variant(0));    
       Dispatch.put(activeWindow, "Left", new Variant(0));    
       Dispatch.put(activeWindow, "Height", new Variant(600));    
       Dispatch.put(activeWindow, "width", new Variant(800));    
   }  
   
   /** 
    * 设置当前选定内容的字体 
    * 
    * @param boldSize 
    * @param italicSize 
    * @param underLineSize 下划线 
    * @param colorSize  字体颜色 
    * @param size 字体大小 
    * @param name 字体名称 
    */  
   public void setFont(boolean bold, boolean italic, boolean underLine,  
       String colorSize, String size, String name,Dispatch selection) {  
       Dispatch font = Dispatch.get(selection, "Font").toDispatch();  
       Dispatch.put(font, "Name", new Variant(name));  
       Dispatch.put(font, "Bold", new Variant(bold));  
       Dispatch.put(font, "Italic", new Variant(italic));  
       Dispatch.put(font, "Underline", new Variant(underLine));  
       Dispatch.put(font, "Color", colorSize);  
       Dispatch.put(font, "Size", size);  
   }
}