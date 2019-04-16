package com.chinadaas.riskbell;

import org.apache.poi.POIXMLDocumentPart;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.NumberToTextConverter;
import org.apache.poi.xssf.usermodel.*;
import org.openxmlformats.schemas.drawingml.x2006.spreadsheetDrawing.CTMarker;

import java.io.*;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.*;

/**
 * Created by pc on 2017/3/13.
 * @para 输入路径（文件名），输出路径（文件夹名）
 * 完成文件的对应输出
 *
 */
public class Els_x_toHtml {

    private final Workbook wb;
    private HSSFWorkbook hwb;
    private XSSFWorkbook xwb;
    private Appendable output;
    private Formatter out;
    private int sheetNo;
    private String sheetNames[];
    private String outputPath;
    private String inputPath;

    public Els_x_toHtml(String inputPath,String outputPath) throws IOException{
       this.outputPath=outputPath;
       this.inputPath=inputPath;
            try {
                wb = WorkbookFactory.create(new FileInputStream(inputPath));
                output= new PrintWriter(new FileWriter(outputPath));
                if (wb == null)
                    throw new NullPointerException("wb");
                if (output == null)
                    throw new NullPointerException("output");
                sheetNo = wb.getNumberOfSheets();
                getNames(sheetNo);
                printPage();
            } catch (Exception e){
                throw new IllegalArgumentException("Cannot create workbook from stream", e);
            }
        }
    private void constructTabHeader()
    {
        for(int i=0;i<sheetNo;i++)
        {
        }
    }

    private void getNames(int num)
    {
        sheetNames = new String[num];
        for(int i=0;i<num;i++)
        {
            String name = wb.getSheetName(i);
            sheetNames[i]=name;
        }
    }

    //判断Excel版本
    private void printsheets(StringBuilder outHtml) throws IOException{
        Map<String, PictureData>  maplist=null;
        if (wb instanceof HSSFWorkbook)
        {
            hwb=new HSSFWorkbook(new FileInputStream(inputPath));
            int sheetnum=hwb.getNumberOfSheets();
//            for(int nums=0;nums<sheetnum;nums++){

                int nums = 0;

                HSSFSheet sheet = hwb.getSheetAt(nums);
                //获取图片并导出
                maplist = getPicturesxls((HSSFSheet) sheet);

                //打印图片
                printImg(maplist);

                Map<Integer,Integer> xyArray = new HashMap<Integer,
                        Integer>();

                //获取图片单元格坐标
                for (Map.Entry<String, PictureData> entry :
                        maplist.entrySet()) {
                    System.out.println("Key = " + entry.getKey() + ", Value = " + entry.getValue());

                    //定义x坐标
                    xyArray.put(Integer.parseInt(entry.getKey().split("-")[0]),
                            Integer.parseInt(entry.getKey().split("-")[1]));
                }





                String name = sheet.getSheetName();
                int numr =sheet.getPhysicalNumberOfRows();
                int numc =sheet.getLeftCol();

                for(int i=0;i<numr;i++){
                    outHtml.append("<table ><tr>\n");

                    int y=0;

                    //如果包含x坐标
                    boolean result = xyArray.containsKey(i);
                    if(result){
                        y=xyArray.get(i);
                    }

                    HSSFRow hssfRow = sheet.getRow(i);
                    numc=hssfRow.getPhysicalNumberOfCells();
                    for(int j=0;j<numc;j++){
                        if(result && j==y){
                            //获取图片后缀名
                            PictureData picture = maplist.get((i)+
                                    "-"+y);

                            String fileName =
                                    picture.getMimeType().substring(6);

                            outHtml.append("<td><img src=\"./"+i+"-"+y+
                                    "."+fileName+
                                    "\" /></td" +
                                    ">\n" );

                            //移除表格内图片
                            maplist.remove((i)+
                                    "-"+y);

                        }else{
                            HSSFCell cell =hssfRow.getCell(j);
                            String str =getCellValue(cell);
                            outHtml.append("<td>"+str+"</td>\n" );
                        }




                    }
                    outHtml.append("</tr>\n</table>");
                }
            for(Map.Entry<String, PictureData> map : maplist.entrySet()){
                outHtml.append("<table ><tr>\n");

                PictureData picture = maplist.get(map.getKey());

                String fileName =
                        picture.getMimeType().substring(6);
                outHtml.append("<td><img src=\"./"+map.getKey()+
                        "."+fileName+
                        "\" /></td" +
                        ">\n" );
                outHtml.append("</tr>\n</table>");
            }
//            }
        }
        else if (wb instanceof XSSFWorkbook)
        {
            xwb=new XSSFWorkbook(new FileInputStream(inputPath));
            int sheetnum=xwb.getNumberOfSheets();


//            for(int nums=0;nums<sheetnum;nums++){
                int nums = 0;
                XSSFSheet sheet = xwb.getSheetAt(nums);

                //获取图片并导出
                 maplist = getPicturesxlsx((XSSFSheet) sheet);

                 //打印图片
                 printImg(maplist);


                Map<Integer,Integer> xyArray = new HashMap<Integer,
                         Integer>();

                  //获取图片单元格坐标
                 for (Map.Entry<String, PictureData> entry :
                         maplist.entrySet()) {
                        System.out.println("Key = " + entry.getKey() + ", Value = " + entry.getValue());

                 //定义x坐标
                     xyArray.put(Integer.parseInt(entry.getKey().split("-")[0]),
                             Integer.parseInt(entry.getKey().split("-")[1]));
                 }


                String name = sheet.getSheetName();
                int numr =sheet.getPhysicalNumberOfRows();
                int numc =sheet.getLeftCol();
//                out.format("<p>工作页码"+nums+"</p>\n" +
//                        "<p>.工作空间名称:"+name+"</p>\n"+
//                        "<p>行列数："+ numr+" ,"+numc+"<p>");
                for(int i=0;i<numr;i++){
                    XSSFRow xssfRow = sheet.getRow(i);
                    numc=xssfRow.getPhysicalNumberOfCells();
//                    out.format("<table ><tr>\n");
                    outHtml.append("<table ><tr>\n");

                    int y=0;

                    //如果包含x坐标
                    boolean result = xyArray.containsKey(i);
                    if(result){
                        y=xyArray.get(i);
                    }

                    for(int j=0;j<numc;j++){
                        if(result && j==y){
                            //获取图片后缀名
                            PictureData picture = maplist.get((i)+
                                    "-"+y);

                            String fileName =
                                    picture.getMimeType().substring(6);

                            outHtml.append("<td><img src=\"./"+i+"-"+y+
                                            "."+fileName+
                                    "\" /></td" +
                                    ">\n" );

                            //移除表格内图片
                            maplist.remove((i)+
                                    "-"+y);

                        }else{
                            XSSFCell cell =xssfRow.getCell(j);
                            String str =getCellValue(cell);
                            outHtml.append("<td>"+str+"</td>\n" );
                        }
                    }
                    outHtml.append("</tr>\n</table>");
                }

                //处理剩下的表格外图片

            for(Map.Entry<String, PictureData> map : maplist.entrySet()){
                outHtml.append("<table ><tr>\n");

                PictureData picture = maplist.get(map.getKey());

                String fileName =
                        picture.getMimeType().substring(6);
                outHtml.append("<td><img src=\"./"+map.getKey()+
                        "."+fileName+
                        "\" /></td" +
                        ">\n" );
                outHtml.append("</tr>\n</table>");
            }
        }
        else
            throw new IllegalArgumentException(
                    "unknown workbook type: " + wb.getClass().getSimpleName());

    }

    public void printPage() throws IOException {
        try {
            ensureOut();
            int sheetnum = wb.getNumberOfSheets();

//            out.format("<body>%n");
//            out.format("<style type=\"text/css\">\n" +
//                    "td\n" +
//                    "  {\n" +
//                    "  height:18px;\n" +
//                    "  width:130px;\n" +
//                    "  vertical-align:middle;\n" +
//                    "  }table,th\n" +
//                    "  {\n" +
//                    "  border: 1px solid black;\n" +
//                    "  font-size:10px;\n" +
//                    "  }table\n" +
//                    "  {\n" +
//                    "  border-collapse:collapse;\n" +
//                    "  width:100%%;\n" +
//                    "  }\n" +
//                    "  </style>\n" +
//                    "  </body>");

            StringBuilder outHtml = new StringBuilder();
            outHtml.append("<!DOCTYPE html>%n");
            outHtml.append("<html>%n");
            outHtml.append("<head>%n");
            outHtml.append("<meta charset=\"utf-8\">%n");
            outHtml.append("<title>Excel转换HTML测试version1</title>%n");
            outHtml.append("</head>%n");
            outHtml.append("<body>%n");
            outHtml.append("<style type=\"text/css\">\n" +
                    "td\n" +
                    "  {\n" +
                    "  height:18px;\n" +
                    "  width:130px;\n" +
                    "  vertical-align:middle;\n" +
                    "  }table,th\n" +
                    "  {\n" +
                    "  border: 1px solid black;\n" +
                    "  font-size:10px;\n" +
                    "  }table\n" +
                    "  {\n" +
                    "  border-collapse:collapse;\n" +
                    "  width:100%%;\n" +
                    "  }\n" +
                    "  </style>\n" +
                    "  </body>");

            outHtml.append("<body>%n");

            printsheets(outHtml);
            if (true) {

                outHtml.append("</body>%n");
                outHtml.append("</html>%n");

                System.out.println(outHtml.toString());
                String str=outHtml.toString();
                out.format(str);
            }
        } finally {
            if (out != null)
                out.close();
            if (output instanceof Closeable) {
                Closeable closeable = (Closeable) output;
                closeable.close();
                wb.close();
            }
        }
    }

    private void ensureOut() {
        if (out == null)
            out = new Formatter(output);
    }

    private String getCellValue(Cell cell) {
        String cellvalue = "";
        if (cell != null) {
            // 判断当前Cell的Type
            switch (cell.getCellType()) {
                // 如果当前Cell的Type为NUMERIC
                case HSSFCell.CELL_TYPE_NUMERIC: {
                    short format = cell.getCellStyle().getDataFormat();
                    if(format == 14 || format == 31 || format == 57 || format == 58){   //excel中的时间格式
                        SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd");
                        double value = cell.getNumericCellValue();
                        Date date = DateUtil.getJavaDate(value);
                        cellvalue = sdf.format(date);
                    }
                    // 判断当前的cell是否为Date
                    else if (HSSFDateUtil.isCellDateFormatted(cell)) {  //先注释日期类型的转换，在实际测试中发现HSSFDateUtil.isCellDateFormatted(cell)只识别2014/02/02这种格式。
                        // 如果是Date类型则，取得该Cell的Date值           // 对2014-02-02格式识别不出是日期格式
                        Date date = cell.getDateCellValue();
                        DateFormat formater = new SimpleDateFormat("yyyy-MM-dd");
                        cellvalue= formater.format(date);
                    } else { // 如果是纯数字
                        // 取得当前Cell的数值
                        cellvalue = NumberToTextConverter.toText(cell.getNumericCellValue());

                    }
                    break;
                }
                // 如果当前Cell的Type为STRIN
                case HSSFCell.CELL_TYPE_STRING:
                    // 取得当前的Cell字符串
                    cellvalue = cell.getStringCellValue().replaceAll("'", "''");
                    break;
                case  HSSFCell.CELL_TYPE_BLANK:
                    cellvalue = "";
                    break;
                // 默认的Cell值
                default:{
                    cellvalue = "";
                }
            }
        } else {
            cellvalue = "";
        }
        return cellvalue;
    }

    /**
     * 获取图片和位置 (xls)
     * @param sheet
     * @return
     * @throws IOException
     */
    private Map<String, PictureData> getPicturesxls (HSSFSheet sheet) throws IOException {
        Map<String, PictureData> map = new HashMap<String, PictureData>();
        List<HSSFShape> list = sheet.getDrawingPatriarch().getChildren();
        for (HSSFShape shape : list) {
            if (shape instanceof HSSFPicture) {
                HSSFPicture picture = (HSSFPicture) shape;
                HSSFClientAnchor cAnchor = (HSSFClientAnchor) picture.getAnchor();
                PictureData pdata = picture.getPictureData();
                String key = cAnchor.getRow1() + "-" + cAnchor.getCol1(); // 行号-列号
                map.put(key, pdata);
            }
        }
        return map;
    }

    /**
     * 获取图片和位置 (xlsx)
     * @param sheet
     * @return
     * @throws IOException
     */
    private Map<String, PictureData> getPicturesxlsx (XSSFSheet sheet) throws IOException {
        Map<String, PictureData> map = new HashMap<String, PictureData>();
        List<POIXMLDocumentPart> list = sheet.getRelations();
        for (POIXMLDocumentPart part : list) {
            if (part instanceof XSSFDrawing) {
                XSSFDrawing drawing = (XSSFDrawing) part;
                List<XSSFShape> shapes = drawing.getShapes();
                for (XSSFShape shape : shapes) {
                    XSSFPicture picture = (XSSFPicture) shape;
                    XSSFClientAnchor anchor = picture.getPreferredSize();
                    CTMarker marker = anchor.getFrom();
                    String key = marker.getRow() + "-" + marker.getCol();
                    map.put(key, picture.getPictureData());
                }
            }
        }
        return map;
    }

    //图片写出
    private void printImg(Map<String, PictureData> sheetList) throws IOException {

        //for (Map<String, PictureData> map : sheetList) {
        Object key[] = sheetList.keySet().toArray();
        for (int i = 0; i < sheetList.size(); i++) {
            // 获取图片流
            PictureData pic = sheetList.get(key[i]);
            // 获取图片索引
            String picName = key[i].toString();
            // 获取图片格式
            String ext = pic.suggestFileExtension();

            byte[] data = pic.getData();

            //图片保存路径
            FileOutputStream out = new FileOutputStream("/Users/lonnielu" +
                    "/Documents/code/excelToHtml/html/" + picName + "." + ext);
            out.write(data);
            out.close();
        }
        // }

    }


}

