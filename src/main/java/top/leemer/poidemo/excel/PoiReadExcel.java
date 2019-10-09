package top.leemer.poidemo.excel;

import org.apache.http.entity.ContentType;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.mock.web.MockMultipartFile;
import org.springframework.web.multipart.MultipartFile;
import top.leemer.poidemo.excel.bean.Roster;

import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;
import java.util.Scanner;

/**
 * @author LEEMER
 * Create Date: 2019-10-09
 */
public class PoiReadExcel {

    private static final String EXCEL_XLS = "xls";
    private static final String EXCEL_XLSX = "xlsx";

    private static void checkExcelVaild(MultipartFile file) throws Exception {
        if (file == null){
            throw new Exception("文件不存在");
        }
        if (!((file.getName().endsWith(EXCEL_XLS) || file.getName().endsWith(EXCEL_XLSX)))){
            throw new Exception("文件不是Excel");
        }
    }

    private static String getCellValue(Cell cell){
        String cellValue = "";
        if(cell == null){
            return cellValue;
        }
        //把数字当成String来读，避免出现1读成1.0的情况
        if(cell.getCellType() == Cell.CELL_TYPE_NUMERIC){
            cell.setCellType(Cell.CELL_TYPE_STRING);
        }
        //判断数据的类型
        switch (cell.getCellType()){
            case Cell.CELL_TYPE_NUMERIC: //数字
                cellValue = String.valueOf(cell.getNumericCellValue());
                break;
            case Cell.CELL_TYPE_STRING: //字符串
                cellValue = String.valueOf(cell.getStringCellValue());
                break;
            case Cell.CELL_TYPE_BOOLEAN: //Boolean
                cellValue = String.valueOf(cell.getBooleanCellValue());
                break;
            case Cell.CELL_TYPE_FORMULA: //公式
                cellValue = String.valueOf(cell.getCellFormula());
                break;
            case Cell.CELL_TYPE_BLANK: //空值
                cellValue = "";
                break;
            case Cell.CELL_TYPE_ERROR: //故障
                cellValue = "非法字符";
                break;
            default:
                cellValue = "未知类型";
                break;
        }
        return cellValue;
    }

    private static List<Roster> readExcel(MultipartFile file) throws Exception {
        List<Roster> rosters = new ArrayList<Roster>();
        try {
            checkExcelVaild(file);
        }catch (Exception e){
            e.printStackTrace();
            return null;
        }

        //文件流
        InputStream is = file.getInputStream();
        Workbook workbook = null;
        if(file.getName().endsWith(EXCEL_XLS)){  //Excel 2003
            workbook = new HSSFWorkbook(is);
        }else if(file.getName().endsWith(EXCEL_XLSX)){  // Excel 2007/2010
            workbook = new XSSFWorkbook(is);
        }

        //获取Excel文件第一个sheet
        Sheet sheet = workbook.getSheetAt(0);

        int count = 0;
        //遍历每一行数据
        for (Row row : sheet){
            //过滤第一行数据
            if (count<1){
                count++;
                continue;
            }
            Roster roster = new Roster();
            //遍历每一列数据
            int c = 0;
            for (Cell cell : row){
                String value = getCellValue(cell);
                switch (c){
                    case 0:
                        roster.setId(value);
                        break;
                    case 1:
                        roster.setName(value);
                        break;
                    case 2:
                        roster.setPhone(value);
                        break;
                    case 3:
                        roster.setSite(value);
                        break;
                }
                c++;
            }
            rosters.add(roster);
        }
        return rosters;
    }

    public static void run(){
        List<Roster> rosters = null;
        Scanner scanner = new Scanner(System.in);
        System.out.println("请输入文件绝对路径(如:E:/花名册.xlsx)：");
        File file = new File(scanner.nextLine());
        try(FileInputStream inputStream = new FileInputStream(file)) {
            //File 转 MultipartFile
            MultipartFile mMultipartFile = new MockMultipartFile(file.getName(),file.getName(),
                    ContentType.APPLICATION_OCTET_STREAM.toString(),inputStream);
            rosters = readExcel(mMultipartFile);
        } catch (Exception e) {
            e.printStackTrace();
        }
        rosters.forEach( roster -> System.out.println(roster.toString()));
    }

}
