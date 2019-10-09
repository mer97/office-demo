package top.leemer.poidemo.excel;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import top.leemer.poidemo.excel.bean.Roster;

import java.io.File;
import java.io.FileOutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.List;
import java.util.Scanner;

/**
 * @author LEEMER
 * Create Date: 2019-10-09
 */
public class PoiWriteExcel {
    private static final Logger LOGGER = LoggerFactory.getLogger(PoiWriteExcel.class);

    private static void writeExcel(List<Roster> rosters, Path filePath){
        String[] title = {"编号","姓名","手机号码","地址"};

        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet("sheet_name");
        XSSFRow firstRow = sheet.createRow(0);
        XSSFCell firstCell = null;
        for(int j = 0; j < title.length; j++){
            firstCell = firstRow.createCell(j);
            firstCell.setCellValue(title[j]);
        }

        for (int i = 0; i < rosters.size(); i++) {
            XSSFRow row = sheet.createRow(i+1);
            XSSFCell cell = row.createCell(0);
            cell.setCellValue(rosters.get(i).getId());
            cell = row.createCell(1);
            cell.setCellValue(rosters.get(i).getName());
            cell = row.createCell(2);
            cell.setCellValue(rosters.get(i).getPhone());
            cell = row.createCell(3);
            cell.setCellValue(rosters.get(i).getSite());
        }

        try(FileOutputStream out = new FileOutputStream(filePath.toFile())) {
            workbook.write(out);
        }catch (Exception e){
            LOGGER.error("Write excel failure", e);
        }
    }

    public static void run(){
        Scanner scanner = new Scanner(System.in);
        System.out.println("请输入文件绝对路径(如:E:/花名册.xlsx)：");
        String filePath = scanner.nextLine();
        File file = new File(filePath);
        Path parentPath = Paths.get(file.getParent());
        try {
            if (!Files.exists(parentPath)) {
                Files.createDirectories(parentPath);
            }
            if (!Files.exists(file.toPath())) {
                Files.createFile(file.toPath());
            }
        }catch (Exception e){
            e.printStackTrace();
        }

        List<Roster> rosters = new ArrayList<Roster>();
        rosters.add(new Roster("001", "杨过", "10088", "杨家村"));
        rosters.add(new Roster("002", "杨康", "10087", "杨家村"));

        writeExcel(rosters, file.toPath());
    }

}
