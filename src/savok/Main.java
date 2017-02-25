package savok;

import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Row;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Scanner;

public class Main {

    public static void main(String[] args) throws IOException {
        System.out.print("Enter path: ");
        Scanner sc = new Scanner(System.in);
        String path = sc.nextLine();
        POIFSFileSystem fs = new POIFSFileSystem(new File(path));
        HSSFWorkbook book = new HSSFWorkbook(fs);
        HSSFSheet sheet = book.getSheetAt(0);
        ArrayList<HSSFRow> rows = new ArrayList<>();
        String question;
        for (Row r_t : sheet) {
            HSSFRow r = (HSSFRow) r_t;
            question = r.getCell(0).toString();
            if (question.length() != 0) {
                while (!Character.isLetter(question.charAt(0)))
                    question = question.substring(1);
                while (!Character.isLetter(question.charAt(question.length() - 1))) {
                    question = question.substring(0, question.length() - 1);
                }
            }
            boolean isExists = false;
            for (HSSFRow r1: rows) {
                if (r1.getCell(0).toString().equals(question)) {
                    isExists = true;
                    break;
                }
            }
            if (!isExists)
                rows.add(r);
        }
        HSSFWorkbook newWorkbook = new HSSFWorkbook();
        HSSFSheet newSheet = newWorkbook.createSheet();
        for (int i = 0; i < rows.size(); i++) {
            HSSFRow row = newSheet.createRow(i);
            HSSFRow row1 = rows.get(i);
            row.createCell(0).setCellValue(row1.getCell(0).toString());
            row.createCell(1).setCellValue(row1.getCell(1).toString());
        }
        FileOutputStream fos = new FileOutputStream(new File(path + "_new.xls"));
        newWorkbook.write(fos);
        fos.close();
        System.out.println("Done!");
    }
}

