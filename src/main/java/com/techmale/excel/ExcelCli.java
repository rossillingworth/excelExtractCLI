package com.techmale.excel; /**
 Copyright [yyyy] [name of copyright owner]

 Licensed under the Apache License, Version 2.0 (the "License");
 you may not use this file except in compliance with the License.
 You may obtain a copy of the License at

 http://www.apache.org/licenses/LICENSE-2.0

 Unless required by applicable law or agreed to in writing, software
 distributed under the License is distributed on an "AS IS" BASIS,
 WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 See the License for the specific language governing permissions and
 limitations under the License.
 */

import com.beust.jcommander.JCommander;
import com.beust.jcommander.Parameter;
import com.beust.jcommander.converters.FileConverter;
import com.beust.jcommander.validators.PositiveInteger;
import com.sun.org.apache.xpath.internal.operations.Number;
import com.techmale.excel.validators.FileExists;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;

public class ExcelCli {

    @Parameter(names = "--file", converter = FileConverter.class, validateWith = FileExists.class,required = true)
    private File xlsFile;

    @Parameter(names = "--ls", description = "Show worksheet names")
    private boolean listWorksheets = false;

    @Parameter(names = "--old", description = "Older Excel file type (2003 or earlier)")
    private boolean oldFileType = false;

    @Parameter(names = "--sheetName", description = "Worksheet name")
    private String worksheetName;

    @Parameter(names = "--sheetNum", description = "Worksheet number",validateWith = PositiveInteger.class)
    private Integer worksheetNumber;

    @Parameter(names = "--row", description = "Data cell row",validateWith = PositiveInteger.class)
    private Integer row;

    @Parameter(names = "--col", description = "Data cell column (eg: A -> ZZZZ)")
    private Integer column;


    // ######################################################
    // ######################################################
    // ######################################################



    public void readAll() {

//        try {
//            FileInputStream excelFile = new FileInputStream(new File(FILE_NAME));
//            Workbook workbook = new XSSFWorkbook(excelFile);
//            Sheet datatypeSheet = workbook.getSheetAt(0);
//            Iterator<Row> iterator = datatypeSheet.iterator();
//
//            while (iterator.hasNext()) {
//
//                Row currentRow = iterator.next();
//                Iterator<Cell> cellIterator = currentRow.iterator();
//
//                while (cellIterator.hasNext()) {
//
//                    Cell currentCell = cellIterator.next();
//                    //getCellTypeEnum shown as deprecated for version 3.15
//                    //getCellTypeEnum ill be renamed to getCellType starting from version 4.0
//                    if (currentCell.getCellTypeEnum() == CellType.STRING) {
//                        System.out.print(currentCell.getStringCellValue() + "--");
//                    } else if (currentCell.getCellTypeEnum() == CellType.NUMERIC) {
//                        System.out.print(currentCell.getNumericCellValue() + "--");
//                    }
//
//                }
//                System.out.println();
//
//            }
//        } catch (FileNotFoundException e) {
//            e.printStackTrace();
//        } catch (IOException e) {
//            e.printStackTrace();
//        }

    }

    public void start(){

        try {

            // check if --ls
            if(listWorksheets){
                listWorksheets();
                return;
            }else{
                showCellInfo();
                return;
            }

        } catch (Exception e){
            e.printStackTrace();
        }

        System.out.println("Unexpected this is ....");

    }


    private void listWorksheets() throws Exception{
        // get book
        // show a list of sheets

        FileInputStream excelFile = new FileInputStream(xlsFile);
        Workbook workbook = new XSSFWorkbook(excelFile);
        int countSheets = workbook.getNumberOfSheets();

        for (int i = 0; i < countSheets; i++) {
            Sheet datatypeSheet = workbook.getSheetAt(i);
            String name = datatypeSheet.getSheetName();
            System.out.println(String.format("%s: '%s'",i,name));
        }

    }

    private void showCellInfo() throws Exception{
        // get book
        // get sheet
        // get cell
        // show it

        assert(xlsFile != null);
        assert(worksheetName!=null || worksheetNumber!=null);
        assert(row!=null);
        assert(column!=null);

        FileInputStream excelFile = new FileInputStream(xlsFile);
        Workbook workbook = new XSSFWorkbook(excelFile);

        Sheet dataTypeSheet;
        if(worksheetName != null && worksheetName.length()>0){
            dataTypeSheet= workbook.getSheet(worksheetName);
        }else{
            dataTypeSheet= workbook.getSheetAt(worksheetNumber);
        }

        // move coords to zero rated
        int r = row -1;
        int c = column - 1;

        Cell cell = dataTypeSheet.getRow(r).getCell(c);

        String value="UNASSIGNED";
        if (cell.getCellTypeEnum() == CellType.STRING) {
            value = cell.getStringCellValue();
        } else if (cell.getCellTypeEnum() == CellType.NUMERIC) {
            value = Double.toString(cell.getNumericCellValue());
        }

        System.out.println(value);

    }

    // ######################################################
    // ######################################################
    // ######################################################

    /**
     * Parse CLI args and then run the CLI.
     *
     * @param args
     */
    public static void main(String[] args) {

        ExcelCli cli = new ExcelCli();

        JCommander.newBuilder()
                .addObject(cli)
                .build()
                .parse(args);
        cli.start();

    }


}