import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.util.ArrayList;

public class CreateDropDown {

    public CreateDropDown() {

    }

    public void PopulateDropDown(XSSFWorkbook workbook, XSSFSheet sheet, String sheetName, String fileLocation, Integer rowNum, Integer columnNum,Integer sheetNum) {
        ImportDropDownLists importLists = new ImportDropDownLists();

        Name namedRange;
        String reference;

        sheetName = sheetName.replaceAll("\\s+","");

        //Hidden sheet for Lookup
        XSSFSheet hiddenDropDown = workbook.createSheet(sheetName);
        ArrayList<String> dropDown = new ArrayList<String>();
        dropDown = importLists.OpenDropDownCSV(fileLocation);

        for (int i = 0; i < dropDown.size(); i++)

        {
            XSSFRow row = hiddenDropDown.createRow(i);
            XSSFCell cell = row.createCell(0);
            cell.setCellValue(dropDown.get(i));
        }

        //create name for Drop Down list constraint
        namedRange = workbook.createName();
        namedRange.setNameName(sheetName);
        reference = sheetName + "!$A$1:$A$" + dropDown.size();
        namedRange.setRefersToFormula(reference);

        //unselect that sheet because we will hide it later
        hiddenDropDown.setSelected(false);

        //data validations
        DataValidationHelper dvHelper = sheet.getDataValidationHelper();
        DataValidationConstraint dvConstraint = dvHelper.createFormulaListConstraint(sheetName);
        CellRangeAddressList addressList = new CellRangeAddressList(1, (rowNum - 1), columnNum, columnNum);
        DataValidation validation = dvHelper.createValidation(dvConstraint, addressList);

        validation.setErrorStyle(DataValidation.ErrorStyle.STOP);
        validation.createErrorBox("Validation Error", "Please select from the provided dropdown list.");
        validation.setShowErrorBox(true);

        sheet.addValidationData(validation);

        hiddenDropDown.protectSheet("torch");

        workbook.setActiveSheet(0);
        workbook.setSheetHidden(sheetNum,true);
    }
}