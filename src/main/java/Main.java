
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.*;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

import java.io.FileOutputStream;
import java.io.IOException;
import java.time.format.DateTimeFormatter;
import java.util.Collections;
import java.util.ArrayList;

public class Main {

    private static String ADDREPLACE = "Add-Replace";
    private static String FISCALYEAR = "Fiscal Year";
    private static String EMPLOYEEID = "Employee ID";
    private static String NCREGION = "Region";
    private static String EMPLOYEENAME = "Employee Name";
    private static String LOCATION = "Location";
    private static String JOBCODE = "Job Code";
    private static String JOBPROFILE = "Job Profile";
    private static String BUSINESSTITLE = "Business Title";
    private static String HIREDATE = "Hire Date";
    private static String TIMEINJOB = "Time In Job";
    private static String PLANELIGIBILITYSTARTDATE = "Plan Eligibility Start Date";
    private static String PLANELIGIBILITYENDDATE = "Plan Eligibility End Date";
    private static String INCENTIVEPLAN = "Incentive Plan";
    private static String BUSINESSUNITLABEL = "Business Unit Label                             ";
    private static String BUSINESSUNITMAXGOAL = "Business Unit Maximum Goal";
    private static String BUSINESSUNITTARGETGOAL = "Business Unit Target Goal";
    private static String BUSINESSUNITMINGOAL = "Business Unit Threshold Goal";
    private static String RQIMAXGOAL = "RQI Maximum Goal";
    private static String RQITARGETGOAL = "RQI Target Goal";
    private static String RQIMINGOAL = "RQI Threshold Goal";
    private static String SBRMAXGOAL = "SBR Maximum Goal";
    private static String SBRTARGETGOAL = "SBR Target Goal";
    private static String SBRMINGOAL = "SBR Threshold Goal";
    private static String SURPLUSLABEL = "Surplus Label                               ";
    private static String SURPLUSMAXGOAL = "Surplus Maximum Goal";
    private static String SURPLUSTARGETGOAL = "Surplus Target Goal";
    private static String SURPLUSMINGOAL = "Surplus Threshold Goal";
    private static String AWRMAXGOAL = "AWR Maximum Goal";
    private static String AWRTARGETGOAL = "AWR Target Goal";
    private static String AWRMINGOAL = "AWR Threshold Goal";
    private static String MAJORGIFTSMAXGOAL = "Major Gifts Maximum Goal";
    private static String MAJORGIFTSTARGETGOAL = "Major Gifts Target Goal";
    private static String MAJORGIFTSMINGOAL = "Major Gifts Threshold Goal";
    private static String INDIVIDUALGOALLABEL = "Individual Goal Label                         ";
    private static String INDIVIDUALGOALMAXGOAL = "Individual Maximum Goal";
    private static String INDIVIDUALGOALTARGETGOAL = "Individual Target Goal";
    private static String INDIVIDUALGOALMINGOAL = "Individual Threshold Goal";
    private static String TEAMWORKLABEL = "Teamwork Label                                ";
    private static String TEAMWORKMAXGOAL = "Teamwork Maximum Goal";
    private static String TEAMWORKTARGETGOAL = "Teamwork Target Goal";
    private static String TEAMWORKMINGOAL = "Teamwork Threshold Goal";
    private static String INITTOGETHERLABEL = "In It Together Label                             ";
    private static String INITTOGETHERMAXGOAL = "In It Together Maximum Goal";
    private static String INITTOGETHERTARGETGOAL = "In It Together Target Goal";
    private static String INITTOGETHERMINGOAL = "In It Together Threshold Goal";
    private static String DEPARTMENT = "Department";


    private static String[] columnsNAT = {ADDREPLACE, FISCALYEAR, INCENTIVEPLAN, EMPLOYEENAME,JOBPROFILE,TIMEINJOB,
                                       PLANELIGIBILITYSTARTDATE,PLANELIGIBILITYENDDATE,BUSINESSUNITLABEL,
                                       BUSINESSUNITMAXGOAL,BUSINESSUNITTARGETGOAL,BUSINESSUNITMINGOAL,RQIMAXGOAL,RQITARGETGOAL,RQIMINGOAL,
                                       SBRMAXGOAL,SBRTARGETGOAL,SBRMINGOAL,SURPLUSLABEL,SURPLUSMAXGOAL,SURPLUSTARGETGOAL,SURPLUSMINGOAL,
                                       AWRMAXGOAL,AWRTARGETGOAL,AWRMINGOAL,MAJORGIFTSMAXGOAL,MAJORGIFTSTARGETGOAL,MAJORGIFTSMINGOAL,
                                       INDIVIDUALGOALLABEL,INDIVIDUALGOALMAXGOAL,INDIVIDUALGOALTARGETGOAL,INDIVIDUALGOALMINGOAL,TEAMWORKLABEL,
                                       TEAMWORKMAXGOAL,TEAMWORKTARGETGOAL,TEAMWORKMINGOAL,INITTOGETHERLABEL,INITTOGETHERMAXGOAL,INITTOGETHERTARGETGOAL,
                                       INITTOGETHERMINGOAL,
                                       EMPLOYEEID, NCREGION, LOCATION,DEPARTMENT,JOBCODE,BUSINESSTITLE,HIREDATE};

    private static String[] columnsRegion = {ADDREPLACE, FISCALYEAR, INCENTIVEPLAN, EMPLOYEENAME,JOBPROFILE,TIMEINJOB,
                                        PLANELIGIBILITYSTARTDATE,PLANELIGIBILITYENDDATE,
                                        INDIVIDUALGOALLABEL,INDIVIDUALGOALMINGOAL,INDIVIDUALGOALMAXGOAL,TEAMWORKLABEL,
                                        TEAMWORKMINGOAL,TEAMWORKMAXGOAL,INITTOGETHERLABEL,
                                        INITTOGETHERMINGOAL,INITTOGETHERMAXGOAL,
                                        EMPLOYEEID, NCREGION, LOCATION,DEPARTMENT,JOBCODE,BUSINESSTITLE,HIREDATE};

    private static String [] columnsStart = {ADDREPLACE, FISCALYEAR, INCENTIVEPLAN, EMPLOYEENAME,JOBPROFILE,TIMEINJOB,
                                            PLANELIGIBILITYSTARTDATE,PLANELIGIBILITYENDDATE};

    private static String [] columnsEnd = {EMPLOYEEID, NCREGION, LOCATION,DEPARTMENT,JOBCODE,BUSINESSTITLE,HIREDATE};

    private static AuditFileEmployees auditFile;

    private static int employeesInInputTemplates;


    public static void main(String[] args) throws IOException, InvalidFormatException {

        // CREATE AUDIT FILE
        DateTimeFormatter timeStampPattern = DateTimeFormatter.ofPattern("yyyyMMddHHmmss");
        auditFile = new AuditFileEmployees("d:\\incentive\\audit\\","ImportTemplateCreationAuditFile-"+timeStampPattern.format(java.time.LocalDateTime.now())+".csv");
        auditFile.createAuditFile();
        employeesInInputTemplates = 0;

        // IMPORT SETTINGS FOR CLASSES LIKE GOALS AND ELIGIBLE EMPLOYEES
        ArrayList<GoalDefinition> goalDefinitions = new ArrayList<>();
        goalDefinitions = SettingImports.ImportGoalSettings("d://incentive//GoalDefinitions.csv");

        ArrayList<GoalColumnDefinition> goalColumnDefinitions = new ArrayList<>();
        goalColumnDefinitions = SettingImports.ImportGoalColumnSettings("d://incentive//goalcolumnsettings.csv");
        Collections.sort(goalColumnDefinitions, new GoalColumnChainedComparator(new GoalColumnOrderComparator(), new GoalColumnGoalOrderComparator()));

        // IMPORT ALL ELIGIBLE EMPLOYEES
        ArrayList<EmployeeIncentiveInput> employeeIncentiveInputs = new ArrayList<>();
        employeeIncentiveInputs = SettingImports.OpenEmployeeImportCSV("d://incentive//Incentive_Input_File.csv");

        try {
            auditFile.writeRow("Load Employees", "Number of Eligible Employees", Integer.toString(employeeIncentiveInputs.size()));
        } catch (IOException e) {
            e.printStackTrace();
        }

        // BREAK UP THE ELIGIBLE EMPLOYEES INTO THEIR VARIOUS FILE SEGMENTS
        ArrayList<EmployeeIncentiveInput> fieldIncentivePlans = new ArrayList<>();
        ArrayList<EmployeeIncentiveInput> natIncentivePlans = new ArrayList<>();
        ArrayList<EmployeeIncentiveInput> natIncentivePlansFieldFundraiser = new ArrayList<>();
        ArrayList<EmployeeIncentiveInput> natIncentivePlansRevenuePlan = new ArrayList<>();
        ArrayList<EmployeeIncentiveInput> natIncentivePlansRevenueGenerationPlan = new ArrayList<>();
        ArrayList<EmployeeIncentiveInput> natIncentivePlansECCCommunityCPRPlan = new ArrayList<>();
        ArrayList<EmployeeIncentiveInput> natIncentivePlansManagementPlan = new ArrayList<>();
        ArrayList<EmployeeIncentiveInput> natIncentivePlansSrManagementPlan = new ArrayList<>();
        ArrayList<EmployeeIncentiveInput> natIncentivePlansSrManagementEmergingPlan = new ArrayList<>();
        ArrayList<EmployeeIncentiveInput> natIncentivePlansExecutiveIncentivePlan = new ArrayList<>();
        ArrayList<EmployeeIncentiveInput> natIncentivePlansCEOIncentivePlan = new ArrayList<>();

        // FIRST LEVEL OF SPLITTING OUT THE EMPLOYEE BY REGION
        for(EmployeeIncentiveInput employees: employeeIncentiveInputs) {
            if (!employees.getRegion().equals("NAT")) {
                if (employees.getIncentivePlan().equals("Field Fundraiser Incentive Plan Below 250k")) {
                    employees.setEmployeeType("Field");
                    fieldIncentivePlans.add(employees);
                } else if (employees.getIncentivePlan().equals("Field Fundraising Management Incentive Plan")) {
                    employees.setEmployeeType("Field");
                    fieldIncentivePlans.add(employees);
                } else if (employees.getIncentivePlan().equals("Field Fundraiser Incentive Plan Above 250k")) {
                    employees.setEmployeeType("Field");
                    fieldIncentivePlans.add(employees);
                } else {
                    // EMPLOYEES WHO DO NOT HAVE THE ABOVE 3 INCENTIVE PLANS WILL BE HANDLED AS NAT EMPLOYEES
                    employees.setEmployeeType("NAT");
                    natIncentivePlans.add(employees);
                }
            } else {
                employees.setEmployeeType("NAT");
                natIncentivePlans.add(employees);
            }
        }

        // SORT THE COLLECTIONS BASED ON THE COMPARE OPERATOR IN THE EMPLOYEE INCENTIVE INPUT CLASS
        Collections.sort(fieldIncentivePlans);
        Collections.sort(natIncentivePlans);

        // PROCESS THE REGIONS THAT ARE NOT NAT AND CREATE XLS INPUT SHEETS FOR EACH REGION
        String regionStore = "";
        for(EmployeeIncentiveInput employees: fieldIncentivePlans) {
            if (employees.getEmployeeType().equals("Field")) {
                if (!regionStore.equals(employees.getRegion())) {
                    regionStore = employees.getRegion();
                    createIncentiveWorkSheet(regionStore, fieldIncentivePlans, "Field", goalDefinitions, goalColumnDefinitions,"General");
                }
            }
        }

        // PROCESS ALL THE NAT CENTER EMPLOYEES.
        // SUB-SEGMENT THE NAT EMPLOYEES BY INCENTIVE PLAN
        for(EmployeeIncentiveInput employees: natIncentivePlans) {
            switch (employees.getIncentivePlan()) {
                case "CEO Incentive Plan":
                    natIncentivePlansCEOIncentivePlan.add(employees);
                    break;
                case "ECC Community CPR Incentive Plan":
                    natIncentivePlansECCCommunityCPRPlan.add(employees);
                    break;
                case "Executive Incentive Plan":
                    natIncentivePlansExecutiveIncentivePlan.add(employees);
                    break;
                case "Field Fundraiser Incentive Plan Above 250k":
                    natIncentivePlansFieldFundraiser.add(employees);
                    break;
                case "Field Fundraising Management Incentive Plan":
                    natIncentivePlansFieldFundraiser.add(employees);
                    break;
                case "Field Fundraiser Incentive Plan Below 250k":
                    natIncentivePlansFieldFundraiser.add(employees);
                    break;
                case "Management Incentive Plan":
                    natIncentivePlansManagementPlan.add(employees);
                    break;
                case "Revenue Generating Incentive Plan":
                    natIncentivePlansRevenueGenerationPlan.add(employees);
                    break;
                case "Revenue Incentive Plan":
                    natIncentivePlansRevenuePlan.add(employees);
                    break;
                case "Sr Management Incentive Plan":
                    natIncentivePlansSrManagementPlan.add(employees);
                    break;
                case "Sr Management Incentive Plan - Emerging & Commercial Bus":
                    natIncentivePlansSrManagementEmergingPlan.add(employees);
                    break;
            }
        }

        // CREATE INPUT SHEETS FOR EACH INCENTIVE PLAN IN THE NAT
        if (natIncentivePlansCEOIncentivePlan.size() > 0) {
            createIncentiveWorkSheet("NAT", natIncentivePlansCEOIncentivePlan, "NAT", goalDefinitions, goalColumnDefinitions,"General");
        }
        if (natIncentivePlansECCCommunityCPRPlan.size() > 0) {
            createIncentiveWorkSheet("NAT", natIncentivePlansECCCommunityCPRPlan, "NAT", goalDefinitions, goalColumnDefinitions,"General");
        }
        if (natIncentivePlansExecutiveIncentivePlan.size() > 0) {
            createIncentiveWorkSheet("NAT", natIncentivePlansExecutiveIncentivePlan, "NAT", goalDefinitions, goalColumnDefinitions,"General");
        }
        if (natIncentivePlansFieldFundraiser.size() > 0) {
            createIncentiveWorkSheet("NAT", natIncentivePlansFieldFundraiser, "NAT", goalDefinitions, goalColumnDefinitions,"General");
        }
        if (natIncentivePlansManagementPlan.size() > 0) {
            createIncentiveWorkSheet("NAT", natIncentivePlansManagementPlan, "NAT", goalDefinitions, goalColumnDefinitions,"General");
        }
        // INCENTIVE PLAN REVENUE GENERATION WILL BE SEGMENTED FURTHER BY DEPARTMENT
        if (natIncentivePlansRevenueGenerationPlan.size() > 0) {

            ArrayList<EmployeeIncentiveInput> natIncentivePlanRevenueGenGeneral = new ArrayList<>();
            ArrayList<EmployeeIncentiveInput> natIncentivePlanRevenueGenMission = new ArrayList<>();
            ArrayList<EmployeeIncentiveInput> natIncentivePlanRevenueGenECC = new ArrayList<>();
            ArrayList<EmployeeIncentiveInput> natIncentivePlanRevenueGenInt = new ArrayList<>();

            for(EmployeeIncentiveInput employees: natIncentivePlansRevenueGenerationPlan) {
                switch (employees.getDepartment()) {
                    case "Field Campaigns - NC":
                    case "Marketing Communications-NC":
                    case "Quality & Health IT - NC":
                    case "Field Operations/Development":
                    case "Field Ops & Development - NC":
                        natIncentivePlanRevenueGenGeneral.add(employees);
                        break;
                    case "Mission Advancement - NC":
                    case "Corporate Relations - NC":
                        natIncentivePlanRevenueGenMission.add(employees);
                        break;
                    case "ECC Field - NC":
                    case "ECC Sci & Prod Dev - NC":
                        natIncentivePlanRevenueGenECC.add(employees);
                        break;
                    default:
                        natIncentivePlanRevenueGenInt.add(employees);
                        break;
                }
            }

            createIncentiveWorkSheet("NAT", natIncentivePlanRevenueGenGeneral, "NAT", goalDefinitions, goalColumnDefinitions,"General Plans");
            createIncentiveWorkSheet("NAT", natIncentivePlanRevenueGenMission, "NAT", goalDefinitions, goalColumnDefinitions, "Mission Advancement");
            createIncentiveWorkSheet("NAT", natIncentivePlanRevenueGenECC, "NAT", goalDefinitions, goalColumnDefinitions, "ECC");
            createIncentiveWorkSheet("NAT", natIncentivePlanRevenueGenInt, "NAT", goalDefinitions, goalColumnDefinitions, "International");

/*            String departmentStore = natIncentivePlansRevenueGenerationPlan.get(0).getDepartment();

            ArrayList<EmployeeIncentiveInput> natIncentivePlansRevenueGenerationPlanDepartment = new ArrayList<>();
            natIncentivePlansRevenueGenerationPlanDepartment.clear();

            for(EmployeeIncentiveInput employees: natIncentivePlansRevenueGenerationPlan) {

                if (!departmentStore.equals(employees.getDepartment())) {
                    departmentStore = employees.getDepartment();
                    if (natIncentivePlansRevenueGenerationPlanDepartment.size() > 0) {
                        createIncentiveWorkSheet("NAT", natIncentivePlansRevenueGenerationPlanDepartment, "NAT", goalDefinitions, goalColumnDefinitions);
                    }
                    natIncentivePlansRevenueGenerationPlanDepartment.clear();
                }
                natIncentivePlansRevenueGenerationPlanDepartment.add(employees);
            }
            createIncentiveWorkSheet("NAT", natIncentivePlansRevenueGenerationPlanDepartment, "NAT", goalDefinitions, goalColumnDefinitions);
*/
        }
        if (natIncentivePlansRevenuePlan.size() > 0) {
            ArrayList<EmployeeIncentiveInput> natIncentivePlansRevenuePlanMissionAdvancement = new ArrayList<>();
            ArrayList<EmployeeIncentiveInput> natIncentivePlansRevenuePlanOtherDepartments = new ArrayList<>();
            for (EmployeeIncentiveInput revenuePlanEmployees : natIncentivePlansRevenuePlan) {
                if (revenuePlanEmployees.getDepartment().equals("Mission Advancement - NC")) {
                    natIncentivePlansRevenuePlanMissionAdvancement.add(revenuePlanEmployees);
                } else {
                    natIncentivePlansRevenuePlanOtherDepartments.add(revenuePlanEmployees);
                }
            }
            if (natIncentivePlansRevenuePlanMissionAdvancement.size() > 0) {
                createIncentiveWorkSheet("NAT", natIncentivePlansRevenuePlanMissionAdvancement, "NAT", goalDefinitions, goalColumnDefinitions,"General");
            }
            if (natIncentivePlansRevenuePlanOtherDepartments.size() > 0) {
                createIncentiveWorkSheet("NAT", natIncentivePlansRevenuePlanOtherDepartments, "NAT", goalDefinitions, goalColumnDefinitions,"General");
            }
        }
        if (natIncentivePlansSrManagementPlan.size() > 0) {
            createIncentiveWorkSheet("NAT", natIncentivePlansSrManagementPlan, "NAT", goalDefinitions, goalColumnDefinitions,"General");
        }
        if (natIncentivePlansSrManagementEmergingPlan.size() > 0) {
            createIncentiveWorkSheet("NAT", natIncentivePlansSrManagementEmergingPlan, "NAT", goalDefinitions, goalColumnDefinitions,"General");
        }

        try {
            auditFile.writeRow("Audit Exit","Number of Employees in Input Templates",Integer.toString(employeesInInputTemplates));
            auditFile.saveAuditFile();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    public static void createIncentiveWorkSheet(String regionName, ArrayList<EmployeeIncentiveInput> employeeIncentiveInputs,String employeeType,
                                                ArrayList<GoalDefinition> goalDefinitions, ArrayList<GoalColumnDefinition> goalColumnDefinitions, String fileDescription) throws IOException, InvalidFormatException {

        // Create a Workbook
        XSSFWorkbook workbook = new XSSFWorkbook();     // new XSSFWorkbook() for generating `.xls` file

        /* CreationHelper helps us create instances for various things like DataFormat,
           Hyperlink, RichTextString etc in a format (HSSF, XSSF) independent way */
        CreationHelper createHelper = workbook.getCreationHelper();

        // Create a Sheet
        XSSFSheet sheet = (XSSFSheet) workbook.createSheet("Employee");

//        CTSheetProtection sheetProtection = sheet.getCTWorksheet().getSheetProtection();
//        sheetProtection.setSort(true);

        // Create a Font for styling header cells
        Font headerFont = workbook.createFont();
        headerFont.setBold(true);
        headerFont.setFontHeightInPoints((short) 11);
        headerFont.setColor(IndexedColors.RED.getIndex());

        Font lockedFont = workbook.createFont();
        lockedFont.setBold(false);
        lockedFont.setFontHeightInPoints((short) 11);
        lockedFont.setColor(IndexedColors.BLUE.getIndex());

        Font unlockedFont = workbook.createFont();
        unlockedFont.setBold(true);
        unlockedFont.setFontHeightInPoints((short) 11);
        unlockedFont.setColor(IndexedColors.BLACK.getIndex());

        // Create a CellStyle with the font
        CellStyle headerCellStyle = workbook.createCellStyle();
        headerCellStyle.setFont(headerFont);
        headerCellStyle.setFillBackgroundColor(IndexedColors.RED.getIndex());

        CellStyle unlockedCellStyle = workbook.createCellStyle();
        unlockedCellStyle.setLocked(false);
        unlockedCellStyle.setFont(unlockedFont);

        DataFormat numberFormat = workbook.createDataFormat();

        CellStyle unlockedCellStyleNumber = workbook.createCellStyle();
        unlockedCellStyleNumber.setLocked(false);
        unlockedCellStyleNumber.setFont(unlockedFont);
        unlockedCellStyleNumber.setDataFormat(numberFormat.getFormat("###,###,###,###.##"));

        CellStyle lockedCellStyle = workbook.createCellStyle();
        lockedCellStyle.setLocked(true);
        lockedCellStyle.setFont(lockedFont);

        // Create a Row
        Row headerRow = sheet.createRow(0);

        ArrayList<String> columns = new ArrayList<>();
        ArrayList<String> goalsToInclude = new ArrayList<>();

        // Creating headers
        Collections.addAll(columns, columnsStart);

        for (EmployeeIncentiveInput employee : employeeIncentiveInputs) {
            if (regionName.equals(employee.getRegion()) || employee.getEmployeeType().equals("NAT")) {
                for (GoalDefinition goals : goalDefinitions) {
                    if (goals.getIncentivePlanName().equals(employee.getIncentivePlan())) {
                        boolean foundGoal = false;
                        for (String goalsToCheck : goalsToInclude) {
                            if (goalsToCheck.equals(goals.getGoalName())) {
                                foundGoal = true;
                            }
                        }
                        if (!goals.getGoalName().equals("Beyond Plan Goal") && !goals.getGoalName().equals("Bonus Adder (> 250K Spot Award)")) {
                            if (!foundGoal) {
                                goalsToInclude.add(goals.getGoalName());
                            }
                        }
                    }
                }
            }
        }

        ArrayList<GoalColumnDefinition> goalColsToInclude = new ArrayList<>();
        for (String goalNameToInclude : goalsToInclude) {
            for (GoalColumnDefinition goalColDef : goalColumnDefinitions) {
                if (goalNameToInclude.equals(goalColDef.getLinkedGoalName())) {
                    if (!regionName.equals(goalColDef.getRegionException1())) {
                        if (!regionName.equals(goalColDef.getRegionException2())) {
                            if (!regionName.equals(goalColDef.getRegionException3())) {
                                if (!regionName.equals(goalColDef.getRegionException4())) {
                                    if (!regionName.equals(goalColDef.getRegionException5())) {
                                        goalColsToInclude.add(goalColDef);
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }

        Collections.sort(goalColsToInclude, new GoalColumnChainedComparator(new GoalColumnOrderComparator(), new GoalColumnGoalOrderComparator()));

        for (GoalColumnDefinition includeGoalColDef : goalColsToInclude) {
            columns.add(includeGoalColDef.getColumnName());
        }

        Collections.addAll(columns, columnsEnd);

/*
        System.out.println("********************Columns to print********************");
        for (String colsToPrint : columns) {
            System.out.println("Column Header:"+colsToPrint);
        }
        System.out.println("********************************************************\n\n");
*/
        for (int i = 0; i < columns.size(); i++) {
            Cell cell = headerRow.createCell(i);
            cell.setCellValue(columns.get(i));
            cell.setCellStyle(headerCellStyle);
        }

        // Cell Style for formatting Date
        CellStyle dateCellStyleUnlocked = workbook.createCellStyle();
        dateCellStyleUnlocked.setDataFormat(createHelper.createDataFormat().getFormat("MM/dd/yyyy"));
        dateCellStyleUnlocked.setLocked(false);
        dateCellStyleUnlocked.setFont(unlockedFont);

        CellStyle dateCellStyleLocked = workbook.createCellStyle();
        dateCellStyleLocked.setDataFormat(createHelper.createDataFormat().getFormat("MM/dd/yyyy"));
        dateCellStyleLocked.setLocked(true);
        dateCellStyleLocked.setFont(lockedFont);

        // Create Other rows and cells with employees data
        int rowNum = 1;

        String incentivePlanStore = "";

        ArrayList<String> comboSelectionFiles = new ArrayList<>();
        ArrayList<Integer> comboSelectionColNum = new ArrayList<>();
        ArrayList<String> comboSelectionColName = new ArrayList<>();

        for(EmployeeIncentiveInput employees: employeeIncentiveInputs) {

            if (regionName.equals(employees.getRegion()) || employees.getEmployeeType().equals("NAT")) {

                Row row = sheet.createRow(rowNum++);
                Integer colNum = 0;

                if (employees.getIncentivePlan().equals("Field Fundraiser Incentive Plan Above 250k") ||
                        employees.getIncentivePlan().equals("Field Fundraising Management Incentive Plan") ||
                        employees.getIncentivePlan().equals("Field Fundraiser Incentive Plan Below 250k")) {

                    incentivePlanStore = "Field Fundraising Plan";
                } else {
                    incentivePlanStore = employees.getIncentivePlan();
                }

                for (String goalColumn : columns) {

                    switch (goalColumn) {
                        case "Add-Replace":
                            row.createCell(colNum).setCellValue("Add");
                            row.getCell(colNum).setCellStyle(unlockedCellStyle);
                            break;
                        case "Fiscal Year":
                            row.createCell(colNum).setCellValue("2019");
                            row.getCell(colNum).setCellStyle(lockedCellStyle);
                            break;
                        case "Incentive Plan":
                            row.createCell(colNum).setCellValue(employees.getIncentivePlan());
                            row.getCell(colNum).setCellStyle(lockedCellStyle);
                            break;
                        case "Employee Name":
                            row.createCell(colNum).setCellValue(employees.getEmployeeName());
                            row.getCell(colNum).setCellStyle(lockedCellStyle);
                            break;
                        case "Job Profile":
                            row.createCell(colNum).setCellValue(employees.getJobProfile());
                            row.getCell(colNum).setCellStyle(lockedCellStyle);
                            break;
                        case "Time In Job":
                            Cell timeInJob = row.createCell(colNum);
                            timeInJob.setCellValue(employees.getTimeInJob());
                            timeInJob.setCellStyle(dateCellStyleLocked);
                            break;
                        case "Plan Eligibility Start Date":
                            Cell eligibilityStartDate = row.createCell(colNum);
                            eligibilityStartDate.setCellValue("07/01/2019");
                            eligibilityStartDate.setCellStyle(dateCellStyleUnlocked);
                            break;
                        case "Plan Eligibility End Date":
                            Cell eligibilityEndDate = row.createCell(colNum);
                            eligibilityEndDate.setCellValue("06/30/2020");
                            eligibilityEndDate.setCellStyle(dateCellStyleUnlocked);
                            break;
                        case "Employee ID":
                            row.createCell(colNum).setCellValue(employees.getEmployeeID());
                            row.getCell(colNum).setCellStyle(lockedCellStyle);
                            break;
                        case "Region":
                            row.createCell(colNum).setCellValue(employees.getRegion());
                            row.getCell(colNum).setCellStyle(lockedCellStyle);
                            break;
                        case "Location":
                            row.createCell(colNum).setCellValue(employees.getLocation());
                            row.getCell(colNum).setCellStyle(lockedCellStyle);
                            break;
                        case "Department":
                            row.createCell(colNum).setCellValue(employees.getDepartment());
                            row.getCell(colNum).setCellStyle(lockedCellStyle);
                            break;
                        case "Job Code":
                            row.createCell(colNum).setCellValue(employees.getJobCode());
                            row.getCell(colNum).setCellStyle(lockedCellStyle);
                            break;
                        case "Business Title":
                            row.createCell(colNum).setCellValue(employees.getBusinessTitle());
                            row.getCell(colNum).setCellStyle(lockedCellStyle);
                            break;
                        case "Hire Date":
                            Cell hiredDate = row.createCell(colNum);
                            hiredDate.setCellValue(employees.getHireDate());
                            hiredDate.setCellStyle(dateCellStyleLocked);
                            break;
                        default:
                            for (GoalColumnDefinition goalColDef : goalColumnDefinitions) {
                                if (goalColDef.getColumnName().equals(goalColumn)) {
                                    String cellValue = "";
                                    if (goalColDef.getDefaultValue() != null && !goalColDef.getDefaultValue().equals("")) {
                                        cellValue = goalColDef.getDefaultValue();
                                    }
                                    row.createCell(colNum).setCellValue(cellValue);
                                    if (goalColDef.isColumnLocked()) {
                                        row.getCell(colNum).setCellStyle(lockedCellStyle);
                                    } else {
                                        row.getCell(colNum).setCellStyle(unlockedCellStyle);
                                    }
                                }
                            }
                    }
                    colNum++;
                }
            }
        }

        // SET THE FILENAME FOR THE XLS INPUT TEMPLATE
        String fileName;
        if(employeeIncentiveInputs.get(0).getIncentivePlan().equals("Revenue Generating Incentive Plan")) {
            employeeType = "NAT Revenue Generation";
        } else if (employeeIncentiveInputs.get(0).getIncentivePlan().equals("Revenue Incentive Plan")) {
            if (!employeeIncentiveInputs.get(0).getDepartment().equals("Mission Advancement - NC")) {
                employeeType = "NAT Revenue Plan Other";
            } else {
                employeeType = "NAT Revenue Plan Mission";
           }
        }

        switch (employeeType) {
            case "NAT":
                if (fileDescription.equals("General")) {
                    fileName = "D:/incentive/output/IncentiveDataGatheringSheet-NAT-" + incentivePlanStore + ".xlsx";
                } else {
                    fileName = "D:/incentive/output/IncentiveDataGatheringSheet-NAT-" + fileDescription + ".xlsx";
                }
                break;
            case "NAT Revenue Generation":
                fileName = "D:/incentive/output/IncentiveDataGatheringSheet-NAT-" + incentivePlanStore + "-"+ fileDescription + ".xlsx";
//                String department = employeeIncentiveInputs.get(0).getDepartment();
//                if (department.equals("Field Operations/Development")) {
//                    department = "Field Operations Development";
//                }
//                fileName = "D:/incentive/output/IncentiveDataGatheringSheet-NAT-" + incentivePlanStore + "-"+department+".xlsx";
                break;
            case "NAT Revenue Plan Other":
                fileName = "D:/incentive/output/IncentiveDataGatheringSheet-NAT-" + incentivePlanStore + "-Other Departments.xlsx";
                break;
            case "NAT Revenue Plan Mission":
                fileName = "D:/incentive/output/IncentiveDataGatheringSheet-NAT-" + incentivePlanStore + "-Mission Advancement.xlsx";
                break;
            default:
                fileName = "D:/incentive/output/IncentiveDataGatheringSheet-" + regionName + ".xlsx";
        }

        // CHECK TO SEE IF THERE IS MORE THAN JUST THE HEADER ROW BEFORE ADDING DROP DOWN FIELDS IN THE LIST
        if (rowNum > 1) {
            int columnFindIncrement = 0;
            for (String colComboFind : columns) {
                if (colComboFind.equals("Add-Replace")) {
                    comboSelectionColNum.add(columnFindIncrement);
                    comboSelectionFiles.add("addreplace.csv");
                    comboSelectionColName.add("AddReplace");
                }
                for (GoalColumnDefinition findGoalColDef : goalColumnDefinitions) {
                    if (colComboFind.equals(findGoalColDef.getColumnName())) {
                        if (findGoalColDef.getColumnType().equals("Dropdown")) {
                            comboSelectionColNum.add(columnFindIncrement);
                            comboSelectionFiles.add(findGoalColDef.getSelectionFile());
                            comboSelectionColName.add(findGoalColDef.getColumnName());
                        }
                    }
                }
                columnFindIncrement++;
            }

            int worksheetNumber = 1;

            int s = 0;
            CreateDropDown createDropDown = new CreateDropDown();
            for (String sheetName : comboSelectionColName) {
                createDropDown.PopulateDropDown(workbook, sheet, sheetName, "d://incentive//" + comboSelectionFiles.get(s), rowNum, comboSelectionColNum.get(s), worksheetNumber++);
                s++;
            }
        }

        // RESIZE ALL COLUMNS TO FIT THE CONTENT SIZE
        for(int i = 0; i < columns.size(); i++) {
            sheet.autoSizeColumn(i);
        }
        String filterRange;

        filterRange = "A1:"+ convertToTitle(columns.size())+rowNum;

        sheet.setAutoFilter(CellRangeAddress.valueOf(filterRange));
        sheet.lockAutoFilter(false);

        sheet.protectSheet("torch");
        sheet.createFreezePane(5, 1);

        // WRITE THE XLS INPUT SHEET FOR THE EXCEL FILE
        FileOutputStream fileOut = new FileOutputStream(fileName);
        workbook.write(fileOut);
        fileOut.close();

        workbook.close();

        //WRITE THE AUDIT ENTRY VALUES
        auditFile.writeRow("Input File Creation","Number of employees in "+fileName,Integer.toString(rowNum-1));
        employeesInInputTemplates = employeesInInputTemplates + rowNum - 1;
    }

    public static int titleToNumber(String s) {
        int result = 0;
        for (char c : s.toCharArray()) {
            result = result * 26 + (c - 'A') + 1;
        }
        return result;
    }

    public static String convertToTitle(int n) {
        if(n <= 0){
            throw new IllegalArgumentException("Input is not valid!");
        }

        StringBuilder sb = new StringBuilder();

        while(n > 0){
            n--;
            char ch = (char) (n % 26 + 'A');
            n /= 26;
            sb.append(ch);
        }

        sb.reverse();
        return sb.toString();
    }
}

