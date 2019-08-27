import java.util.Date;
import org.apache.commons.lang3.StringUtils;

public class EmployeeIncentiveInput implements Comparable< EmployeeIncentiveInput > {

    private String employeeID;
    private String employeeName;
    private String region;
    private String jobCode;
    private String jobProfile;
    private String businessTitle;
    private Date hireDate;
    private Date timeInJob;
    private String incentivePlan;
    private String department;
    private String location;

    private String sortCriteria1;
    private String sortCriteria2;
    private String sortCriteria3;

    private String employeeType;

    public String getSortCriteria1() {
        return sortCriteria1;
    }

    public String getSortCriteria2() {
        return sortCriteria2;
    }

    public String getSortCriteria3() {
        return sortCriteria3;
    }

    public String getEmployeeType() {
        return employeeType;
    }

    public void setEmployeeType(String employeeType) {
        this.employeeType = employeeType;
    }

    public EmployeeIncentiveInput() {
        this.employeeID = "";
        this.employeeName = "";
        this.region = "";
        this.jobCode = "";
        this.jobProfile = "";
        this.businessTitle = "";
        this.hireDate = null;
        this.timeInJob = null;
        this.incentivePlan = null;
        this.department = "";
        this.location = "";
        this.sortCriteria1 = null;
        this.sortCriteria2 = null;
        this.sortCriteria3 = null;
        this.employeeType = "";
    }

    public EmployeeIncentiveInput(String employeeID, String employeeName, String region, String jobCode, String jobProfile,
                                  String businessTitle, Date hireDate, Date timeInJob, String incentivePlan, String department, String location, String sortCriteria1, String sortCriteria2, String sortCriteria3, String employeeType) {
        this.employeeID = employeeID;
        this.employeeName = employeeName;
        this.region = region;
        this.jobCode = jobCode;
        this.jobProfile = jobProfile;
        this.businessTitle = businessTitle;
        this.hireDate = hireDate;
        this.timeInJob = timeInJob;
        this.incentivePlan = incentivePlan;
        this.department = department;
        this.location = location;
        this.sortCriteria1 = sortCriteria1;
        this.sortCriteria2 = sortCriteria2;
        this.sortCriteria3 = sortCriteria3;
        this.employeeType = employeeType;
    }

    public String getEmployeeID() {
        return employeeID;
    }

    public void setEmployeeID(String employeeID) {
        this.employeeID = employeeID;
    }

    public String getEmployeeName() {
        return employeeName;
    }

    public void setEmployeeName(String employeeName) {
        this.employeeName = employeeName;
    }

    public String getRegion() {
        return region;
    }

    public void setRegion(String region) {
        this.region = region;
    }

    public String getJobCode() {
        return jobCode;
    }

    public void setJobCode(String jobCode) {
        this.jobCode = jobCode;
    }

    public String getJobProfile() {
        return jobProfile;
    }

    public void setJobProfile(String jobProfile) {
        this.jobProfile = jobProfile;
    }

    public String getBusinessTitle() {
        return businessTitle;
    }

    public void setBusinessTitle(String businessTitle) {
        this.businessTitle = businessTitle;
    }

    public Date getHireDate() {
        return hireDate;
    }

    public void setHireDate(Date hireDate) {
        this.hireDate = hireDate;
    }

    public Date getTimeInJob() {
        return timeInJob;
    }

    public void setTimeInJob(Date timeInJob) {
        this.timeInJob = timeInJob;
    }

    public String getIncentivePlan() {
        return incentivePlan;
    }

    public void setIncentivePlan(String incentivePlan) {
        this.incentivePlan = incentivePlan;
    }

    public String getDepartment() {
        return department;
    }

    public void setDepartment(String department) {
        this.department = department;
    }

    public String getLocation() {
        return location;
    }

    public void setLocation(String location) {
        this.location = location;
    }

    public void setSortCriteria1(String sortCriteria) {
        this.sortCriteria1 = sortCriteria;
    }

    public void setSortCriteria2(String sortCriteria) {
        this.sortCriteria2 = sortCriteria;
    }

    public void setSortCriteria3(String sortCriteria) {
        this.sortCriteria3 = sortCriteria;
    }

    @Override
    public String toString() {

        String sortValue = "";
        String[] getSortValuesArray;

        if (!sortCriteria1.isEmpty()) {
            getSortValuesArray = GetSortValues(sortCriteria1);
            sortValue = "Employee [" + getSortValuesArray[0] + " = " +
                    getSortValuesArray[1];
            if (!sortCriteria2.isEmpty()) {
                getSortValuesArray = GetSortValues(sortCriteria2);
                sortValue = sortValue + " - " + getSortValuesArray[0] + " = " +
                        getSortValuesArray[1];
                if (!sortCriteria3.isEmpty()) {
                    getSortValuesArray = GetSortValues(sortCriteria2);
                    sortValue = sortValue + " - " + getSortValuesArray[0] + " = " +
                            getSortValuesArray[2];
                }
            }
            sortValue = sortValue + "]";
        }

        return sortValue;

    }

    private String[] GetSortValues(String searchCriteria) {

        String[] sortValue = new String[2];

        switch (searchCriteria) {
            case "employeID":
                sortValue[0] = StringUtils.rightPad("Employee ID", 50);
                sortValue[1] = StringUtils.rightPad(getEmployeeID(), 50);
                break;
            case "employeeName":
                sortValue[0] = StringUtils.rightPad("Employee Name", 50);
                sortValue[1] = StringUtils.rightPad(getEmployeeName(), 50);
                break;
            case "jobCode":
                sortValue[0] = StringUtils.rightPad("Job Code", 50);
                sortValue[1] = StringUtils.rightPad(getJobCode(), 50);
                break;
            case "jobProfile":
                sortValue[0] = StringUtils.rightPad("Job Profile", 50);
                sortValue[1] = StringUtils.rightPad(getJobProfile(), 50);
                break;
            case "businessTitle":
                sortValue[0] = StringUtils.rightPad("Business Title", 50);
                sortValue[1] = StringUtils.rightPad(getBusinessTitle(), 50);
                break;
            case "hireDate":
                sortValue[0] = StringUtils.rightPad("Hire Date", 50);
                sortValue[1] = StringUtils.rightPad(getHireDate().toString(), 50);
                break;
            case "timeInJob":
                sortValue[0] = StringUtils.rightPad("Time in Job", 50);
                sortValue[1] = StringUtils.rightPad(getTimeInJob().toString(), 50);
                break;
            case "incentivePlan":
                sortValue[0] = StringUtils.rightPad("Incentive Plan", 50);
                sortValue[1] = StringUtils.rightPad(getIncentivePlan(), 50);
                break;
            case "department":
                sortValue[0] = StringUtils.rightPad("Department", 50);
                sortValue[1] = StringUtils.rightPad(getDepartment(), 50);
                break;
            case "location":
                sortValue[0] = StringUtils.rightPad("Location", 50);
                sortValue[1] = StringUtils.rightPad(getLocation(), 50);
                break;
            default:
                sortValue[0] = StringUtils.rightPad("Region", 50);
                sortValue[1] = StringUtils.rightPad(getRegion(), 50);
        }

        return sortValue;
    }

    @Override
    public int compareTo(EmployeeIncentiveInput employeeIncentiveInput) {

        int i;

        switch (employeeIncentiveInput.getEmployeeType()) {
            case "Field":
                i = this.getRegion().compareTo(employeeIncentiveInput.getRegion());
                if (i != 0) return i;
                i = this.getLocation().compareTo(employeeIncentiveInput.getLocation());
                if (i != 0) return i;
                i = this.getIncentivePlan().compareTo(employeeIncentiveInput.getIncentivePlan());
                if (i != 0) return i;
                return this.getJobProfile().compareTo(employeeIncentiveInput.getJobProfile());
            default:
                i = this.getIncentivePlan().compareTo(employeeIncentiveInput.getIncentivePlan());
                if (i != 0) return i;
                i = this.getDepartment().compareTo(employeeIncentiveInput.getDepartment());
                if (i != 0) return i;
                return this.getRegion().compareTo(employeeIncentiveInput.getRegion());
        }
    }
}


