
public class GoalColumnDefinition {

    private String columnName;
    private String columnType;
    private String linkedGoalName;
    private Integer columnOrder;
    private Integer columnGoalOrder;
    private boolean columnLocked;
    private String defaultValue;
    private String selectionFile;
    private String regionException1;
    private String regionException2;
    private String regionException3;
    private String regionException4;
    private String regionException5;

    public GoalColumnDefinition(String columnName, String columnType, String linkedGoalName, Integer columnOrder, Integer columnGoalOrder,
                                      boolean columnLocked, String defaultValue, String selectionFile, String regionException1,
                                      String regionException2, String regionException3, String regionException4, String regionException5) {

        this.columnName = columnName;
        this.columnType = columnType;
        this.linkedGoalName = linkedGoalName;
        this.columnOrder = columnOrder;
        this.columnGoalOrder = columnGoalOrder;
        this.columnLocked = columnLocked;
        this.defaultValue = defaultValue;
        this.selectionFile = selectionFile;
        this.regionException1 = regionException1;
        this.regionException2 = regionException2;
        this.regionException3 = regionException3;
        this.regionException4 = regionException4;
        this.regionException5 = regionException5;

    }

    public String getColumnType() {
        return columnType;
    }

    public void setColumnType(String columnType) {
        this.columnType = columnType;
    }

    public String getLinkedGoalName() {
        return linkedGoalName;
    }

    public void setLinkedGoalName(String linkedGoalName) {
        this.linkedGoalName = linkedGoalName;
    }

    public Integer getColumnOrder() {
        return columnOrder;
    }

    public void setColumnOrder(Integer columnOrder) {
        this.columnOrder = columnOrder;
    }

    public Integer getColumnGoalOrder() {
        return columnGoalOrder;
    }

    public void setColumnGoalOrder(Integer columnGoalOrder) {
        this.columnGoalOrder = columnGoalOrder;
    }

    public boolean isColumnLocked() {
        return columnLocked;
    }

    public void setColumnLocked(boolean columnLocked) {
        this.columnLocked = columnLocked;
    }

    public String getDefaultValue() {
        return defaultValue;
    }

    public void setDefaultValue(String defaultValue) {
        this.defaultValue = defaultValue;
    }

    public String getColumnName() {
        return columnName;
    }

    public void setColumnName(String columnName) {
        this.columnName = columnName;
    }

    public String getSelectionFile() {
        return selectionFile;
    }

    public void setSelectionFile(String selectionFile) {
        this.selectionFile = selectionFile;
    }

    public String getRegionException1() {
        return regionException1;
    }

    public void setRegionException1(String regionException1) {
        this.regionException1 = regionException1;
    }

    public String getRegionException2() {
        return regionException2;
    }

    public void setRegionException2(String regionException2) {
        this.regionException2 = regionException2;
    }

    public String getRegionException3() {
        return regionException3;
    }

    public void setRegionException3(String regionException3) {
        this.regionException3 = regionException3;
    }

    public String getRegionException4() {
        return regionException4;
    }

    public void setRegionException4(String regionException4) {
        this.regionException4 = regionException4;
    }

    public String getRegionException5() {
        return regionException5;
    }

    public void setRegionException5(String regionException5) {
        this.regionException5 = regionException5;
    }
}
