public class GoalDefinition {

    private Integer fiscalYear;
    private String incentivePlanName;
    private String planCategory;
    private Float max;
    private Float target;
    private Float min;
    private String goalName;
    private Float goalWeight;
    private Float flatAmount;

    public GoalDefinition() {
        this.fiscalYear = 0;
        this.incentivePlanName = "";
        this.planCategory = "";
        this.max = 0f;
        this.target = 0f;
        this.min = 0f;
        this.goalName = "";
        this.goalWeight = 0f;
        this.flatAmount = 0f;
    }

    public GoalDefinition(Integer fiscalYear, String incentivePlanName, String planCategory, Float max, Float target, Float min, String goalName, Float goalWeight, Float flatAmount) {
        this.fiscalYear = fiscalYear;
        this.incentivePlanName = incentivePlanName;
        this.planCategory = planCategory;
        this.max = max;
        this.target = target;
        this.min = min;
        this.goalName = goalName;
        this.goalWeight = goalWeight;
        this.flatAmount = flatAmount;
    }

    public Integer getFiscalYear() {
        return fiscalYear;
    }

    public void setFiscalYear(Integer fiscalYear) {
        this.fiscalYear = fiscalYear;
    }

    public String getIncentivePlanName() {
        return incentivePlanName;
    }

    public void setIncentivePlanName(String incentivePlanName) {
        this.incentivePlanName = incentivePlanName;
    }

    public String getPlanCategory() {
        return planCategory;
    }

    public void setPlanCategory(String planCategory) {
        this.planCategory = planCategory;
    }

    public Float getMax() {
        return max;
    }

    public void setMax(Float max) {
        this.max = max;
    }

    public Float getTarget() {
        return target;
    }

    public void setTarget(Float target) {
        this.target = target;
    }

    public Float getMin() {
        return min;
    }

    public void setMin(Float min) {
        this.min = min;
    }

    public String getGoalName() {
        return goalName;
    }

    public void setGoalName(String goalName) {
        this.goalName = goalName;
    }

    public Float getGoalWeight() {
        return goalWeight;
    }

    public void setGoalWeight(Float goalWeight) {
        this.goalWeight = goalWeight;
    }

    public Float getFlatAmount() {
        return flatAmount;
    }

    public void setFlatAmount(Float flatAmount) {
        this.flatAmount = flatAmount;
    }
}
