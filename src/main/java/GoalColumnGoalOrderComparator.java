import java.util.Comparator;

public class GoalColumnGoalOrderComparator implements Comparator<GoalColumnDefinition> {

    public int compare(GoalColumnDefinition goalColumnDef1, GoalColumnDefinition goalColumnDef2) {
        return goalColumnDef1.getColumnGoalOrder() - goalColumnDef2.getColumnGoalOrder();
    }

}
