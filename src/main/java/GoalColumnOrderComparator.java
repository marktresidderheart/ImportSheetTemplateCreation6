import java.util.Comparator;

public class GoalColumnOrderComparator implements Comparator<GoalColumnDefinition> {

    public int compare(GoalColumnDefinition goalColumnDef1, GoalColumnDefinition goalColumnDef2) {
        return goalColumnDef1.getColumnOrder() - goalColumnDef2.getColumnOrder();
    }
}
