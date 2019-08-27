import java.util.Arrays;
import java.util.Comparator;
import java.util.List;

/**
 * This is a chained comparator that is used to sort a list by multiple
 * attributes by chaining a sequence of comparators of individual fields
 * together.  This comparator is specifically for the GoalColumnDefinition Collection.
 */

public class GoalColumnChainedComparator implements Comparator<GoalColumnDefinition> {

    private List<Comparator<GoalColumnDefinition>> listComparators;

    @SafeVarargs
    public GoalColumnChainedComparator(Comparator<GoalColumnDefinition>... comparators) {
        this.listComparators = Arrays.asList(comparators);
    }

    @Override
    public int compare(GoalColumnDefinition goalColumnDefinition1, GoalColumnDefinition goalColumnDefinition2) {
        for (Comparator<GoalColumnDefinition> comparator : listComparators) {
            int result = comparator.compare(goalColumnDefinition1, goalColumnDefinition2);
            if (result != 0) {
                return result;
            }
        }
        return 0;
    }

}