# automated-statistical-and-visual-analysis
This Python code performs statistical analysis on an Excel sheet as follows:

1. Select an Excel sheet and choose a group column through a user-friendly window interface
2. Test the normal distribution of the data using appropriate statistical tests
3. If the data is normally distributed:
  - Test for variance using Levene's test
    - If the variances are equal:
    - Perform a one-way ANOVA to test for differences in means among the groups
    - Conduct a Tukey post-hoc test to identify which groups are significantly different from each other
  - If the variances are unequal:
    - Conduct a Welch's ANOVA, which is a robust version of one-way ANOVA that can handle unequal variances
    - Perform a Games-Howell post-hoc test to compare groups' means and identify significant differences
  If the data is not normally distributed:
    - Perform a Kruskal-Wallis test, which is a non-parametric alternative to one-way ANOVA for non-normal data
    - Conduct a Mann-Whitney U test with a Sidak multiple comparison correction to identify which pairs of groups are significantly different
 4. Boxplot alongside p-values will be plotted.
