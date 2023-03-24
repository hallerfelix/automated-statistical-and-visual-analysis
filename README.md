# automated-statistical-and-visual-analysis
This Python code performs statistical analysis on an Excel sheet as follows:

1. Select an Excel sheet and choose a group column through a user-friendly window interface

![image](https://user-images.githubusercontent.com/80318329/227542366-6d76a53a-372c-4de1-a005-31ed49919c2e.png)
![image](https://user-images.githubusercontent.com/80318329/227542527-d5911440-a458-4c13-825d-4e53cc367770.png)
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
 4. Create a new folder called Plots and plot a Boxplot of each column with depending p-values.
 
 ![image](https://user-images.githubusercontent.com/80318329/227545430-9bbf821f-4df2-45e4-b2d4-2a6a66b9fbfd.png)
   - if you run the code in the terminal you will get this output:
  
  ![image](https://user-images.githubusercontent.com/80318329/227543697-c21911ef-d564-449b-a3c1-4d6a7d717ae3.png)
