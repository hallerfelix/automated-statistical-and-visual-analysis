# import importlib 
import importlib

# A list of packages to be imported
packages = ["PyQt5", "pandas", "numpy", "scipy", "seaborn", "matplotlib", "pingouin", "openpyxl"]

# Iterating over the list of packages
for package in packages:
    try:
        # Attempting to import the package
        importlib.import_module(package)
    except ModuleNotFoundError:
        # If import fails due to FileNotFoundError, then install the package
        # using the subprocess library
        import subprocess
        subprocess.run(["pip", "install", package])

import os
import sys
from PyQt5.QtWidgets import QApplication, QFileDialog, QMainWindow, QPushButton, QVBoxLayout, QWidget, QListWidget, QLabel
from PyQt5.QtGui import QPalette, QColor, QFont
from PyQt5.QtCore import Qt
from pandas import read_excel, core
from numpy import array, around, percentile
from scipy.stats import shapiro, levene
from seaborn import boxplot, stripplot
from matplotlib import pyplot as plt
from pingouin import anova, welch_anova, kruskal, pairwise_tukey, pairwise_gameshowell, ttest, mwu, pairwise_tests
from warnings import filterwarnings
from os import path, chdir, makedirs, getcwd

filterwarnings("ignore", module='pingouin',message='Not prepending group keys to the result index of transform-like apply')

# Define ASVA class
class ASVA:
    # constructor method
    def __init__(self, data:core.frame.DataFrame, group_column:str):
        # store input data and group column as class attributes
        self.data = data
        self.group_column = group_column
        # get the unique values of the group column and store as class attribute
        self.group_values = sorted(data[group_column].unique())

    # method to test for normal distribution
    def distribution_test(self, column):
        # store column as class attribute
        self.column = column
        # perform Shapiro-Wilk test on data for each group
        norm_p_vals = [shapiro(self.data[self.data[self.group_column] == i][self.column])[1] for i in self.group_values]
        # set significance level
        alpha = 0.05      
        # return "nonparametric" if all p-values are less than alpha
        return "nonparametric" if any(x < alpha for x in norm_p_vals) else "parametric"
        
    # method to test for equal variance
    def variance_test(self, column):
        # store column as class attribute
        self.column = column
        # create list of tuples of data for each group
        variance_list = [tuple(array(self.data[self.data[self.group_column] == g][self.column].tolist())) for g in self.group_values]
        # perform Levene's test on data in variance_list
        variance_p = levene(*variance_list)[1]
        print(variance_p)
        # return "equal_variance" if p-value is greater than or equal to 0.05
        return "equal_variance" if variance_p >= 0.05 else "unequal_variance"
    
    # method to determine which test to run
    def determine_test(self, column):
        print("")
        print("")
        # store column as class attribute
        self.column = column
        if len(self.group_values) > 2:
            # check if data in column is normally distributed
            if self.distribution_test(column=self.column) == "parametric":
                #  check if variances are equal
                if self.variance_test(column=self.column) =="equal_variance":
                    # perform One-Way ANOVA with posthoc Tukey
                    p, pc = self.anova_tukey(column=self.column)
                # if variances are not equal
                elif self.variance_test(column=self.column) == "unequal_variance":
                    # perform Welch's ANOVA with posthoc Tamhane T2
                    p, pc = self.welchs_anova_tamhane(column=self.column)      
            elif self.distribution_test(column=self.column) == "nonparametric":
                # perform Kruskal-Wallis-Test with Mann-Whitney-U-test
                p, pc = self.kruskal_mannwhitneyu(column=self.column)
        elif len(self.group_values) == 2:
            if self.distribution_test(column=self.column) == "parametric":
                p = "not applicable"
                pc = self.t_test(column=self.column)
            elif self.distribution_test(column=self.column) == "nonparametric":
                p = "not applicable"
                pc = self.mannwhinteyu(column=self.column)
        elif len(self.group_values) <= 2:
            raise ValueError ("Two or fewer groups")
        return p, pc
    
    # method to run ANOVA and Tukey's post-hoc test
    def anova_tukey(self, column):
        # Perform a one-way ANOVA on the specified column
        aov = anova(dv=self.column, between=self.group_column, data=self.data)
        # Store the resulting p-value in the object's 'p' attribute
        self.p = aov["p-unc"].values
        # Perform a Tukey's post hoc test on the specified column
        self.pc = pairwise_tukey(dv=self.column, between=self.group_column, data=self.data, effsize="none")
        # Rename the resulting p-value column to 'pval'
        self.pc.rename(columns={"p-tukey":"pval"}, inplace=True)
        # Create a text string containing some information about the results of the tests
        self.text = ["The data is normal distributed with equal variances.\n",
                     "One-Way ANOVA:p={:.3f}\n".format(self.p.item()),
                     "--> Tukey post hoc Test"]
        self.text = "".join(self.text)
        # Print the name of the column being analyzed and the text string
        print(f"## {self.column } ##")
        print(self.text)
        # Return the 'p' and 'pc' attributes
        return self.p, self.pc
    
    # method to run Welch's ANOVA and Tamhane's post-hoc test
    def welchs_anova_tamhane(self, column):
        # Perform a Welch's ANOVA on the specified column
        aov = welch_anova(dv=self.column, between=self.group_column, data=self.data)
        # Store the resulting p-value in the object's 'p' attribute
        self.p = aov["p-unc"].values
        # Perform a Games-Howell post hoc test on the specified column
        self.pc = pairwise_gameshowell(dv=self.column, between=self.group_column, data=self.data, effsize="none")
        # Create a text string containing some information about the results of the tests
        self.text = ["The data is normal distributed with unequal variances.\n",
                     "Welch's ANOVA:p={.3f}\n".format(round(float(self.p.item()), 3)),
                     "--> Games-Howell post hoc Test"]
        self.text = "".join(self.text)
        # Print the name of the column being analyzed and the text string
        print(f"## {self.column} ##")
        print(self.text)
        # Return the 'p' and 'pc' attributes
        return self.p, self.pc

    # method to run Kruskal-Wallis test and Mann-Whitney U post-hoc test
    def kruskal_mannwhitneyu(self, column):
        # Perform a Kruskal-Wallis test on the specified column
        aov = kruskal(dv=self.column, between=self.group_column, data=self.data)
        # Store the resulting p-value in the object's 'p' attribute
        self.p = aov["p-unc"].values
        # Perform Mann-Whitney U tests with a sidak multiple comparison correction
        self.pc = pairwise_tests(data=self.data, dv=self.column, between=self.group_column, parametric=False,
                                 alternative='two-sided', padjust="sidak")
        # Rename the resulting p-value column to 'pval'
        self.pc.rename(columns={"p-corr":"pval"}, inplace=True)
        # Create a text string containing some information about the results of the tests
        self.text = ["The data is not normally distributed.\n",
                    "Kruskal-Wallis-Test:p={:.3f}\n".format(self.p.item()),
                    "--> Mann–Whitney U test with sidak multiple comparison correction"]
        self.text = "".join(self.text)
        # Print the name of the column being analyzed and the text string
        print(f"## {self.column} ##")
        print(self.text)
        # Return the 'p' and 'pc' attributes
        return self.p, self.pc

    def mannwhinteyu(self, column):
        # Set the object's 'p' attribute to 0.01
        self.p = 0.01
        # Perform Mann-Whitney U tests without a multiple comparison correction
        self.pc = pairwise_tests(data=self.data, dv=self.column, between=self.group_column, parametric=False,
                            alternative='two-sided', padjust="none")
        # Rename the resulting p-value column to 'pval'
        self.pc.rename(columns={"p-unc":"pval"}, inplace=True)
        # Create a text string containing some information about the test
        self.text = [f"The data is not normally distributed.\n"
                     f"--> Mann–Whitney U test"]
        self.text = "".join(self.text)
        # Print the name of the column being analyzed and the text string
        print(f"## {self.column} ##")
        print(self.text)
        # Return the 'pc' attribute
        return self.pc

    def t_test(self, column):
        self.p = 0.01
        # Perform t-tests without a multiple comparison correction
        self.pc = pairwise_tests(data=self.data, dv=self.column, between=self.group_column, parametric=True,
                             alternative='two-sided', padjust="none", correction="auto")
        # Rename the resulting p-value column to 'pval'
        self.pc.rename(columns={"p-unc":"pval"}, inplace=True)
        # Create a text string containing some information about the test
        self.text = [f"The data is normally distributed.\n"
                     f"--> t-test"]
        self.text = "".join(self.text)
        # Print the name of the column being analyzed and the text string
        print(f"## {self.column} ##")
        print(self.text)
        # Return the 'pc' attribute
        return self.pc

    def plot_boxes(self):
        # Initialize an empty list called boxes
        self.boxes = []
        # Check the length of group_values and assign values to boxes accordingly
        if len(self.group_values) < 2:
            # If group_values has fewer than 2 elements, raise a ValueError
            raise ValueError ("Too few groups")
        elif len(self.group_values) == 2:
            # If group_values has 2 elements, assign a single list with 4 elements to boxes
            self.boxes = [[0,0,1,1]]
        elif len(self.group_values) == 3:
            # If group_values has 3 elements, assign 2 lists with 4 elements to boxes
            self.boxes = [[0,0,2,2],
                          [0,0,0.9,0.9], [1.1,1.1,2,2]]
        elif len(self.group_values) == 4:
            # If group_values has 4 elements, assign 2 lists with 4 elements and 2 lists with 5 elements to boxes
            self.boxes = [[0,0,3,3],
                          [0,0,1.9,1.9], [2.1,2.1,3,3],
                          [0,0,0.9,0.9], [1.1,1.1,3,3],
                          [1,1,2,2]]
        else:
            # If group_values has more than 4 elements, raise a ValueError
            raise ValueError ("Too many groups")

        # Return the value of boxes
        return self.boxes
    def plot_figure(self, column):
        # Set the size of the figure
        plt.rcParams["figure.figsize"] = 2.5+0.5*len(self.group_values),8+len(self.group_values)*0.5
        # Plot the boxplot
        ax1 = boxplot(x=self.group_column, y=self.column, order=self.group_values, 
                     data=self.data, color="grey", linewidth=2, fliersize=0)
        # Plot the stripplot
        ax1 = stripplot(x=self.group_column, y=self.column, order=self.group_values, 
                     data=self.data, color="white", linewidth=1.5, zorder=1, edgecolor="k", size=18)
        
        # Set the linewidth and fontsize
        linewidth = 2
        fontsize = 25
        # Set the labels
        xlabel = " "
        ylabel = self.column
        title = " "
        
        # Set the y-axis limits
        max_y = self.data[self.column].max()
        min_y = self.data[self.column].min()
        if min_y > 0:
            ylim = [min_y-0.03*( max_y*1.5), max_y*(1 + 0.2*len(self.group_values))]
        elif min_y == 0:
            ylim = [max_y - max_y*1.05, max_y*(1 + 0.2*len(self.group_values))]
        elif min_y < 0:
            raise ValueError ("Values are too small")

        # Modify the plot aesthetics
        ax1.spines['right'].set_visible(False)
        ax1.spines['top'].set_visible(False)
        plt.setp(ax1.spines.values(), linewidth=linewidth)
        plt.xticks(fontsize=fontsize)
        plt.yticks(fontsize=fontsize)
        plt.xlabel(xlabel, fontsize=fontsize)
        plt.ylabel(ylabel, fontsize=fontsize, labelpad=10)
        ax1.tick_params(direction='out', length=4, width=linewidth, colors='k',
                     grid_color='k', grid_alpha=0.5)
        plt.setp(ax1.artists, edgecolor = 'k')
        plt.setp(ax1.lines, color='k')
        ax1.set_ylim(bottom=ylim[0], top=ylim[1])
        ax1.yaxis.set_tick_params(labelsize=fontsize)
        ax1.set_title(title, fontsize=fontsize, pad=15)

        # Set the x-tick labels
        labels = self.group_values
        ax1.set_xticklabels(labels)
        plt.xticks(fontsize=25,rotation=45,ha='right', rotation_mode="anchor")

        #Create a list of rounded integers for each sublist in self.boxes
        integer_box = [[int(round(num)) for num in sublst] for sublst in self.boxes]

        #Loop through the list of integers and assign variables to the pvalues associated with each pair of corresponding A and B groups
        for i, liste in enumerate(integer_box):
            a = self.group_values[liste[0]]
            b = self.group_values[liste[3]]
            variable_name = "row" + str(i)
            globals()[variable_name] = self.pc.loc[(self.pc["A"] == a) & (self.pc["B"] == b)].pval.values
            
        if self.p < 0.05:
            #If two groups are present, draw a line on the graph and add the associated pvalue to the graph
            if len(self.group_values) == 2:
                ax1.plot(self.boxes[0], [0.9*(ylim[1]-ylim[0])+ylim[0],
                                         0.924*(ylim[1]-ylim[0])+ylim[0],
                                         0.924*(ylim[1]-ylim[0])+ylim[0],
                                         0.9*(ylim[1]-ylim[0])+ylim[0]], linewidth= 2, color="k")
                ax1.text(0.15, 0.94*(ylim[1]-ylim[0])+ylim[0],
                         "p={:.3f}".format(row0.item()), fontsize=15)

            #If three groups are present, draw two lines on the graph and add the associated pvalues to the graph
            elif len(self.group_values) == 3:
                ax1.plot(self.boxes[0], [0.900*(ylim[1]-ylim[0])+ylim[0],
                                         0.924*(ylim[1]-ylim[0])+ylim[0],
                                         0.924*(ylim[1]-ylim[0])+ylim[0],
                                         0.900*(ylim[1]-ylim[0])+ylim[0]], linewidth= 2, color="k")
                ax1.text(0.5, 0.94*(ylim[1]-ylim[0])+ylim[0],
                         "p={:.3f}".format(row0.item()), fontsize=15)

                ax1.plot(self.boxes[1], [0.82*(ylim[1]-ylim[0])+ylim[0],
                                         0.844*(ylim[1]-ylim[0])+ylim[0],
                                         0.844*(ylim[1]-ylim[0])+ylim[0],
                                         0.82*(ylim[1]-ylim[0])+ylim[0]], linewidth= 2, color="k")
                ax1.text(0, 0.86*(ylim[1]-ylim[0])+ylim[0],
                         "p={:.3f}".format(row1.item()), fontsize=15)
                ax1.plot(self.boxes[2], [0.82*(ylim[1]-ylim[0])+ylim[0],
                                         0.844*(ylim[1]-ylim[0])+ylim[0],
                                         0.844*(ylim[1]-ylim[0])+ylim[0],
                                         0.82*(ylim[1]-ylim[0])+ylim[0]], linewidth= 2, color="k")
                ax1.text(1.1, 0.86*(ylim[1]-ylim[0])+ylim[0],
                         "p={:.3f}".format(row2.item()), fontsize=15)

            #If four groups are present, draw four lines on the graph and add the associated pvalues to the graph
            elif len(self.group_values) == 4:
                ax1.plot(self.boxes[0], [0.900*(ylim[1]-ylim[0])+ylim[0],
                                         0.924*(ylim[1]-ylim[0])+ylim[0],
                                         0.924*(ylim[1]-ylim[0])+ylim[0],
                                         0.900*(ylim[1]-ylim[0])+ylim[0]], linewidth= 2, color="k")
                ax1.text(1, 0.94*(ylim[1]-ylim[0])+ylim[0],
                         "p={:.3f}".format(row0.item()), fontsize=15)

                ax1.plot(self.boxes[1], [0.820*(ylim[1]-ylim[0])+ylim[0],
                                         0.844*(ylim[1]-ylim[0])+ylim[0],
                                         0.844*(ylim[1]-ylim[0])+ylim[0],
                                         0.820*(ylim[1]-ylim[0])+ylim[0]], linewidth= 2, color="k")
                ax1.text(0.45, 0.86*(ylim[1]-ylim[0])+ylim[0],
                         "p={:.3f}".format(row1.item()), fontsize=15)
                ax1.plot(self.boxes[2], [0.820*(ylim[1]-ylim[0])+ylim[0],
                                         0.844*(ylim[1]-ylim[0])+ylim[0],
                                         0.844*(ylim[1]-ylim[0])+ylim[0],
                                         0.820*(ylim[1]-ylim[0])+ylim[0]], linewidth= 2, color="k")
                ax1.text(2.05, 0.86*(ylim[1]-ylim[0])+ylim[0],
                         "p={:.3f}".format(row2.item()), fontsize=15)

                ax1.plot(self.boxes[3], [0.740*(ylim[1]-ylim[0])+ylim[0],
                                         0.764*(ylim[1]-ylim[0])+ylim[0],
                                         0.764*(ylim[1]-ylim[0])+ylim[0],
                                         0.740*(ylim[1]-ylim[0])+ylim[0]], linewidth= 2, color="k")
                ax1.text(-0.05, 0.78*(ylim[1]-ylim[0])+ylim[0],
                         "p={:.3f}".format(row3.item()), fontsize=15)
                ax1.plot(self.boxes[4], [0.740*(ylim[1]-ylim[0])+ylim[0],
                                         0.764*(ylim[1]-ylim[0])+ylim[0],
                                         0.764*(ylim[1]-ylim[0])+ylim[0],
                                         0.740*(ylim[1]-ylim[0])+ylim[0]], linewidth= 2, color="k")
                ax1.text(1.55, 0.78*(ylim[1]-ylim[0])+ylim[0],
                         "p={:.3f}".format(row4.item()), fontsize=15)

                ax1.plot(self.boxes[5], [0.660*(ylim[1]-ylim[0])+ylim[0],
                                         0.684*(ylim[1]-ylim[0])+ylim[0],
                                         0.684*(ylim[1]-ylim[0])+ylim[0],
                                         0.660*(ylim[1]-ylim[0])+ylim[0]], linewidth= 2, color="k")
                ax1.text(1, 0.70*(ylim[1]-ylim[0])+ylim[0],
                         "p={:.3f}".format(row5.item()), fontsize=15)

        #Create two empty lists to store the 25th percentile and 75th percentile
        twentyfive = []
        seventyfive = []

        #Loop through the different groups and create variables for each dataframe subset by group 
        # and store the 25th percentile and 75th percentile values in the empty lists
        for i, group in enumerate(self.group_values):
            df_variable_name = f"df_{group}"
            globals()[df_variable_name] = self.data[self.data[self.group_column]==group][self.column]
            twentyfive_var_name = f"twentyfive_{group}"
            twentyfive.append(twentyfive_var_name)
            seventyfive_var_name = f"seventyfive_{group}"
            seventyfive.append(seventyfive_var_name)
            globals()[twentyfive_var_name] = percentile(globals()[df_variable_name], 25)
            globals()[seventyfive_var_name] = percentile(globals()[df_variable_name], 75)

        #Loop through the two lists to plot the 25th percentile and 75th percentile lines for each group
        for i, (twfi, sefi) in enumerate(zip(twentyfive, seventyfive)):
            ax1.plot([-0.4+1*i,-0.4+1*i,0.4+1*i, 0.4+1*i],
                     [globals()[twfi],globals()[twfi],globals()[twfi],globals()[twfi]],
                     color="k", linewidth=2, zorder=10)
            ax1.plot([-0.4+1*i,-0.4+1*i,0.4+1*i, 0.4+1*i],
                     [globals()[sefi],globals()[sefi],globals()[sefi],globals()[sefi]],
                     color="k", linewidth=2, zorder=10)

        #Save the plot and show it
        # plt.savefig(f"Plots/{self.column}.jpg", format="jpg",bbox_inches="tight")
        return ax1

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        # Set up the user interface
        self.initUI()

        # Initialize the dataframe and group column
        self.df = None
        self.group_column = None
        
    def closeEvent(self, event):
        # Set the accept parameter to True to close the window
        event.accept()

    def initUI(self):
        # Create a label to display the instruction "Choose the group column"
        label = QLabel("Choose the group column:")

        # Create a button that, when clicked, will open a file dialog
        self.btn = QPushButton('Select Excel sheet', self)
        self.btn.clicked.connect(self.showDialog)

        # Create an "End" button to close the window
        self.end_btn = QPushButton('End', self)
        self.end_btn.clicked.connect(self.close)

        # Create a list widget to display the columns of the imported data
        self.column_list = QListWidget(self)
        self.column_list.itemClicked.connect(self.setGroupColumn)

        # Create a vertical layout to hold the widgets
        layout = QVBoxLayout()
        layout.addWidget(self.btn)  # Add the "Select Excel sheet" button below the label
        layout.addWidget(label)  # Add the label
        layout.addWidget(self.column_list)
        layout.addWidget(self.end_btn)  # Add the "End" button to the bottom of the layout

        # Create a central widget to hold the layout
        central_widget = QWidget()
        central_widget.setLayout(layout)
        self.setCentralWidget(central_widget)

        self.setGeometry(300, 300, 300, 200)
        # self.setWindowTitle('Automated Statistical and Visual Analysis')
        self.setWindowTitle('<html><head/><body><p><span style=" font-family:\'Arial\'; font-size:18pt; font-style:italic; font-weight:600;">Automated Statistical and Visual Analysis</span></p></body></html>')
        self.show()

    def showDialog(self):
        # Open a file dialog to select an Excel sheet
        options = QFileDialog.Options()
        options |= QFileDialog.ReadOnly
        file_name, _ = QFileDialog.getOpenFileName(self,"QFileDialog.getOpenFileName()", "","Excel Files (*.xlsx);;All Files (*)", options=options)
        if file_name:
            # If a file is selected, import it using pandas and display the columns in the list widget
            self.df = read_excel(file_name, index_col="ID")
            self.column_list.clear()
            self.column_list.addItems(self.df.columns)
            
                # Set the directory of the Excel sheet as the current working directory
        file_directory = path.dirname(file_name)
        chdir(file_directory)
        print(getcwd())
        
        # Create the "Plots" folder if it does not exist
        if not path.exists("Plots"):
            makedirs("Plots")
            

    def setGroupColumn(self, item):
        # Set the group column to the selected item
        self.group_column = item.text()
        print(f'Selected column: {self.group_column}')
        
        # Print the values of the selected column if they have not been printed before
        if not hasattr(self, 'values_printed'):
            self.values_printed = True
        
        if self.df[self.group_column].dtypes == "O":
            print(self.df)
            # Define group_column variable
            # Create list of columns, excluding group_column
            columns = [col for col in self.df.columns.tolist() if col != self.group_column]
            # Loop through columns
            for col in columns:
                # Create ASVA object for column
                asva = ASVA(data=self.df, group_column=self.group_column)
                # Determine which test to run
                p, pc = asva.determine_test(column=col)
                box = asva.plot_boxes()
                ax1 = asva.plot_figure(column=col)
                plt.savefig(f"Plots/{col}.jpg", format="jpg",bbox_inches="tight")
                plt.show(block=False)
                plt.close()
                


if __name__ == '__main__':
    # This block is executed only if the script is run directly, rather than imported as a module.
    # Create a QApplication with the command-line arguments passed in sys.argv.
    app = QApplication(sys.argv) 
    # Create an instance of the MainWindow class.
    window = MainWindow() 
    # Show the main window.
    window.show() 
    # Start the event loop and exit when it is finished.
    # The return value is passed to sys.exit() and used as the script's exit status.
    sys.exit(app.exec_()) 
