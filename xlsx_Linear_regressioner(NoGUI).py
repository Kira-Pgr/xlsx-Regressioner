import numpy as np
import matplotlib.pyplot as plt
from scipy import stats
import pandas as pd

# Read data from Excel file
# Ensure the Excel file is in the same directory as your Python script, or provide the full path
df = pd.read_excel('test.xlsx')  

# Extract x and y values from the DataFrame
x = df['t'].values
y = df['g'].values
x = x[~np.isnan(x)]
y = y[~np.isnan(y)]
# Perform linear regression
slope, intercept, r, p, std_err = stats.linregress(x, y)

# Define linear regression function
def myfunc(x):
    return slope * x + intercept

# Generate model predictions
mymodel = list(map(myfunc, x))

equation = f"y = {slope:.5f}x + {intercept:.5f}"
# Plot the data and the linear regression line
plt.scatter(x, y, label='Data')
plt.plot(x, mymodel, color='red', label='Linear Regression')
plt.xlabel('x')
plt.ylabel('y')
plt.title('Linear Data and Linear Regression')
plt.text(0.95, 0.05, equation, transform=plt.gca().transAxes, 
         verticalalignment='bottom', horizontalalignment='right', 
         fontsize=10, bbox=dict(boxstyle="round", alpha=0.5, facecolor="white"))
plt.legend()

plt.grid(True)
plt.show()
