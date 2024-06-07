import pandas as pd
import numpy as np
from tkinter import Tk, filedialog
import matplotlib.pyplot as plt

# Calculate control limits (upper, mean, lower) based on standard deviation (Sd)
def limits(vals, Sd):
    lm1 = np.mean(vals[0])        # Mean of the first set of values (could be range, variance, or std)
    ls1 = lm1 + 3 * Sd            # Upper control limit
    li1 = lm1 - 3 * Sd            # Lower control limit
    if li1 < 0:
        li1 = 0                   # Ensure lower control limit is not negative
    L1=[ls1,lm1,li1]

    lm2 = np.mean(vals[1])        # Mean of the second set of values (Xmean
    ls2 = lm2 + 3 * Sd            # Upper control limit
    li2 = lm2 - 3 * Sd            # Lower control limit
    L2 = [ls2, lm2, li2]

    return(L1,L2)

# Compute standard deviation and mean for each sample in the data
def std_values(dados):
    s = []          # List to store standard deviations
    xmean = []      # List to store means
    i = 0
    while i < len(dados):
        s.append(np.std(dados[i]))
        xmean.append(np.mean(dados[i]))
        i = i + 1
    return s, xmean

# Compute variance and mean for each sample in the data
def var_values(dados):
    v = []          # List to store variances
    xmean = []      # List to store means
    i = 0
    while i < len(dados):
        v.append(np.var(dados.iloc[i]))
        xmean.append(np.mean(dados.iloc[i]))
        i=i+1
    return (v,xmean)

# Compute range and mean for each sample in the data
def amplitude_values(dados):
    r = []          # List to store ranges
    xmean = []      # List to store means
    i=0
    while i < len(dados):
        r.append(np.max(dados.iloc[i]) - np.min(dados.iloc[i]))
        xmean.append(np.mean(dados.iloc[i]))
        i=i+1
    print(r)
    print(xmean)
    return (r,xmean)


# Call the appropriate method based on the choice (amplitude, var, or std) and store the values
def values(Choice, Dados):
    if Choice == 'amplitude':
        val = amplitude_values(Dados)
    elif Choice == 'var':
        val = var_values(Dados)
    else:
        val = std_values(Dados)
    return val



# Determine the appropriate graph type based on the characteristics of the data
def grafchoice(DD):
    data = np.array(DD)
    if np.isnan(data).any() == True:
        choice = 'std'            # Choose 'std' if any NaN values are present
    else:
        if len(DD.iloc[0]) < 10:
            choice = 'amplitude'  # Choose 'amplitude' if sample size is less than 10
        else:
            choice = 'var'        # Choose 'var' for larger sample sizes
    return choice

# Plot the control charts based on the type, data, and control limits
def PlotGrafr (type, data, limit):
    plt.figure(figsize=(10, 6))

    # Plot x̄ control chart
    plt.subplot(211)
    plt.plot(data[1], label='Dados', color='blue')
    plt.scatter(range(len(data[1])), data[1], color='blue')
    plt.axhline(limit[1][1], color='orange', linestyle='--', label='Mean')
    plt.axhline(limit[1][0], color='red', linestyle='--', label='Control\n Limits')
    plt.axhline(limit[1][2], color='red', linestyle='--')
    plt.title('x̄ control chart')
    plt.legend(loc='upper left', bbox_to_anchor=(1, 0.5))
    plt.grid(True)

    # Plot control chart for range, variance, or std
    plt.subplot(212)
    plt.plot(data[0], label='Dados', color='blue')
    plt.scatter(range(len(data[0])), data[0], color='blue')
    plt.axhline(limit[0][1], color='orange', linestyle='--', label='Mean')
    plt.axhline(limit[0][0], color='red', linestyle='--', label='Upper Control Limit')
    plt.axhline(limit[0][2], color='red', linestyle='--', label='Lower Control Limit')
    if type == 'amplitude':
        plt.title('R Control Chart')
    elif type == 'var':
        plt.title('S² Control Chart')
    else:
        plt.title('Sd Control Chart')
    plt.grid(True)
    plt.show()

# Open file dialog to select an Excel file and return its path
def select_sheet():
    root = Tk()
    root.withdraw()
    sheet = filedialog.askopenfilename( title="Choose the datasheet",filetypes=[("Excel", "*.xlsx")])
    root.destroy()
    return sheet

# Main execution starts here
dd = pd.read_excel(select_sheet())   # Read data from the selected Excel file
graph_type = grafchoice(dd)          # Determine the type of graph to plot

if graph_type == 'std':
    total_sd = dd.values.flatten()          # Flatten the DataFrame to a 1D array
    total_sd = total_sd[~pd.isna(total_sd)] # Remove NaN values
    total_sd = total_sd.std()               # Calculate the standard deviation of the entire dataset
    i = 0
    dd = np.array(dd)

    # Clean data samples of NaN values
    cdata = []
    for _ in dd:
        x = dd[i]
        x = x[~np.isnan(x)]
        cdata.append(x)
        i = i + 1
    dd = cdata
else:
    total_sd = dd.values.flatten().std()    # Calculate the standard deviation for non-std types
val = values(graph_type,dd)                 # Compute the required values (amplitude, variance, or std)
lim = limits(val,total_sd)                  # Calculate control limits
n=PlotGrafr(graph_type, val, lim)           # Plot the control charts
