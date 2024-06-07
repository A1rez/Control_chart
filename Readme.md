# Control Chart Plotter
This repository contains a Python script that reads data from an Excel file, performs statistical analysis, and plots control charts. The script supports calculating control limits and plotting x̄, R, S², and standard deviation (Sd) control charts based on the characteristics of the data.

## Features
- **Data Import**: Select and import data from an Excel file.
- **Statistical Analysis**: Compute range, variance, or standard deviation for each sample in the dataset.
- **Control Limits Calculation**: Calculate upper, mean, and lower control limits.
- **Control Charts Plotting**: Plot x̄, R, S², and Sd control charts.

## Requirements
- Python 3.x
- pandas
- numpy
- matplotlib
- tkinter

## Installation
1. Clone the repository:

git clone https://github.com/yourusername/control-chart-plotter.git
cd control-chart-plotter

2. Install the required Python packages:

pip install pandas numpy matplotlib

## Usage
1. Run the script:

python control_chart_plotter.py

2. A file dialog will appear. Select the Excel file containing your data.

3. The script will determine the appropriate type of control chart based on your data and plot the charts.

# Script Breakdown
## Functions
limits(vals, Sd): Calculate control limits based on standard deviation.
std_values(dados): Compute standard deviation and mean for each sample.
var_values(dados): Compute variance and mean for each sample.
amplitude_values(dados): Compute range and mean for each sample.
values(Choice, Dados): Call the appropriate method based on the choice and store the values.
grafchoice(DD): Determine the appropriate graph type based on the characteristics of the data.
PlotGrafr(type, data, limit): Plot the control charts based on the type, data, and control limits.
select_sheet(): Open a file dialog to select an Excel file and return its path.

## Main Execution
Import data from an Excel file.
Determine the type of graph to plot.
Clean data samples if necessary.
Compute the required values.
Calculate control limits.
Plot the control charts.

##Example
python

# Example of how to run the script
python control_chart_plotter.py

## Contributing
Contributions are welcome! Please create an issue or pull request if you have any improvements or new features to add.
