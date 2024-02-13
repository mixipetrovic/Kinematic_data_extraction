# Kinematic_data_extraction
Extracting the data given in Excel dataframe

When getting tabular data with each sheet containing several columns of variables, and each of the sheets is one repetition
(recording of the movement).

First thing you want to do is to generate another sheet that will calculate mean across all sheets and log them in 'mean'
sheet.

Second thing is to transpose the data and store them in 'transposed_mean' sheet which
them is useful in feeding the data to some pre-made scripts in Matlab/Python.

After that, you want to iterate through all files in your working directory, extract the 'transposed_mean' sheet,
and store it in variable_data Excel file so that each variable goes in it's designated sheet. 

The data in your new variable_data Excel file is stored so that each sheet is a different variable containing data from
all iterated files...