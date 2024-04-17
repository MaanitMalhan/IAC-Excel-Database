# Excel Data Extraction for the SNE-IAC

### Solving the problem
The data file from the `iac.university` website holds data about the work from all the Industrial Assessment Centers nationwide. This file is essentially useless due to the combination of every center's data. This software filters all pieces of data not from the SNE IAC and creates its own data points and graphs to help the SNE IAC get insight into their work. 

### Usage
To use this software go into `src/app.py` the first variable you see after all the imports on `line 17` will be called `universal_dir` change this variable value to wherever you want the output files saved. Use the `PATH` format for the variable look at the placeholder value to check the format. It is recommended you output them into the already created `files` folder. After setting where the files are saved you just need to run `app.py`. You can see the results in the `files` folder. 

### Process

This is a more detailed breakdown of the process the software goes through and uses in chronological order. 

1. `app.py` creates the Excel file where all the output data is saved this file is called `SNE_IAC_Database.xlsx`

2. The data file containing every center's data is downloaded from the IAC server. This is a ZIP file so it is extracted.

3. The extracted file is in `xls` format. This is an outdated version which is not compatible with most Python APIs so we converted to the newer `xlsx` format.

4. After both the input and output files are ready and prepped the software goes through the entire `assessment` sheet on the converted version of the original file and copies all relevant data to the SNE IAC-specific Excel file. 

5. The last step is repeated for the `recommendations` and `terms` sheets. The `terms` sheet also gets screenshots highlighting the center's work from the official IAC website. 

6. The top rows for sheets are relabeled for easier use.

7. A new sheet containing `ARC codes` is created 8. All major ARC codes are put into this sheet 

8. A `Calculations` sheet. is created to house the new data points created while other relevant calculations are left inside their respective sheets.

9. Calculations are done(example: Number of recommendations, Amount saved per recommendation, etc).

10. The value of these calculations replaces the formula of the respective calculation so we can use the values to generate plots.

11. Plots are generated using data from the SNE IAC Excel file. 

12. All work is saved.  
