# Comparable Company Analysis Generator
CCA Calculator for companies that are not able to be valued effectively by DCF

### For users:
To use exe:
- Click and run the exe file that can be downloaded from https://github.com/sylchw/comparable_company_analysis/releases (packaged with dependencies)

Sample csv file:
- sample_companies.csv

### Step by step guide:
1. If there is a windows smart screen check, just click more info and run anyway
2. Follow instructions on screen. As of this stage, at any point you wish to undo a command, there is no way. Just close and restart the app.
3. For a CSV template, refer to sample_companies.csv. Ideally it should be companies from the same industry or of similar nature

### For developers:
To run:
- Create fresh conda python 3.8 environment
- Install all python dependencies on requirements.txt

To compile to exe:
pyinstaller -F --paths=C:\Users\User\anaconda3\envs\ccacalculator\Lib\site-packages  cca_calculator.py