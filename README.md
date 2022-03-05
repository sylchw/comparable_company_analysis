# Comparable Company Analysis Generator
CCA Calculator for companies that are not able to be valued effectively by DCF

To run:
- Create fresh conda python 3.8 environment
- Install all python dependencies on requirements.txt

To use exe:
- Click and run the exe file inside CCA_Calculator/dist (packaged with dependencies)

Sample csv file:
- sample_companies.csv

To compile to exe:
pyinstaller -F --paths=C:\Users\User\anaconda3\envs\ccacalculator\Lib\site-packages  cca_calculator.py