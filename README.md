# Whole-Enrichment
Complete enrichment analysis, by sample, cell type and organ

## Table of Contents
## LocalSetup
1) Install dependencies
'pip3 install -r requirements.txt'

2) Run file either:
	
	a) (Optional) Create an executable file and run file through it
	* 'pip3 install pyinstaller'
	* 'pyinstaller --hidden-import=cmath --onefile -w Enrichment_interface.py'

	If issue arise, check https://www.pyinstaller.org/
	
	Additional resource: https://www.youtube.com/watch?v=t51bT7WbeCM

	b) Run file on terminal
	* 'python3 Enrichment_interface.py'

## ToDo
1) Back-end:
	* removal of run-aways
	* removal of outlier mice
	* renormalization

2) front-end:
  	* select if user wants outliers removed
  	* print out outliers removed (barcodes and mice sample names)
