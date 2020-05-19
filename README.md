# AUTOMATION TRAINING MATERIAL
# Internship Project

<h3>Acknowledgements</h3>
I would like to express my Thanks and gratitude to Mr. Anand Pitchaikani, Managing Director, Modelon Engineering Private limited for giving me an opportunity and guiding me during this Internship period, giving me the valuable time and inputs to complete my project.
I express my deepest Thanks to Mr. Madhan Kumar, Simulation Engineer, Modelon Engineering Private limited for guiding me in completing the tasks performed in this period and completing the project.
And additional Thanks to Mr. Abhilash Kumar Nair, Simulation Engineer and Mr. Sattenapalli Hemanth, Software Engineer, Modelon Engineering Private limited for supports in this period.
Finally, I would like to Thanks my parents and my friends for their encouragement to complete this internship successfully.
</br>
<h2>Introduction</h2>
A python and python based graphical interface build using PyQt5 to reduce manual steps and make entire workflow as automation that is if user provides all the necessary inputs then this tool will handle all other process in background using developed python scripts.
all the input files (*.ppt and *.doc) provided by the user needs to be converted as *.pdf format and according to user’s requirement those pdfs needs to be combined as single file.
<p align="center">
  <img  width="660" height="500" src="https://github.com/ak224001/Automation-training-material/blob/master/Images/AutomationGUI.png">
</p>
<h3> Working Process</h3>
1. Remove all the page Number from ppt files and doc files.</br>
2. convert all .doc, .docx, .ppt, .pptx files into pdfs and using one Lecture(*ppt->*pdf) and one workshop(*.docs->pdf) create one sessions(pdf's) and finally merge all the sessions into one final pdf file.</br>
3.	Add watermark to every pdf page if required.</br>
4.	Add blank page if Sessions(pdf's) ends with odd number of pages.</br>	
5.	Add page number to final pdf.
<h3>Drag and Drop files in GUI</h3>
<p align="center">
  <img  width="660" height="500" src="https://github.com/ak224001/Automation-training-material/blob/master/Images/DragNDrop.png">
</p>
<h3>Purpose of this project</h3>
<h4>No data stolen problem</h4>It is possible to do all this online but Data,Contents,idea might be stolen. </br>
<h4>Saving time</h4>
Number of .ppts,.docs files is approx 40-50.</br>
1.	Manually Converting all Lecture.ppt file and Workshop.doc file into pdf after that merging one Lecture.pdf and Workshop.pdf into one session.
                                                                                            ---approximate time: 10 minutes</br>
2.	Checking all Sessions if it ended with odd number of pages add one blank page.
                                                                                            ---approximate time: 5 minutes</br>
3.	Writing page number on every pdf page (400 pages)
                                                                                            ---approximate time: 5 minutes</br>
4.	Adding watermark to every page.
                                                                                          ----approximate time: 5 minutes</br>
<b>Previously time spend for this work is 25 minute approx., but using this project it reduced to 3-4 mins.</b>
<h3>Challenges</h3>
<b>Python Graphical user interface GUI not responding.</b>
•	Gui was not responding When running multiple script using python subprocess because there was only one thread that is main thread.</br>
•	Using QThread solved this problem it creates a new thread.
<h3>Prerequisites</h3>
•	Python 3.7 or latest version of python.</br>
•	Activated PowerPoint and Word app.
<h3>Python libraries and packages</h3>
•	PyQt5: For creating python Graphical User Interface.</br>
•	Comtypes: To convert all .ppts, .pptxs, .docs, .docxs files into pdf.</br>
•	PyPDF2: A Pure-Python library built as a PDF toolkit.</br>
•	Reportlab: An Open Source Python library for generating PDFs and graphics.</br>
•	Pdfrw: Python library and utility that reads and writes PDF files.</br>
<h3>Warning message</h3>
<p align="center">
  <img  width="660" height="500" src="https://github.com/ak224001/Automation-training-material/blob/master/Images/warningMessage.png">
</p>
<h3>Two page per sheet</h3>
<p align="center">
  <img   src="https://github.com/ak224001/Automation-training-material/blob/master/Images/TwoPagePerSlide.png">
</p>
<h3>Add page Number and blank page</h3>
<kbd>
<p align="center">
  <img src="https://github.com/ak224001/Automation-training-material/blob/master/Images/addBlankpage.png" >
</p>
  </kbd>
<h3>Watermark</h3>
<p align="center">
  <img src="https://github.com/ak224001/Automation-training-material/blob/master/Images/addwatermark.png" >
</p>
<h3>Watermark</h3>
<p align="center">
  <img src="https://github.com/ak224001/Automation-training-material/blob/master/Images/watermark1.png" >
</p>
<h3>Gui With watermark input</h3>
<p align="center">
  <img  width="660" height="500" src="https://github.com/ak224001/Automation-training-material/blob/master/Images/watermarkinput.png">
</p>
<h3>Table of contents</h3>
<p align="center">
  <img src="https://github.com/ak224001/Automation-training-material/blob/master/Images/TableOfContents.png" >
</p>
<h3>Thank you</h3>
