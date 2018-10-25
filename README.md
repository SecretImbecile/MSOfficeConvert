# MSOfficeConvert

VBA Batch conversion scripts for legacy Microsoft Office Formats

----

## Instructions

Fistly, copy all of the documents you wish to convert into one folder.

Then, open the Office program for the type of file you want to convert. Create a blank file and save it in the same folder.

Press `ALT+F11` to open *visual basic for applications*

Click `File -> Import File...` and load the .bas file for the corresponding file format.

If you wish to track the progress of the conversion for large numbers of files, press `Ctrl + G` to open a window titled *Immediate*.

Then click `Run -> Run Macro`. Check an function named similarly to *TranslateDocIntoDocx* is present, then press `Run`.

The script will run, printing the file progress in the *immediate* window.