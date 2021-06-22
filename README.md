**UnifyMB**

UnifyMB is the first of three sub-programs that make up CrystalMB, a LIMS system that takes data directly from
instrument output files (in .csv or .xml format), converts them into a unified excel format for editing, pushes the
edited data into a SQL server database, and then draws on that database to make reports using CystalReports.

UnifyMB is the sub-program that converts a variety of different data formats into one unified excel format, that can
also be saved as a .csv file or a .xml file, with field names in the 'A' row, and one analyte concentration per row,
with (n) rows per sample where (n) = the amount of analytes in the particular method being converted.

**How to Use**

- Export data you want to unify from a Waters Instrument, or an Agilent ICP/MS instrument.
    - note: the instruments are currently (22June21) set up to produce data that works with the methods in this program.
    - documentation on what those settings should be for each instrument will be produced at a later date (22June21)

- Put data in the CSVFilesToUnify directory in the overall project folder (CrystalMB)

- Run Unify.TK, select data file to convert. If converted successfully, the file will be in UnifiedExcelFiles.

**Supported Instruments**

as of June 22, 2021:

Supported Instruments:

- Waters Instruments (UPLC-UV, UPLC-MS/MS)

- Agilent Instruments (ICP-MS)

Not Yet Supported Instruments:

- Agilent Instruments (GC-MS)

- Capillary Electrophoresis

- Alex's UV-Vis instrument

**Documentation**

As of June 22nd, 2021:

Fully Documented Files:

- AgilentUnify.py
- UnifyTK.py
- WatersUnify.py

Partially Documented Files:

Undocumented Files:

documentation can be found written into the python files themselves.
Files have help() accessible docstrings as well as comments throughout.

**Contact Information**

This readme and program was written by Peter Levett (MB Laboratories ltd.). This readme
was last updated June 9th, 2021. Any questions or concerns about this program
can be sent to peterlevett@gmail.com.

[Home](http://StavromularBeta.github.io)
