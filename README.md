# Drumlin
An Excel based system for modifying EnergyPlus IDF files.

# Installation

The dumlin.xlsm file is the only file you need. You are welcome to rename the file. Make sure you follow the configuration steps shown below. It is working with Excel 2016 but should work with other versions of Excel also. You should download the 5ZoneAirCooled.idf file also to run the macro the first time to understand it. It is the same file as comes as an example with EnergyPlus.

# Configuration

The upper left corner of the spreadsheet that uses Drumlin should look something like this:

![Upper Left Corner](/images/DrumlinExample01.PNG)

The A1 cell always needs to say "Drumlin" just to make sure the macro that is being used is being used on the right sheet. The cells after that are paired with a Key in column A and the corresponding Value in column B. The keys are:

- IDD - The value should be the path to the Energy+.idd file. It should be the same version as the ORIGINAL file. The IDD key should only appear once.
- ORIGINAL - The original idf file that is the source before any changes. The ORIGINAL  key should only appear once.
- REVISED - The revised idf file that is based on the source but has modifications. The revised file gets overwritten every time the macro is run so you should probably not modify that file by hand without copying it to a new name first. The REVISED key should only appear once.
- OBJECT - An object that you want to modify or pull values from. This can be repeated for each object you want shown in the spreadsheet.

# Running the Macro

The macro can be run by clicking on the Run Drumlin button at the top of the sheet.

![Run Drumlin Button](/images/DrumlinExample02-RunDumlin.PNG)

Every time the button is pressed several things happen:

- The Keys and Values described in the Configuration section are used
- The Energy+.idd file is read
- The ORIGINAL file is read and put on the spreadsheet in the sections marked with the name of the object and the word "[ORIGINAL]"
- The REVISED file is written based on the copying the ORIGINAL file and substituting the objects shown in the sections markets with the name of the object and the word "[REVISED]"

# Example

For example, the spreadsheet below the will look like after the Run Drumlin button is pressed:

![Original and Revised Objects](/images/DrumlinExample03-OrigRevObjects.PNG)

The section that says "Lights [ORIGINAL]" is from the ORIGINAL file and shows the values for those fields kind of like the IDF Editor. The section that says "Lights [REVISED]" shows the same values as the original but can be modified and those values will be get written to the REVISED file. 

# Revised Formulas

Most of the cells in the "Lights [REVISED]" section are simple formulas that link to the corresponding rows in the "Lights [ORIGINAL]" section such as "=E11". If special formulas are used those are automatically highlighted. For the Lighting Level row in the example shown above and in the Drumlin.xlsm file has a more complicated formula "=E15*0.95" taking the original Lighting Level value and multiplying it by 0.95. If you look at the  REVISED file, the 5ZoneAirCooled-out.idf in the original spreadsheet. You will see that the lighting levels have all been modified to these new values.

Any cell in the "object-name [REVISED]" blocks can have a new formula or simply a value and this is how simple energy efficiency measures can be implemented.










