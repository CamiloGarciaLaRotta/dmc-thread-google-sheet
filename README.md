# DMC Thread Google Sheet
A Google Sheet to keep inventory of DMC threads. It will automatically colour cells and add the colour name for any valid DMC code.

### How to use
1. Create a copy of [the base Google Sheet](https://docs.google.com/spreadsheets/d/11LGlq5iOEaUjL3uIchl868p0G1mrW2Gz5RfZK5EjYhM/edit?usp=sharing): `File -> Make a copy`.  
  The default layout has a column per colour group, but you can aggregate all colours in a single column if you want.
2. Add a valid DMC code in any of the `Number` columns. If you added it to the cell `A3`, then `A3` will be coloured with the appropriate colour and `B3` will be filled up with the DMC name.

![how to use](https://user-images.githubusercontent.com/17187770/103323734-50c6ba80-4a12-11eb-9d9d-156752b7926c.gif)

### How to modify the code
Click on `Tools -> Script Editor` to open the code. From here you can modify:
- the name of the sheets that contain the colours and the sheet that contains the codes
- the columns that will be considered to be colored over
- the columns that define the DMC code, name and HEX codes. By default, these codes live in the `colour_codes` sheet, feel free to modify it as needed.

![dmcedit](https://user-images.githubusercontent.com/17187770/103323846-af8c3400-4a12-11eb-9b95-2b3fcd671d19.gif)


♥️ Shoutout to @adrianj, as I took the DMC-to-HEX mapping from his [`CrossStitchCreator`](https://github.com/adrianj/CrossStitchCreator/blob/master/CrossStitchCreator/Resources/DMC%20Cotton%20Floss%20converted%20to%20RGB%20Values.csv) project.

