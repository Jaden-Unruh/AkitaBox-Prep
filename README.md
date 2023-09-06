# AkitaBox Preparing Tool
A tool to prepare data for upload to AkitaBox
## Setup
Java SE 17 is required to run this program. If you don't have Java 17 or a newer version installed, you can download an installer for Temurin/OpenJDK 17 from [here](https://github.com/adoptium/temurin17-binaries/releases/download/jdk-17.0.8%2B7/OpenJDK17U-jdk_x64_windows_hotspot_17.0.8_7.msi). This is an open-source version of java. Once downloaded, you can run the installer by double-clicking, it will open a window guiding you through the installation. Leaving everything as the defaults and just clicking through the pages should work perfectly.

The JAR for the program itself is located within the folder [target](https://github.com/Jaden-Unruh/AkitaBox-Prep/tree/main/target) in the GitHub repository. It is the only `.jar` there, it should be called something like `AkitaBox-Prep-1.0.x-jar-with-dependencies.jar`. The reason I don't link to it directly is because the file will change as I update the program. Click on the name of the file there, and click the download button (an arrow pointing downwards towards a tray) in the top-right. The button will say "Download raw file" when you hover over it. You can rename the file to whatever you'd like after it's downloaded.\

Once Temurin/Java 17 and the program .jar are installed, double click the `.jar` to run.
## GUI and How to Use
After double-clicking the `.jar`, a window titled "AkitaBox Prep Tool" will open. It will have 2 prompts, as described below:
1. `Select Component Inventory Sheet`
	* Click on the select button to open a file prompt, navigate to and select the Component Inventory Spreadsheet (`ASPxGridViewInventory.xlsx`). Note that this must be a `*.xlsx` file, rather than `*.xlsb` or any other spreadsheet filetype - see [Troubleshooting](/#troubleshooting) for more. The contents of this spreadsheet should be as described under [Files->Component Inventory](/#component-inventory).
2. `Select CA Sheet`
	* Click on the select button to open a file prompt, navigate to and select the CA Spreadsheet (`CA-20YY-MM-DD_IAXXX_NXX-XX-00.xlsx`)[^1]. Again, this file has to be a `*.xlsx`. The contents of this spreadsheet should be as described under [Files->CA](/#ca).

The other contents of this window are the `Close` and `Run` buttons, which are fairly self-explanatory, and an info text box. This info box will not be visible when the window is first opened, but will display relevant information as the program runs.
	
[^1]: I'm not sure of this file's actual name, I'm just referring to it as the CA spreadsheet because the file name of the spreadsheet I was given to make this program started with 'CA'. If you'd like to change what shows in the GUI to reflect a different name, you can edit the file `messages.properties` in the `.jar` located at `*.jar\us\akana\tools\AkitaBoxPrep\messages.properties` - open it in a text editor, and change the text after `Main.Window.CAPrompt=` to whatever you'd like the new prompt to be. You can do the same for any other user-visible Strings in the program, they are all located in this file. Note that you will have to uncompress the `.jar` before you do this, and recompress it afterwards.

## Files
The program requires two spreadsheets to run, as follows:
### Component Inventory
This `.xlsx` spreadsheet should have columns lettered A through R. Data should begin on row two, with the IDs used in the temporary asset numbers in column R and the maximo IDs in column I.
### CA
This `.xlsx` spreadsheet should have three sheets - the program only uses the second. This second sheet should have columns lettered A through D. Temporary asset numbers ending in "NEW" should be in column B, and column C will be for the corresponding maximo IDs. It is ok if some are already there, the program will compare what it finds to any that are already there and leave a comment if they are different.

## Troubleshooting
> Nothing's happening when I double click the `.JAR` file

Ensure you've installed Java as specified under [Setup](/#Setup). If you believe you have, try checking your java version:
1. Press Win+R, type `cmd` and press enter - this will open a command prompt window
2. Type `java -version` and press enter
3. If you've installed java as specified, the first line under your typing should read `openjdk version "17.0.8" 2023-07-18`[^2]. If, instead, it says `'java' is not recognized as an internal...` then java is not installed.

[^2]: If you had a version of java other than the one specified in Setup, this may show a different version, but should be similar. However, you probably wouldn't be in this troubleshooting step if this is the case.

---
> I only have spreadsheets of type `*.xlsb` or `*.csv` (or any other spreadsheet type) and the program won't open them

Open the spreadsheets in Microsoft Excel and select 'File -> Save As -> This PC' and choosing 'Excel Workbook (.xlsx)' from the drop-down. A full list of filetypes that Excel supports (and thus can be converted to .xlsx) can be found [here](https://learn.microsoft.com/en-us/deployoffice/compat/office-file-format-reference#file-formats-that-are-supported-in-excel).

---
> `Run` isn't doing anything

Ensure that you've selected two `*.xlsx` files. Spreadsheets of a different type will not work.

---
> I'm getting an error message popping up when I run the file

If you're getting an error message and you can't figure out what it's saying or how to fix it, reach out to me. If you click `More Info` on the error popup and copy the big text box, that text (a full stack trace on the error) can help me figure out what's going on.

My first guess for where an error could arise is if your files don't match the format described in [Files](/#files).

## Details of what it does
First, the program runs some initialization - notably, loading the sheets into memory in a format the rest of the program can read. Then, for each row in the CA spreadsheet, it will pull the text from the second column and confirm it ends in "NEW" (if it doesn't, the program will leave a comment saying such and skip to the next row). It will cut the 'NEW' off the end and search for the remaining string in column R of the Component Inventory spreadsheet. Once it finds it, it will copy the Maximo ID from Column I of that spreadsheet and compare it to column C of the CA spreadsheet. If column C is empty, it will just write the String. If column C already has the same Maximo ID, it will move on, and if column C has a different String it will write a comment explaining this.

If the program can't find the temporary asset id in the Component Inventory spreadsheet, it will write a different comment in column C explaining this.

## Changing the Code
The `.JAR` file is compiled and compressed, meaning all the code is not human-readable. You can decompile and recompile the file to change certain parts, like some of the GUI text, but all of the code itself is not editable. Instead, all of the program files are included in a [github repository](https://github.com/Jaden-Unruh/AkitaBox-Prep) so that anyone other than me could download them and open them in an IDE (I use Eclipse). See [In the GitHub](/#in-the-github) for more.

## In the GitHub
I did this one a little differently from previous projects, to make it easier to work on this as an ongoing project - as I believe it will be, from what my Dad said. I linked the GitHub repository directly to my eclipse workspace files, so what you see here is everything located on my computer within `\eclipse-workspace\AkitaBox-Prep\`. This makes it easy for me to keep this updated as I make changes to the project (and should make things easy if anyone in the future needs to make changes). This `README.md`, and an identical HTML version, `README.html` are located directly in the parent directory.

The compiled JAR, the file that a user will be running to actually use the program, is located within the folder `target`, called something like `AkitaBox-Prep-1.0.x-jar-with-dependencies.jar`. This file can be renamed to anything you like after you download it, and once the program reaches a more final form I may copy this, renamed, to the parent directory. The current location and naming is just the default for Maven so I don't have to rename and move things around every time I recompile after making changes.

The repository also includes javadoc, detailed documentation for all of my java code. This may be useful to anyone trying to change the code in any way. This is located within the folder [doc](https://github.com/Jaden-Unruh/AkitaBox-Prep/tree/main/doc), you can view it by opening `index.html`after downloading the contents.