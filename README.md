# WordScript for BibDesk

A comprehensive suite of AppleScript tools for integrating BibDesk reference manager with Microsoft Word.

## Quick start guide

### Initial set up

WordScript has three parts, store them somewhere convenient, such as the Applications folder.

1. **WordScript.applescript** is the main programme and can be run from the MacOS Script Editor
2. **WordScript Templates** define how the references and bibliography will appear in Word
3. **BibDesk Templates** provide a convenient way to copy and paste cite commands from BibDesk

To set up the BibDesk Templates:

1. Go to **BibDesk > Preferences > Templates** and add the BibDesk Templates
2. Go to **BibDesk > Preferences > Citation** and use the following settings:

> Default Format: **Template**
>
> Template: **BibDesk Template (citep)**
>
> Format when holding Option key: **Template**
>
> Template: **BibDesk Template (citet)**

### First run

**Warning: always save your work before using WordScript!**

1. Make sure you have documents open in both BibDesk and Word
2. Add cite commands to your Word document (see Cheat sheet below)
3. Run WordScript
4. Select 'Format References' and click 'OK'
5. Choose the requested WordScript Template when prompted

**Note:** the first time you use WordScript, Microsoft may ask for permission to store the bibliography template file. The process of granting permission may cause WordScript to fail. **Simply delete the reference list that was created and run WordScript again.**

### Cheat sheet

WordScript uses cite commands, which can be typed or pasted into your Word document.

- The anatomy of a cite command is: ```\cite-type[prefix][suffix]{citekey}```.

- There are four possible cite-types: ```citep```, ```citet```, ```citealp```, ```citealt``` (see below)

- The ```[prefix]``` and ```[suffix]``` can contain any string of text; latin terms can be italicised

- Each ```{citekey}``` must match an item in your BibDesk library, with multiple keys separated by a comma

| citep (parenthetical)                 | Author-YEAR (with parentheses) |
| ------------------------------------- | ------------------------------ |
| ```\citep[][]{citekey}```             | (Author, 2000)                 |
| ```\citep[][]{citekey1,citekey2}```   | (Author1, 2001; Author2, 2002) |
| ```\citep[cf. ][, pp.1–2]{citekey}``` | (*cf.* Author, 2000, pp.1–2)   |

| citet (textual)                       | YEAR (with parentheses) |
| ------------------------------------- | ----------------------- |
| ```\citet[][]{citekey}```             | (2000)                  |
| ```\citet[][]{citekey1,citekey2}```   | (2001; 2002)            |
| ```\citet[cf. ][, pp.1–2]{citekey}``` | (*cf.* 2000, pp.1–2)    |

| citealp (alternative parenthetical)     | Author-YEAR without parentheses |
| --------------------------------------- | ------------------------------- |
| ```\citealp[][]{citekey}```             | Author, 2000                    |
| ```\citealp[][]{citekey1,citekey2}```   | Author1, 2001; Author2, 2002    |
| ```\citealp[cf. ][, pp.1–2]{citekey}``` | *cf.* Author, 2000, pp.1–2      |

| citealt (alternative textual)           | YEAR without parentheses |
| --------------------------------------- | ------------------------ |
| ```\citealt[][]{citekey}```             | 2000                     |
| ```\citealt[][]{citekey1,citekey2}```   | 2001; 2002               |
| ```\citealt[cf. ][, pp.1–2]{citekey}``` | *cf.* 2000, pp.1–2       |

| Combined cite-types                                                  | Complex chaining                               |
| -------------------------------------------------------------------- | ---------------------------------------------- |
| ```\citealp[(][]{citekey1}\citealt[; ][)]{citekey2}```               | (Author, 2001; 2002)                           |
| ```\citealp[(cf. ][, p.1]{citekey1}\citealp[; ][, p.2)]{citekey2}``` | (*cf.* Author1, 2001, p.1; Author2, 2002, p.2) |


