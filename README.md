## LibreOffice Calc template and Basic macro

This repository contains a LO Basic macro and a LO Calc template file which can be used to generate DSpace [SAF](https://wiki.lyrasis.org/display/DSDOC6x/Importing+and+Exporting+Items+via+Simple+Archive+Format) import package.

#### What is this script for?

We are periodically facing the necessity to create DSpace compatible import packages from a given set of text data like CSV or Excel sheets. This need aimed the script to born.

The  process of the import package creation is the following:

* The issuer (aka the customer) has two options: he gives us the metadata in textual format or fills out our LO Calc template.
* Check the data and correct them if needed according to the metadata rules.
* Run the macro.
* Compress the generated folders into a ZIP package and import it via the UI or the console.

The script is able to produce import packages where the items have different metadata schemas. For example, if the items have schema DC and DCTERMS mixed it will then create different XML metadata files according to the DSpace manual. But do not forget to create these metadata entries within the DSpace server before importing the package.

Any number of metadata value can be added to the items. The only one limit is the LibreOffice Calc's limit of the available column number.

#### The basic macro

We chose LibreOffice because it can be freely available for anybody. Unfortunately it is not compatible with the Excel from the Microsoft Office package but it works on the same level as its counterpart.

Installation of the macro:

* Open the macro organization window: _Tools -> Macros -> Organize macros -> Basic..._
* Click on the _Organizer..._ button on the right.
* Within the _Module_ list, highlight the _Standard_ item and then press _New_ on the bottom right.
*  Enter a name for the module. E.g. _Table2SAF_ or such and click on OK.
* Close the _Basic Macro Organizer_ window and highlight the newly created module and press the _Edit_ button on the right.
* Clear all of the lines in the code window and select the menu _File -> Import Basic_.
* Choose the file _calc2saf.bas_ and press _Open_
* Finally save it.

The macro does not use any kind of special classes and it should run from LibreOffice version 6.3 and up without issues.

#### The template

The basic macro can process the LibreOffice Calc template file. This file consist of two sheets:

1. The sheet storing the actual metadata values.

   The first line and the first column contain mandatory data and they should not be moved.

   * The first column must contain some kind of identification data which identifies the different items within the sheet. It can be a file name or an increasing sequence of numbers.
   * The following columns are optional and fully customizable by the configuration parameters. Within these, there are six special columns and the ones with the actual metadata.
   * The first line of the metadata columns must contain the metadata key. E.g. _dc.contributor.author_ . Above this line, you may have additional lines but this metadata line must go before all of the item lines.

2. The sheet for the configuration parameters of the macro.

   We may have items in two ways.

   1. When each item has exactly one attachment (or bitstream with DSpace words). In this case the identification column (column A) must contain the file names and no any real file names column be presented.
   2. When at least one item has two or more different attachments. In this case one column must be filled with the real file names belonging to the items. The very first, identification column must not hold the real file names but any kind of sequence.

   Configurable parameters on the sheet _Config_

   ##### Processing related parameters

   1. Metadata separator character: the script will split the metadata keys using this character. Defaults to '.'
   2. Base folder: the full directory path where the attachments of the items are placed. It is recommended to create a folder next to the metadata table and place the files in them.
   3. Sheet name: the name of the sheet where the metadata are placed.
   4. Metadata row: the line where the metadata keys are found. Recommended the very first line.
   5. First data row: there can be additional lines between the metadata keys line and the actual item lines. E.g. descriptions or explanations may be written for the librarians. This parameters tells the script where to start reading the actual item lines.
   6. Last data column: the last column where data can be found.
   7. Skip these columns: comma separated list column numbers which should be excluded from the processing. These columns are for the process only and not for the items, see below.

   ##### Items related parameters (optional)

   Leave the parameter empty if not in use.

   1. Bundle column: the bundle string. E.g. _ORIGINAL_
   2. Permissions column: the permissions string. Persmissions string may be a name of a group.
   3. Permissions string column: which kind of rights has the group for the item? _w_, _r_, or _r|w_
   4. Description column: this text will be displayed under the bitstream image on the DSpace' UI.
   5. Primary column: _true_ means the bitstream will be the first one when displaying the item and _false_ means the opposite.
   6. File name column: this is the actual file name which belongs to the item. Use it only when there are at least one item having multiple attachments.
   7. Assetstore number column: if this is not empty then the bitstream will be registered for the item as an existing on the specified assetstore. Leave empty if the bitstream can be stored in the default store. Check out [DSpace the manual](https://wiki.lyrasis.org/display/DSDOC5x/Registering+Bitstreams+via+Simple+Archive+Format#RegisteringBitstreamsviaSimpleArchiveFormat-RegisteringItemsUsingtheItemImporter)
   8. Collecation handle: this collection will contain the items.

#### How to import the generated package?

At the end of the generation, there will be different folders under the _base folder_ for each items. The folders will be named based on the identification column of the metadata sheet. Within these folders there will be few files only:

* _contents_ file: this stores the bitstream [information](https://wiki.lyrasis.org/display/DSDOC5x/Importing+and+Exporting+Items+via+Simple+Archive+Format#ImportingandExportingItemsviaSimpleArchiveFormat-DSpaceSimpleArchiveFormat).
* _collections_ file: the collection handle which will contain the items.
* _dublin_core.xml_, _metadata_XXX.xml_: the actual metadata entries in XML format. Each namespace has exactly one XML file.
* the attachments of the item. If the item has attachment which is already stored under an assetstore 1 ... N then this file will not be copied int this directory but the entry will be added into the _contents_ file.

Finally the package my be imported by DSpace in two different ways. It can be packed by the ZIP algorithm and uploaded via the user interface or it can be uploaded directly to the DSpace server (w/o packaging) and [import from the console](https://wiki.lyrasis.org/display/DSDOC5x/Importing+and+Exporting+Items+via+Simple+Archive+Format#ImportingandExportingItemsviaSimpleArchiveFormat-ImportingItems). E.g. `[dspace-dir]/bin/dspace import -add --eperson=uploader@domain.com --collection=ANY-HANDLE --source=<package-path> -m <mapfile-path> --zip=<zip-path-if-any>`

