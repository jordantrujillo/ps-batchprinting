# PS-BatchPrinting
Batch Printing with PowerShell

# Updates
I will be updating this code soon. I slapped it together while doing research and it came out very linear and I plan on creating it with functions. 

# Inspiration
I needed something simple to do batch printing, that I could share with end users. I did a bunch of research and couldn't find anything usable for average users that wasn't a costly feature rich commercial product. 

# Supported Platforms
Windows Only

Tested on win10 64bit and Windows Server 2019 64bit. 

# How does it work? 
It relies on the default printer, default printer preference settings, and file associations.
The script will allow you to select your printer, as well as set your default print settings for the batch, and revert the default printer change after it is done, but not the preferences :(. When finished, it will ask if you want to move the files to an archive folder created within the folder directory you selected the files from. A log will be created in the same directory. 

The script supports multiple file types, and many more can be added in the future, since the print jobs rely on good file associations: 
```*.pdf;*.doc;*.docx;*.xls;*.xlsx;*.ppt;*.pptx;*.csv;*.txt```

Known programs that work with the file associations above: ```Notepad, Word, Excel, Adobe Acrobat Reader DC```

# Known Errors
Using browsers such as chrome for PDF file association will result in the following Error:

```"This command cannot be run due to the error: No application is associated with the specified file for this operation."```

# Deployment
For deployment, I used the [PS2EXE project](https://github.com/MScholtes/PS2EXE) to encapsulate the script into an executable that was easy for end users to launch. Some AVs may have some false positives during this process, and the program may need to be whitelisted. For this script, I recommend using the ```-noConsole``` and ```-noOutput``` Parameters.
