# config_formatter

## **PURPOSE** 
Rockwell Automation's Configuration Tools utility used for exporting configuration summary files 
from .ACD files currently has the following issues:

    - missing parameter descriptions for general instructions (e.g. PVLV, PDI, PAO, etc.) that are 
      helpful when reviewing.

    - cells containing parameter descriptions are not formatted making navigating the config file 
      difficult to navigate (i.e. no text wrapping on long parameter descriptions resulting in 
      columns with widths that span an entire laptop monitor -> hard to see/scroll through information).

## **USAGE** 
Follow steps below to use this script to format a configuration file:

    1. copy the config file to the project directory's 'stage' folder.
    
    2. open a terminal and run the config_formatter.py script.

