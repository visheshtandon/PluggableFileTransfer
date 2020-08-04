# ***Pluggable File Transfer Utility***

    > This utility is a pluggable solution for transferring file(s) & directories across a source and destination.
    
    > This will help in cutting down development effort & time for custom data ingestion across multiple engagement allowing easy integration & minimal code changes.

&nbsp;

#  ***Features of Version 1.1*** 

|   Source   |   Destination   |
|    :-:     |       :-:       |
| Local System / Network Shared Folder     | Local System / Network Shared Folder     |
| Outlook     | Local System / Network Shared Folder     |
| Dropbox     | Local System / Network Shared Folder     |


&nbsp;

    * Transfer of file(s) & directory between Local System/Network Shared Folder.

    * Transfer of attachments from Outlook to Local System/Network Shared Folder.
    
    * Transfer of file(s) & directory from Dropbox to Local System/Network Shared Folder.

&nbsp;

#  ***User Guide*** 

To run the python script:

Type the following on command prompt -

`python FileTransfer.py`

&nbsp;

There are multiple modes of running the utility:

* ### ***"local"*** :

    > This mode will activate  transfer of file(s) & directory between Local System or Shared Network Folder.

    Type the following on command prompt:
    
    `python FileTransfer.py --local`

    &nbsp;

    Only on passing arguments ***"--local"*** , the utility will enable file(s)/directory transfer based on a configuration.

    > Configuration for local system transfer (Depending on the requirement):

        {
            "movement_type":
            {
                "local machine":
                [
                    {
                        "overwrite": "---",
                        "src_dir": "---",
                        "dest_dir": "---",
                        "file_name": "---"
                    },
                    {
                        "overwrite": "---",
                        "src_dir": "---",
                        "dest_dir": "---",
                        "file_name": "---"
                    },
                    ...
                    ...
                    ...
                ]
            }
        }
    
    
    > Mandatory Attributes to be passed to Configuration JSON:

    1. ***"overwrite"*** (String): Overwrite flag to assert whether the file(s)/directory should be overwritten or not. 

        &nbsp;
        ***List of Values: ["yes", "no"]***

    2. ***"src_dir"*** (String): Specifying the absolute source directory.

    3. ***"dest_dir"*** (String): Specifying the absolute destination directory.

    4. ***"file_name"*** (String): Specifying the file name(s) to be transferred.

        4.1. **Directory** (String): Mentioning file_name as **""** -> Specifying the directory to be transferred instead of file(s).        

        4.2. **File (s)** (String): Mentioning the keyword or the file name as it is (irrespective of the Letter Case).

    &nbsp;

* ### ***"outlook"*** :

    > This mode will activate  transfer of file(s) from Outlook to Local System or Shared Network Folder.

    Type the following on command prompt:
    
    `python FileTransfer.py --outlook`

    &nbsp;

    Only on passing arguments ***"--outlook"*** , the utility will enable file(s) transfer based on a configuration.

    > Configuration for Outlook transfer (Depending on the requirement):

        {
            "movement_type":
            {
                "outlook":
                [
                    {
                        "overwrite": "---",
                        "dest_dir": "---",
                        "emailSettings":
                        {
                            "subjects": ["---", "---"],
                            "email_saved_folder": "---",
                            "user_email": "---"
                        }
                    },
                    {
                        "overwrite": "---",
                        "dest_dir": "---",
                        "emailSettings":
                        {
                            "subjects": [],
                            "email_saved_folder": "---",
                            "user_email": "---"
                        }
                    }
                    ...
                    ...
                    ...
                ]
            }
        }
    
    
    > Mandatory Attributes to be passed to Configuration JSON:

    1. ***"overwrite"*** (String): Overwrite flag to assert whether the file(s)/directory should be overwritten or not. 

        &nbsp;
        ***List of Values: ["yes", "no"]***

    2. ***"dest_dir"*** (String): Specifying the absolute destination directory.
        String DataType.

    3. ***"emailSettings"*** : Outlook email settings required for connecting to the User's Outlook Account.
        
        3.1. ***"subjects"*** (List): Specifying the list of subjects to filter the emails. If list is empty, all the emails in the specified **"email_saved_folder"** will be retrieved.
    
        3.2. ***"email_saved_folder"*** (String): Specifying the saved folder to search the emails for.

        3.3. ***"user_email"*** (String): Specifying the user's email to transfer file(s) from.

&nbsp;

* ### ***"dropbox"*** :

    > This mode will activate  transfer of file(s)/Directory from Dropbox to Local System or Shared Network Folder.

    Type the following on command prompt:
    
    `python FileTransfer.py --dropbox`

    &nbsp;
    
    Only on passing arguments ***"--dropbox"*** , the utility will enable file(s)/Dropbox transfer based on a configuration.

    > Configuration for Dropbox transfer (Depending on the requirement):

        {
            "movement_type":
            {
                "dropbox":
                [
                    {
                        "overwrite": "---",
                        "src_dir": "---",
                        "dest_dir": "---",
                        "file_name": "",
                        "dropboxSettings":
                        {
                            "app_key":"---",
                            "app_secret":"---",
                            "access_token":"---"
                        }
                    },
                    {
                        "overwrite": "---",
                        "src_dir": "---",
                        "dest_dir": "---",
                        "file_name": "---",
                        "dropboxSettings":
                        {
                            "app_key":"---",
                            "app_secret":"---",
                            "access_token":"---"
                        }
                    },
                    ...
                    ...
                    ...
                ]
            }
        }
    
    
    > Mandatory Attributes to be passed to Configuration JSON:

    1. ***"overwrite"*** (String): Overwrite flag to assert whether the file(s)/directory should be overwritten or not.

        &nbsp;
        ***List of Values: ["yes", "no"]***

    2. ***"src_dir"*** (String): Specifying the absolute source directory.

    3. ***"dest_dir"*** (String): Specifying the absolute destination directory.

    4. ***"file_name"*** (String): Specifying the file name(s) to be transferred.
        
        4.1. **Directory** (String): Mentioning file_name as **""** -> Specifying the directory to be transferred instead of file(s.)

        4.2. **File(s)** (String): Mentioning the keywords or the file name as it is (irrespective of the Letter Case).

    5. ***"dropboxSettings"*** : Dropbox settings required for connecting to the User's Dropbox Account.

        5.1. ***"app_key"*** (String): Specifying the application key required to verify account credentials of the user.
    
        5.2. ***"app_secret"*** (String): Specifying the application secret required to verify account credentials of the user.
            String DataType.
    
        5.3. ***"access_token"*** (String): Specifying the access token required to authenticate account credentials of the user.

&nbsp;

# ***Scenarios***

* > By Default, all the file(s)/ directory transfer modes get activated unless otherwise specified.
    
        Type the following on command prompt:

    * `python FileTransfer.py`
        
        * Activating all the transfer modes (**local**, **outlook**, **dropbox**)

&nbsp;

* > Activating more than 1 mode for file(s) / directory transfer.
    
        Type the following on command prompt:
    
    * `python FileTransfer.py --local --dropbox`
        
        * Activating **local** and **dropbox** transfer modes.

    * `python FileTransfer.py --outlook --dropbox`

        * Activating **outlook** and **dropbox** transfer modes.

    * `python FileTransfer.py --local --outlook`

        * Activating **local** and **outlook** transfer modes.