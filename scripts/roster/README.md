# Searider Roster Script
    Requirements:
        Python 3
        openpyxl (pip install -r requirements.txt)


<h3>SRoster.py</h3>
This program utilizes infinate campus extract to create usernames and passwords of students. Also sorts students by teacher/terms and outputs a papercutID.txt to use as a import for papercut ID numbers (student IDs)

    Requirements:

    Infinite Campus Export Field Order:
        Student Number  Last Name	First Name	Birthdate	Term Start	Term End
        Teacher Full Name	Period Start	Grade	Race Ethnicity	Gender	
        Room Name	Start Date	End Date

    Example:

    python sroster.py extractfile.csv outputfile.xlsx

        Optional flags:
        
                -p : Create PaperCut import file


<h3>PaperCut Import</h3>

To update papercut server to add in new users after Google user syncs. Run the sroster above to generate username and IDs for the students, a papercutID file shall be produced.

    1. From the directory of the export file:
        scp <export file> papercut@<server IP>:/home/papercut/server/bin/linux-x64/paperimport/<export file>
        
        This would move the file from your local computer to the remote papercut server
      
    2. SSH into papercut server.
        ssh papercut@<server ip> -l papercut
        
    3. Browse to bin directory:
        cd /server/bin/linux-x64/
      
    4. Run import script
        ./server-command batch-import-user-card-id-numbers /paperimport/<export file>

