# Searider Roster Script


<h3>SRoster.py</h3>
This program utilizes infinate campus extract to create usernames and passwords of students. Also sorts students by teacher/terms and outputs a papercutID.txt to use as a import for papercut ID numbers (student IDs)

    Example:

    python sroster.py extractfile.csv outputfile.xlsx

        Optional flags:
        
                -p : PaperCut import file


<h3>PaperCut Import</h3>

To update papercut server to add in new users after Google user syncs. Run the sroster above to generate username and IDs for the students, a papercutID file shall be produced.

    1. Take the <papercut import file> and drop is into /Application/PaperCut MF/server/bin/mac/imports folder on the papercut server.

    2. On papercut server, browse via terminal to /Applications/PaperCut MF/server/bin/mac

    3. Run command sudo ./server-command batch-import-user-card-id-numbers import/<papercut import file>

