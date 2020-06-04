# OREGON STATE CAPSTONE TEAM 612 FOR TEKFAB
## Data analyitics tool

### USAGE
#### getStates.exe
The getStates executable will create a randomized testing order for the experiment given.
First, create a 'states.txt' file in the CreateTestingOrder folder by copy/pasting a design specification found here:

[Engineering Statistics Handbook](https://www.itl.nist.gov/div898/handbook/pri/section3/pri3347.htm)

Please ensure that the text is left justified (no whitespace on the left side)
Here is an example of a correctly formated 'states.txt' file:

![Imgur](https://i.imgur.com/fPMtC7T.png)

Then the user must create a key, defining which level and factor corresponds to each number. The key must be in csv format. This key must be placed in the CreateTestingOrder folder as well. An example of a correctly formatted key.csv file is here:

![Imgur](https://i.imgur.com/Shtwl6L.png)

Once the states.txt file and the key.csv file are in the CreateTestingOrder folder, you may run the getStates.exe file.

After it runs, it will output an excel file named TestingOrder.xlsx and a key.json file. The TestingOrder.xlsx file will be where the user can input the results from the experiment. An example of the TestingOrder.xlsx given the above inputs can be seen here:

![Imgur](https://i.imgur.com/e1wtB70.png)

Now the user may fill in data anywhere to the right of the testing order. The user must supply a column heading, and each row of data must line up with the test that yielded that result.

A correctly filled TestingOrder.xlsx may look something like this (note that the data is top justified, to the right of the testingorder, and has column headers):

![Imgur](https://i.imgur.com/ChbVFBm.png)

Once data acquisition has completed, move the key.json and the filled TestingOrder.xlsx file into the CreateOutput folder.

#### CreateOutput.exe
In order to run the createOuput.exe, the user must have a filled TestingOrder.xlsx file, and the generated key.json file.
Running this exe will generate a results.txt file, along with accompanying graphs that show while the program is running.
This program will prompt the user to supply how many factors were tested. In this example case, there were 6 factors. Then the program will prompt the user to supply the column where the data is located in the TestingOrder.xlsx. In this example case, we are looking at dwell time, which is in column O.