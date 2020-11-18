# Sheets Comparer
App to compare the rota and availability email excel sheets to find which engineers are not present in them and automatically update the availability of engineers in the template.

# Quickstart
The app receives 2 parameters:

- The filename of the Rota excel
- The filename of the Availability excel

With these 2 files, the app compares the engineers in the rota to the list of engineers in the availability email to find if there are engineers missing in either one of those files.

Then the app updates the availability template excel with the information it gets from the Rota and saves a new excel file with the updated engineers availability in a new file with the suffix "-UPDATED".

The app also warns the user when the managers have put engineers in RED in the rota, meaning that those engineers have left the team. This means that the PP should remove these people from the availability template.

# Runnning
In order to run this app, simply clone this repository or just download the compare.py app to your computer.

Next, install the openpyxy library with pip:
```
# sudo pip install openpyxl
```

Next, make sure to download the latest versions of the rota and the availability email template and place them in the same folder of the compare.py app.

Then, simply open a terminal in that folder and type:
```
# python compare.py
```

In case you have saved the files in a different folder or with different names, you can also specify them like this:
```
# python compare.py <rota.xlsx> <availability.xlsx>
```

Please make sure to specify the files in the argument line in the order described above.
After running, the app will print an output similar to this and save the new availability template excel in a file with the suffix "-UPDATED" in the current folder:
```
$ python compare.py Deployment\ 2020\ Rota.xlsx Availability_email_template_2020.xlsx

Number of engineers in the Rota: 60
Number of engineers in Availability template: 57
Engineers available today: 44

In the Availability template but missing in rota:
afanh
trishans
tavarodr
spreetes
sshurya
yyoriki
ommechla

In the rota but missing in the availability template:
sabaring
ykuntsya
rmoskzb
lehant
mtiarnan
corcall
joannduf
malrasik
cchostak
bddrysda

Warning: these people have left the team!
Please remove them from the Availability Template!
khatd
gouvitor
aareshar
sampaig
pegaslee
sperlt
```

Happy PP'ing! :D