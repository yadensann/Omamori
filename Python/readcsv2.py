import os 
import csv

userhome = os.path.expanduser('~')
csvpath = userhome + '/Desktop/netflixratings.csv'

with open(csvpath,'rU') as f:
    csvreader = csv.reader(f, delimiter = ',')
    for row in csvreader:
        print(row)

