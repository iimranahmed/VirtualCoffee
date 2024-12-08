#Import the right modules
#import time
import numpy as np
import pandas as pd


# === Start this Thing === #

#Set the initial setting to roll again "0"
doagain=0

# Read CSV file
data = pd.read_csv('virtual_coffee_test_data.csv')

# Make a function
def rollfunc():
    rando=np.random.choice(len(data), 2, replace=False)
    print ""
    print  "Lucky Person 1 is: " + data['Name'][rando[0]]
    print  "Lucky Person 2 is: " + data['Name'][rando[1]]
    return rando

# Run this thing
while doagain!=1:
    index=rollfunc()
    doagain=input('Send email (1) or roll again (0):' )

print  "Need to make function to send email to",  data['Email'][index[0]], 'and',  data['Email'][index[1]]
